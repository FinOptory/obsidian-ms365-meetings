import {
	App,
	Notice,
	Plugin,
	PluginSettingTab,
	Setting,
	TFile,
	moment,
} from "obsidian";

// ─── Types ────────────────────────────────────────────────────────────────────

interface MeetingEvent {
	id: string;
	subject: string;
	start: string;       // ISO datetime (Vienna)
	end: string;         // ISO datetime (Vienna)
	isCancelled: boolean;
	attendees: string[]; // display names or emails
	joinUrl?: string;    // Teams/Zoom link
	location?: string;
	isAllDay: boolean;
}

interface TokenCache {
	accessToken: string;
	expiresAt: number; // unix ms
	refreshToken?: string;
}

interface PluginSettings {
	clientId: string;
	tenantId: string;
	dailyNoteFolder: string;
	dailyNoteDateFormat: string;
	pollIntervalMinutes: number;
	sectionHeading: string;
	skipAllDay: boolean;
	skipPrivate: boolean;
}

// ─── Constants ────────────────────────────────────────────────────────────────

const DEFAULT_SETTINGS: PluginSettings = {
	clientId: "aa4ae0d0-02b3-4a4a-8859-f99395bb753d",
	tenantId: "common",
	dailyNoteFolder: "01_daily",
	dailyNoteDateFormat: "YYYY-MM-DD",
	pollIntervalMinutes: 15,
	sectionHeading: "Meetings",
	skipAllDay: true,
	skipPrivate: true,
};

const REDIRECT_URI  = "obsidian://ms365-meetings/callback";
const GRAPH_BASE    = "https://graph.microsoft.com/v1.0";
const BLOCK_START   = (id: string) => `<!-- ms365-meeting:${id} -->`;
const BLOCK_END_TAG = "<!-- ms365-meeting-end -->";
const SECTION_START = "<!-- ms365-meetings-section -->";
const SECTION_END   = "<!-- ms365-meetings-section-end -->";

// ─── Auth ────────────────────────────────────────────────────────────────────

class MSAuthClient {
	private cache: TokenCache | null = null;
	private codeVerifier = "";
	private pendingResolve: ((code: string) => void) | null = null;

	constructor(
		private clientId: string,
		private tenantId: string,
		private plugin: MS365MeetingsPlugin
	) {}

	private base64url(buf: ArrayBuffer): string {
		return btoa(String.fromCharCode(...new Uint8Array(buf)))
			.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
	}

	private async generatePKCE(): Promise<{ verifier: string; challenge: string }> {
		const array = new Uint8Array(32);
		crypto.getRandomValues(array);
		const verifier = this.base64url(array.buffer);
		const encoder = new TextEncoder();
		const digest = await crypto.subtle.digest("SHA-256", encoder.encode(verifier));
		const challenge = this.base64url(digest);
		return { verifier, challenge };
	}

	private tokenUrl(): string {
		return `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
	}

	async getAccessToken(): Promise<string> {
		// Return cached token if still valid (5 min buffer)
		if (this.cache && Date.now() < this.cache.expiresAt - 300_000) {
			return this.cache.accessToken;
		}

		// Try refresh
		if (this.cache?.refreshToken) {
			try {
				return await this.refreshAccessToken(this.cache.refreshToken);
			} catch {
				this.cache = null;
			}
		}

		// Full auth flow
		return await this.startAuthFlow();
	}

	private async refreshAccessToken(refreshToken: string): Promise<string> {
		const body = new URLSearchParams({
			client_id: this.clientId,
			grant_type: "refresh_token",
			refresh_token: refreshToken,
			scope: "Calendars.Read offline_access",
		});

		const res = await fetch(this.tokenUrl(), {
			method: "POST",
			headers: { "Content-Type": "application/x-www-form-urlencoded" },
			body: body.toString(),
		});

		if (!res.ok) throw new Error("Refresh failed");
		const data = await res.json() as { access_token: string; expires_in: number; refresh_token?: string };
		this.saveToken(data);
		return data.access_token;
	}

	private saveToken(data: { access_token: string; expires_in: number; refresh_token?: string }) {
		this.cache = {
			accessToken: data.access_token,
			expiresAt: Date.now() + data.expires_in * 1000,
			refreshToken: data.refresh_token,
		};
		// Persist encrypted in plugin data
		this.plugin.savePluginData({ tokenCache: this.cache });
	}

	async startAuthFlow(): Promise<string> {
		const { verifier, challenge } = await this.generatePKCE();
		this.codeVerifier = verifier;

		const params = new URLSearchParams({
			client_id: this.clientId,
			response_type: "code",
			redirect_uri: REDIRECT_URI,
			scope: "Calendars.Read offline_access",
			code_challenge: challenge,
			code_challenge_method: "S256",
			response_mode: "query",
		});

		const authUrl = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/authorize?${params}`;

		// Open browser
		window.open(authUrl);
		new Notice("Browser geöffnet — bitte mit Microsoft anmelden. Das Fenster schließt sich automatisch.", 10000);

		// Wait for redirect callback
		return new Promise((resolve, reject) => {
			this.pendingResolve = async (code: string) => {
				try {
					const token = await this.exchangeCode(code);
					resolve(token);
				} catch (e) {
					reject(e);
				}
			};
			// Timeout after 5 min
			setTimeout(() => reject(new Error("Auth timeout")), 300_000);
		});
	}

	async handleCallback(code: string): Promise<void> {
		if (this.pendingResolve) {
			this.pendingResolve(code);
			this.pendingResolve = null;
		}
	}

	private async exchangeCode(code: string): Promise<string> {
		const body = new URLSearchParams({
			client_id: this.clientId,
			grant_type: "authorization_code",
			code,
			redirect_uri: REDIRECT_URI,
			code_verifier: this.codeVerifier,
			scope: "Calendars.Read offline_access",
		});

		const res = await fetch(this.tokenUrl(), {
			method: "POST",
			headers: { "Content-Type": "application/x-www-form-urlencoded" },
			body: body.toString(),
		});

		if (!res.ok) {
			const err = await res.text();
			throw new Error(`Token exchange failed: ${err}`);
		}

		const data = await res.json() as { access_token: string; expires_in: number; refresh_token?: string };
		this.saveToken(data);
		return data.access_token;
	}

	restoreCache(cache: TokenCache | null) {
		this.cache = cache;
	}
}

// ─── Graph Client ─────────────────────────────────────────────────────────────

class GraphCalendarClient {
	constructor(private auth: MSAuthClient) {}

	async getEventsForDate(dateStr: string): Promise<MeetingEvent[]> {
		const token = await this.auth.getAccessToken();

		// Build Vienna-aware time window
		const start = `${dateStr}T00:00:00`;
		const end   = `${dateStr}T23:59:59`;

		// Convert to UTC for Graph (Vienna = UTC+1 in CET, UTC+2 in CEST)
		const viennaOffset = this.getViennaOffset(new Date(`${dateStr}T12:00:00`));
		const startUtc = new Date(`${start}${viennaOffset}`).toISOString();
		const endUtc   = new Date(`${end}${viennaOffset}`).toISOString();

		const params = new URLSearchParams({
			startDateTime: startUtc,
			endDateTime:   endUtc,
			$select: "id,subject,start,end,isAllDay,isCancelled,attendees,onlineMeeting,location,sensitivity,showAs",
			$orderby: "start/dateTime",
			$top: "50",
		});

		const res = await fetch(`${GRAPH_BASE}/me/calendarView?${params}`, {
			headers: {
				Authorization: `Bearer ${token}`,
				Prefer: `outlook.timezone="Europe/Vienna"`,
			},
		});

		if (!res.ok) {
			const err = await res.text();
			throw new Error(`Graph API error ${res.status}: ${err}`);
		}

		const data = await res.json() as { value: GraphEvent[] };
		return data.value
			.filter(e => {
				// Skip private events if configured
				if (e.sensitivity === "private") return false;
				return true;
			})
			.map(this.mapEvent.bind(this));
	}

	private getViennaOffset(date: Date): string {
		const parts = new Intl.DateTimeFormat("en", {
			timeZone: "Europe/Vienna",
			timeZoneName: "shortOffset",
		}).formatToParts(date);
		const offsetStr = parts.find(p => p.type === "timeZoneName")?.value ?? "GMT+1";
		// offsetStr is like "GMT+2" or "GMT+1"
		const sign = offsetStr.includes("-") ? "-" : "+";
		const hours = offsetStr.replace(/GMT[+-]?/, "") || "1";
		return `${sign}${hours.padStart(2, "0")}:00`;
	}

	private mapEvent(e: GraphEvent): MeetingEvent {
		const attendees = (e.attendees || [])
			.filter(a => a.type !== "resource")
			.map(a => a.emailAddress.name || a.emailAddress.address)
			.slice(0, 10); // max 10 für die Anzeige

		return {
			id: e.id,
			subject: e.subject || "(Kein Titel)",
			start: e.start.dateTime.substring(0, 16).replace("T", " "),
			end:   e.end.dateTime.substring(0, 16).replace("T", " "),
			isCancelled: e.isCancelled ?? false,
			attendees,
			joinUrl: e.onlineMeeting?.joinUrl,
			location: e.location?.displayName,
			isAllDay: e.isAllDay ?? false,
		};
	}
}

interface GraphEvent {
	id: string;
	subject: string;
	start: { dateTime: string; timeZone: string };
	end:   { dateTime: string; timeZone: string };
	isAllDay: boolean;
	isCancelled: boolean;
	sensitivity: string;
	showAs: string;
	attendees: { type: string; emailAddress: { name: string; address: string } }[];
	onlineMeeting?: { joinUrl: string };
	location?: { displayName: string };
}

// ─── Daily Note Writer ────────────────────────────────────────────────────────

class DailyNoteWriter {
	constructor(private settings: PluginSettings) {}

	buildEventBlock(event: MeetingEvent): string {
		const startTime = event.start.substring(11, 16);
		const endTime   = event.end.substring(11, 16);

		// H2: Abgesagt mit Strikethrough
		const titleLine = event.isCancelled
			? `## ~~${event.subject}~~ [ABGESAGT] | ${startTime} – ${endTime}`
			: `## ${event.subject} | ${startTime} – ${endTime}`;

		// Teilnehmer-Zeile
		const maxShow = 4;
		let attendeeLine = "";
		if (event.attendees.length > 0) {
			const shown = event.attendees.slice(0, maxShow);
			const rest  = event.attendees.length - maxShow;
			attendeeLine = `Teilnehmer: ${shown.join(", ")}${rest > 0 ? `, +${rest}` : ""}`;
		}

		// Link
		const linkLine = event.joinUrl
			? `[Meeting beitreten](${event.joinUrl})`
			: event.location
				? `Ort: ${event.location}`
				: "";

		const lines = [
			BLOCK_START(event.id),
			titleLine,
		];
		if (attendeeLine) lines.push(attendeeLine);
		if (linkLine) lines.push(linkLine);
		lines.push("");
		lines.push("_Notizen:_");
		lines.push("");
		lines.push(BLOCK_END_TAG);

		return lines.join("\n");
	}

	/**
	 * Merged neue Events in bestehende Daily Note.
	 * Bestehende Notizen bleiben erhalten.
	 */
	mergeIntoNote(existingContent: string, events: MeetingEvent[], settings: PluginSettings): string {
		// Baue Map: event-id → existierende Notizen
		const existingNotes = this.extractExistingNotes(existingContent);

		// Baue neue Meetings-Sektion
		const heading = `# ${settings.sectionHeading}`;
		const eventBlocks = events
			.filter(e => !(e.isAllDay && settings.skipAllDay))
			.sort((a, b) => a.start.localeCompare(b.start))
			.map(e => {
				const block = this.buildEventBlock(e);
				// Preserve existing notes
				const savedNotes = existingNotes.get(e.id);
				if (savedNotes) {
					return block.replace(
						"_Notizen:_\n\n" + BLOCK_END_TAG,
						`_Notizen:_\n${savedNotes}\n${BLOCK_END_TAG}`
					);
				}
				return block;
			});

		const section = [
			SECTION_START,
			heading,
			"",
			...eventBlocks.flatMap(b => [b, ""]),
			SECTION_END,
		].join("\n");

		// Ersetze bestehende Sektion oder füge am Ende an
		if (existingContent.includes(SECTION_START)) {
			return existingContent.replace(
				new RegExp(`${SECTION_START}[\\s\\S]*?${SECTION_END}`),
				section
			);
		}

		return `${existingContent.trimEnd()}\n\n${section}\n`;
	}

	private extractExistingNotes(content: string): Map<string, string> {
		const notes = new Map<string, string>();
		const blockRe = /<!-- ms365-meeting:([^>]+) -->([\s\S]*?)<!-- ms365-meeting-end -->/g;

		let match;
		while ((match = blockRe.exec(content)) !== null) {
			const id = match[1];
			const blockContent = match[2];
			// Extract content between "_Notizen:_" and the end tag
			const notesMatch = blockContent.match(/_Notizen:_\n([\s\S]*)/);
			if (notesMatch && notesMatch[1].trim()) {
				notes.set(id, notesMatch[1]);
			}
		}
		return notes;
	}
}

// ─── Plugin ───────────────────────────────────────────────────────────────────

interface PluginData {
	settings: PluginSettings;
	tokenCache: TokenCache | null;
}

export default class MS365MeetingsPlugin extends Plugin {
	settings!: PluginSettings;
	private auth!: MSAuthClient;
	private graph!: GraphCalendarClient;
	private writer!: DailyNoteWriter;
	private pollTimer: number | null = null;
	private pluginData: PluginData = { settings: DEFAULT_SETTINGS, tokenCache: null };

	async onload() {
		await this.loadSettings();

		this.auth = new MSAuthClient(this.settings.clientId, this.settings.tenantId, this);
		this.auth.restoreCache(this.pluginData.tokenCache);
		this.graph = new GraphCalendarClient(this.auth);
		this.writer = new DailyNoteWriter(this.settings);

		// Handle OAuth callback: obsidian://ms365-meetings/callback?code=...
		this.registerObsidianProtocolHandler("ms365-meetings", async (params) => {
			if (params.action === "callback" && params.code) {
				await this.auth.handleCallback(params.code);
				new Notice("✅ Microsoft-Anmeldung erfolgreich");
			}
		});

		// Update wenn Daily Note geöffnet wird
		this.registerEvent(
			this.app.workspace.on("file-open", async (file) => {
				if (file && this.isDailyNote(file)) {
					await this.updateDailyNote(file);
				}
			})
		);

		// Polling
		this.startPolling();

		// Settings Tab
		this.addSettingTab(new MS365MeetingsSettingTab(this.app, this));

		// Command: manuelles Update
		this.addCommand({
			id: "update-meetings",
			name: "Termine jetzt aktualisieren",
			callback: async () => {
				const file = this.app.workspace.getActiveFile();
				if (file && this.isDailyNote(file)) {
					await this.updateDailyNote(file);
					new Notice("Termine aktualisiert");
				} else {
					new Notice("Keine Daily Note aktiv");
				}
			},
		});

		// Command: Anmelden
		this.addCommand({
			id: "sign-in",
			name: "Mit Microsoft anmelden",
			callback: async () => {
				try {
					await this.auth.startAuthFlow();
					new Notice("✅ Anmeldung erfolgreich");
				} catch (e) {
					new Notice(`Anmeldung fehlgeschlagen: ${e}`);
				}
			},
		});
	}

	onunload() {
		if (this.pollTimer) window.clearInterval(this.pollTimer);
	}

	private isDailyNote(file: TFile): boolean {
		const folder = this.settings.dailyNoteFolder;
		const fmt    = this.settings.dailyNoteDateFormat;
		const dateStr = moment().format(fmt);
		// Check if file is in daily note folder and matches date format
		return file.path.startsWith(folder) && file.basename.match(/^\d{4}-\d{2}-\d{2}/) !== null;
	}

	private getTodayDateStr(): string {
		return moment().format(this.settings.dailyNoteDateFormat);
	}

	async updateDailyNote(file: TFile, dateStr?: string) {
		const date = dateStr ?? file.basename.substring(0, 10);
		try {
			const events = await this.graph.getEventsForDate(date);
			const content = await this.app.vault.read(file);
			const updated = this.writer.mergeIntoNote(content, events, this.settings);
			if (updated !== content) {
				await this.app.vault.modify(file, updated);
			}
		} catch (e) {
			console.error("[ms365-meetings] Update failed:", e);
			if (String(e).includes("Auth")) {
				new Notice("MS365 Meetings: Bitte neu anmelden (Cmd+P → Mit Microsoft anmelden)");
			}
		}
	}

	private startPolling() {
		if (this.pollTimer) window.clearInterval(this.pollTimer);
		const ms = this.settings.pollIntervalMinutes * 60 * 1000;
		this.pollTimer = window.setInterval(async () => {
			// Update heute's Daily Note wenn offen
			const active = this.app.workspace.getActiveFile();
			if (active && this.isDailyNote(active)) {
				await this.updateDailyNote(active);
			}
		}, ms);
	}

	async loadSettings() {
		const raw = await this.loadData() as PluginData | null;
		this.pluginData = Object.assign({ settings: DEFAULT_SETTINGS, tokenCache: null }, raw ?? {});
		this.settings = Object.assign({}, DEFAULT_SETTINGS, this.pluginData.settings);
	}

	async saveSettings() {
		this.pluginData.settings = this.settings;
		await this.saveData(this.pluginData);
		this.startPolling(); // Polling-Interval neu starten
	}

	async savePluginData(patch: Partial<PluginData>) {
		Object.assign(this.pluginData, patch);
		await this.saveData(this.pluginData);
	}
}

// ─── Settings Tab ─────────────────────────────────────────────────────────────

class MS365MeetingsSettingTab extends PluginSettingTab {
	constructor(app: App, private plugin: MS365MeetingsPlugin) {
		super(app, plugin);
	}

	display(): void {
		const { containerEl } = this;
		containerEl.empty();
		containerEl.createEl("h2", { text: "MS365 Meetings" });

		new Setting(containerEl)
			.setName("Daily Note Ordner")
			.setDesc("Relativer Pfad im Vault, z.B. 01_daily")
			.addText(t => t
				.setValue(this.plugin.settings.dailyNoteFolder)
				.onChange(async v => { this.plugin.settings.dailyNoteFolder = v; await this.plugin.saveSettings(); }));

		new Setting(containerEl)
			.setName("Datumsformat")
			.setDesc("Format der Daily Note Dateinamen, z.B. YYYY-MM-DD")
			.addText(t => t
				.setValue(this.plugin.settings.dailyNoteDateFormat)
				.onChange(async v => { this.plugin.settings.dailyNoteDateFormat = v; await this.plugin.saveSettings(); }));

		new Setting(containerEl)
			.setName("Abschnittstitel")
			.setDesc("H1-Überschrift der Meetings-Sektion")
			.addText(t => t
				.setValue(this.plugin.settings.sectionHeading)
				.onChange(async v => { this.plugin.settings.sectionHeading = v; await this.plugin.saveSettings(); }));

		new Setting(containerEl)
			.setName("Poll-Intervall (Minuten)")
			.setDesc("Wie oft Termine aktualisiert werden")
			.addSlider(s => s
				.setLimits(5, 60, 5)
				.setValue(this.plugin.settings.pollIntervalMinutes)
				.setDynamicTooltip()
				.onChange(async v => { this.plugin.settings.pollIntervalMinutes = v; await this.plugin.saveSettings(); }));

		new Setting(containerEl)
			.setName("Ganztägige Termine überspringen")
			.addToggle(t => t
				.setValue(this.plugin.settings.skipAllDay)
				.onChange(async v => { this.plugin.settings.skipAllDay = v; await this.plugin.saveSettings(); }));

		new Setting(containerEl)
			.setName("Anmeldung")
			.setDesc("Mit Microsoft 365 anmelden")
			.addButton(b => b
				.setButtonText("Anmelden")
				.onClick(async () => {
					try {
						await this.plugin.auth.startAuthFlow();
						new Notice("✅ Anmeldung erfolgreich");
					} catch (e) {
						new Notice(`Fehler: ${e}`);
					}
				}));
	}
}
