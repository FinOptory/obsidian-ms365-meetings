import {
	App,
	Notice,
	Plugin,
	PluginSettingTab,
	Setting,
	TFile,
	moment,
} from "obsidian";
import * as http from "http";
import * as https from "https";
import * as net from "net";

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

interface CalendarInfo {
	id: string;
	name: string;
	isDefault: boolean;
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
	selectedCalendarIds: string[]; // leer = alle Kalender
}

// ─── Constants ────────────────────────────────────────────────────────────────

const DEFAULT_SETTINGS: PluginSettings = {
	clientId: "aa4ae0d0-02b3-4a4a-8859-f99395bb753d",
	tenantId: "organizations",
	dailyNoteFolder: "01_daily",
	dailyNoteDateFormat: "YYYY-MM-DD",
	pollIntervalMinutes: 15,
	sectionHeading: "Meetings",
	skipAllDay: true,
	skipPrivate: true,
	selectedCalendarIds: [],
};

// Redirect URI: lokaler HTTP-Server, Port wird dynamisch gewählt
// Microsoft erlaubt http://localhost ohne spezifischen Port in Public Client Apps
function buildRedirectUri(port: number): string {
	return `http://localhost:${port}`;
}
const GRAPH_BASE    = "https://graph.microsoft.com/v1.0";
const MEETINGS_H1 = (heading: string) => `# ${heading}`;

// ─── Auth ────────────────────────────────────────────────────────────────────

class MSAuthClient {
	private cache: TokenCache | null = null;

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

	// Node.js HTTPS statt window.fetch — vermeidet Origin: app://obsidian.md Header
	private nodePost(url: string, body: string): Promise<Record<string, unknown>> {
		return new Promise((resolve, reject) => {
			const urlObj = new URL(url);
			const bodyBuf = Buffer.from(body, "utf-8");
			const req = https.request({
				hostname: urlObj.hostname,
				path: urlObj.pathname + urlObj.search,
				method: "POST",
				headers: {
					"Content-Type": "application/x-www-form-urlencoded",
					"Content-Length": bodyBuf.byteLength,
				},
			}, (res) => {
				let data = "";
				res.on("data", (chunk: Buffer) => { data += chunk.toString(); });
				res.on("end", () => {
					try { resolve(JSON.parse(data)); }
					catch { reject(new Error(`Invalid JSON: ${data.substring(0, 200)}`)); }
				});
			});
			req.on("error", reject);
			req.write(bodyBuf);
			req.end();
		});
	}

	async getAccessToken(): Promise<string> {
		if (this.cache && Date.now() < this.cache.expiresAt - 300_000) {
			return this.cache.accessToken;
		}
		if (this.cache?.refreshToken) {
			try { return await this.refreshAccessToken(this.cache.refreshToken); }
			catch { this.cache = null; }
		}
		// Kein automatischer Auth-Flow — nur explizit via Settings "Anmelden"
		throw new Error("Not authenticated");
	}

	isAuthenticated(): boolean {
		return !!(this.cache?.accessToken && Date.now() < this.cache.expiresAt - 300_000)
			|| !!(this.cache?.refreshToken);
	}

	private async refreshAccessToken(refreshToken: string): Promise<string> {
		const body = new URLSearchParams({
			client_id: this.clientId,
			grant_type: "refresh_token",
			refresh_token: refreshToken,
			scope: "Calendars.Read offline_access",
		});
		const data = await this.nodePost(this.tokenUrl(), body.toString()) as { access_token: string; expires_in: number; refresh_token?: string; error?: string };
		if (data.error) throw new Error(`Refresh failed: ${data.error}`);
		this.saveToken(data);
		return data.access_token;
	}

	private saveToken(data: { access_token: string; expires_in: number; refresh_token?: string }) {
		this.cache = {
			accessToken: data.access_token,
			expiresAt: Date.now() + data.expires_in * 1000,
			refreshToken: data.refresh_token,
		};
		this.plugin.savePluginData({ tokenCache: this.cache });
	}

	async startAuthFlow(): Promise<string> {
		const { verifier, challenge } = await this.generatePKCE();

		// Starte lokalen HTTP-Server auf einem freien Port
		const { port, waitForCode, closeServer } = await this.startLocalServer();
		const redirectUri = buildRedirectUri(port);

		const params = new URLSearchParams({
			client_id: this.clientId,
			response_type: "code",
			redirect_uri: redirectUri,
			scope: "Calendars.Read offline_access",
			code_challenge: challenge,
			code_challenge_method: "S256",
			response_mode: "query",
		});

		const authUrl = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/authorize?${params}`;
		window.open(authUrl);
		new Notice("Browser geöffnet — bitte mit Microsoft anmelden. Das Fenster schließt sich automatisch.", 10000);

		try {
			const code = await waitForCode();
			return await this.exchangeCode(code, redirectUri, verifier);
		} finally {
			closeServer();
		}
	}

	private startLocalServer(): Promise<{
		port: number;
		waitForCode: () => Promise<string>;
		closeServer: () => void;
	}> {
		return new Promise((resolve) => {
			let resolveCode: (code: string) => void;
			let rejectCode: (e: Error) => void;
			const codePromise = new Promise<string>((res, rej) => {
				resolveCode = res;
				rejectCode = rej;
			});

			const server = http.createServer((req, res) => {
				const url = new URL(req.url ?? "/", `http://localhost`);
				const code = url.searchParams.get("code");
				const error = url.searchParams.get("error");

				if (code) {
					res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
					res.end("<html><body style='font-family:sans-serif;padding:2em'><h2>✅ Anmeldung erfolgreich</h2><p>Du kannst dieses Fenster schließen und zu Obsidian zurückkehren.</p></body></html>");
					resolveCode(code);
				} else if (error) {
					res.writeHead(400, { "Content-Type": "text/html; charset=utf-8" });
					res.end(`<html><body><h2>Fehler: ${error}</h2></body></html>`);
					rejectCode(new Error(`Auth error: ${error}`));
				} else {
					res.writeHead(404);
					res.end();
				}
			});

			// Timeout: 5 Minuten
			const timeout = setTimeout(() => {
				rejectCode(new Error("Auth timeout nach 5 Minuten"));
				server.close();
			}, 300_000);

			server.listen(0, "localhost", () => {
				const addr = server.address() as net.AddressInfo;
				resolve({
					port: addr.port,
					waitForCode: () => codePromise.finally(() => clearTimeout(timeout)),
					closeServer: () => server.close(),
				});
			});
		});
	}

	private async exchangeCode(code: string, redirectUri: string, verifier: string): Promise<string> {
		const body = new URLSearchParams({
			client_id: this.clientId,
			grant_type: "authorization_code",
			code,
			redirect_uri: redirectUri,
			code_verifier: verifier,
			scope: "Calendars.Read offline_access",
		});
		const data = await this.nodePost(this.tokenUrl(), body.toString()) as { access_token: string; expires_in: number; refresh_token?: string; error?: string; error_description?: string };
		if (data.error) throw new Error(`Token exchange failed: ${data.error_description || data.error}`);
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

	async getCalendars(): Promise<CalendarInfo[]> {
		const token = await this.auth.getAccessToken();
		const res = await fetch(
			`${GRAPH_BASE}/me/calendars?$select=id,name,isDefaultCalendar&$top=50`,
			{ headers: { Authorization: `Bearer ${token}` } }
		);
		if (!res.ok) throw new Error(`Graph calendars error ${res.status}`);
		const data = await res.json() as { value: { id: string; name: string; isDefaultCalendar: boolean }[] };
		return data.value.map(c => ({ id: c.id, name: c.name, isDefault: c.isDefaultCalendar }));
	}

	async getEventsForDate(dateStr: string, selectedCalendarIds: string[] = []): Promise<MeetingEvent[]> {
		// Wenn bestimmte Kalender ausgewählt: pro Kalender einzeln abfragen und zusammenführen
		if (selectedCalendarIds.length > 0) {
			const results = await Promise.all(
				selectedCalendarIds.map(id => this.getEventsForCalendar(id, dateStr))
			);
			const all = results.flat();
			all.sort((a, b) => a.start.localeCompare(b.start));
			return all;
		}
		// Sonst: alle Kalender via /me/calendarView
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

	private async getEventsForCalendar(calendarId: string, dateStr: string): Promise<MeetingEvent[]> {
		const token = await this.auth.getAccessToken();
		const viennaOffset = this.getViennaOffset(new Date(`${dateStr}T12:00:00`));
		const startUtc = new Date(`${dateStr}T00:00:00${viennaOffset}`).toISOString();
		const endUtc   = new Date(`${dateStr}T23:59:59${viennaOffset}`).toISOString();

		const params = new URLSearchParams({
			startDateTime: startUtc,
			endDateTime:   endUtc,
			$select: "id,subject,start,end,isAllDay,isCancelled,attendees,onlineMeeting,location,sensitivity",
			$top: "50",
		});

		const res = await fetch(
			`${GRAPH_BASE}/me/calendars/${calendarId}/calendarView?${params}`,
			{
				headers: {
					Authorization: `Bearer ${token}`,
					Prefer: `outlook.timezone="Europe/Vienna"`,
				},
			}
		);
		if (!res.ok) return [];
		const data = await res.json() as { value: GraphEvent[] };
		return data.value.map(this.mapEvent.bind(this));
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
// Kein Anker-Kommentar. Notizen-Persistenz via Zeit-Matching (HH:MM – HH:MM).
// Der Meetings-Block wird durch den H1 "# Meetings" begrenzt.

class DailyNoteWriter {
	constructor(private settings: PluginSettings) {}

	private timeKey(event: MeetingEvent): string {
		return `${event.start.substring(11, 16)}–${event.end.substring(11, 16)}`;
	}

	// Extrahiert Notizen aus bestehenden Meeting-Blöcken anhand der Zeit.
	private extractNotesByTime(content: string): Map<string, string> {
		const notes = new Map<string, string>();
		// Suche: ## Irgendwas | HH:MM – HH:MM ... _Notizen:_ \n <inhalt> \n bis zum nächsten ## oder EOF
		const blockRe = /^## .+?\| (\d{2}:\d{2}) – (\d{2}:\d{2})[\s\S]*?_Notizen:_\n([\s\S]*?)(?=^## |\Z)/gm;
		let match;
		while ((match = blockRe.exec(content)) !== null) {
			const key = `${match[1]}–${match[2]}`;
			const noteText = match[3].replace(/\n+$/, "").trim();
			if (noteText) notes.set(key, noteText);
		}
		return notes;
	}

	private buildBlock(event: MeetingEvent, existingNotes: Map<string, string>): string {
		const start = event.start.substring(11, 16);
		const end   = event.end.substring(11, 16);

		const header = event.isCancelled
			? `## ~~${event.subject}~~ [ABGESAGT] | ${start} – ${end}`
			: `## ${event.subject} | ${start} – ${end}`;

		const lines: string[] = [header];

		if (event.attendees.length > 0) {
			const shown = event.attendees.slice(0, 4);
			const rest  = event.attendees.length - 4;
			lines.push(`Teilnehmer: ${shown.join(", ")}${rest > 0 ? `, +${rest}` : ""}`);
		}

		if (event.joinUrl) {
			lines.push(`[Meeting beitreten](${event.joinUrl})`);
		} else if (event.location) {
			lines.push(`Ort: ${event.location}`);
		}

		lines.push("");
		lines.push("_Notizen:_");

		const savedNotes = existingNotes.get(this.timeKey(event));
		if (savedNotes) {
			lines.push(savedNotes);
		} else {
			lines.push("");
		}

		return lines.join("\n");
	}

	private totalDuration(events: MeetingEvent[]): string {
		let totalMin = 0;
		for (const e of events) {
			if (e.isAllDay) continue;
			const [sh, sm] = e.start.substring(11, 16).split(":").map(Number);
			const [eh, em] = e.end.substring(11, 16).split(":").map(Number);
			const mins = (eh * 60 + em) - (sh * 60 + sm);
			if (mins > 0) totalMin += mins;
		}
		const h = Math.floor(totalMin / 60);
		const m = totalMin % 60;
		return `${h}:${String(m).padStart(2, "0")}`;
	}

	mergeIntoNote(existingContent: string, events: MeetingEvent[], settings: PluginSettings): string {
		// Bereinige alte Anker-Kommentare aus früheren Versionen
		let content = existingContent
			.replace(/<!-- ms365-meetings-section -->[\s\S]*?<!-- ms365-meetings-section-end -->\n?/g, "")
			.replace(/%% ms365-meetings-section %%[\s\S]*?%% ms365-meetings-section-end %%\n?/g, "");

		const existingNotes = this.extractNotesByTime(content);

		const visible = events
			.filter(e => !(e.isAllDay && settings.skipAllDay))
			.sort((a, b) => a.start.localeCompare(b.start));

		if (visible.length === 0) return content;

		const duration = this.totalDuration(visible);
		const h1 = `# ${settings.sectionHeading} (${duration})`;
		const blocks = visible.map(e => this.buildBlock(e, existingNotes)).join("\n\n");
		const section = `${h1}\n\n${blocks}\n`;

		// Ersetze bestehenden H1-Block oder hänge an
		// Matcht "# Meetings" und optional "(6:30)" dahinter
		const h1Re = new RegExp(`^# ${settings.sectionHeading}[^\\n]*\\n[\\s\\S]*?(?=^# |\\Z)`, "m");
		if (h1Re.test(content)) {
			return content.replace(h1Re, section);
		}
		return `${content.trimEnd()}\n\n${section}`;
	}
}

// ─── Plugin ───────────────────────────────────────────────────────────────────

interface PluginData {
	settings: PluginSettings;
	tokenCache: TokenCache | null;
}

export default class MS365MeetingsPlugin extends Plugin {
	settings!: PluginSettings;
	auth!: MSAuthClient;
	graph!: GraphCalendarClient;
	private writer!: DailyNoteWriter;
	private pollTimer: number | null = null;
	private pluginData: PluginData = { settings: DEFAULT_SETTINGS, tokenCache: null };

	async onload() {
		await this.loadSettings();

		this.auth = new MSAuthClient(this.settings.clientId, this.settings.tenantId, this);
		this.auth.restoreCache(this.pluginData.tokenCache);
		this.graph = new GraphCalendarClient(this.auth);
		this.writer = new DailyNoteWriter(this.settings);

		// Update wenn Daily Note geöffnet wird — nur wenn eingeloggt
		this.registerEvent(
			this.app.workspace.on("file-open", async (file) => {
				if (file && this.isDailyNote(file) && this.auth.isAuthenticated()) {
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
			const events = await this.graph.getEventsForDate(date, this.settings.selectedCalendarIds);
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
			if (!this.auth.isAuthenticated()) return;
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
		// tenantId + clientId immer aus DEFAULT_SETTINGS — nie aus altem Cache
		this.settings.tenantId = DEFAULT_SETTINGS.tenantId;
		this.settings.clientId = DEFAULT_SETTINGS.clientId;
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
	private calendars: CalendarInfo[] = [];

	constructor(app: App, private plugin: MS365MeetingsPlugin) {
		super(app, plugin);
	}

	display(): void {
		const { containerEl } = this;
		containerEl.empty();
		containerEl.createEl("h2", { text: "MS365 Meetings" });

		// ── Anmeldung ──────────────────────────────────────────────
		containerEl.createEl("h3", { text: "Anmeldung" });

		const isLoggedIn = this.plugin.auth.isAuthenticated();
		new Setting(containerEl)
			.setName("Microsoft 365")
			.setDesc(isLoggedIn ? "✅ Angemeldet" : "Noch nicht angemeldet — Token wird nach Anmeldung gespeichert.")
			.addButton(b => {
				b.setButtonText(isLoggedIn ? "Neu anmelden" : "Anmelden");
				if (!isLoggedIn) b.setCta();
				b.onClick(async () => {
					try {
						await this.plugin.auth.startAuthFlow();
						new Notice("✅ Anmeldung erfolgreich");
						this.calendars = []; // Reset damit auto-load greift
						this.display();
					} catch (e) {
						new Notice(`Fehler: ${e}`);
					}
				});
			});

		// ── Kalender ───────────────────────────────────────────────
		containerEl.createEl("h3", { text: "Kalender" });

		const calSection = containerEl.createDiv();
		this.renderCalendarSection(calSection);

		// Kalender automatisch laden wenn eingeloggt und noch nicht geladen
		if (this.plugin.auth.isAuthenticated() && this.calendars.length === 0) {
			this.loadCalendars(calSection);
		}

		new Setting(containerEl)
			.setName("Kalender aktualisieren")
			.setDesc("Kalenderliste neu laden")
			.addButton(b => b
				.setButtonText("Laden")
				.onClick(() => this.loadCalendars(calSection)));

		// ── Termine ────────────────────────────────────────────────
		containerEl.createEl("h3", { text: "Termine" });

		new Setting(containerEl)
			.setName("Ganztägige Termine anzeigen")
			.setDesc("Ganztägige Events (z.B. Urlaub) in der Daily Note ein- oder ausblenden")
			.addToggle(t => t
				.setValue(!this.plugin.settings.skipAllDay)
				.onChange(async v => {
					this.plugin.settings.skipAllDay = !v;
					await this.plugin.saveSettings();
				}));

		// ── Daily Note ─────────────────────────────────────────────
		containerEl.createEl("h3", { text: "Daily Note" });

		new Setting(containerEl)
			.setName("Ordner")
			.setDesc("Relativer Pfad im Vault, z.B. 01_daily")
			.addText(t => t
				.setValue(this.plugin.settings.dailyNoteFolder)
				.onChange(async v => { this.plugin.settings.dailyNoteFolder = v; await this.plugin.saveSettings(); }));

		new Setting(containerEl)
			.setName("Datumsformat")
			.setDesc("Format der Dateinamen, z.B. YYYY-MM-DD")
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
			.setDesc("Wie oft Termine automatisch aktualisiert werden")
			.addSlider(s => s
				.setLimits(5, 60, 5)
				.setValue(this.plugin.settings.pollIntervalMinutes)
				.setDynamicTooltip()
				.onChange(async v => { this.plugin.settings.pollIntervalMinutes = v; await this.plugin.saveSettings(); }));
	}

	private renderCalendarSection(container: HTMLElement) {
		container.empty();
		if (this.calendars.length === 0) {
			container.createEl("p", {
				text: "Noch keine Kalender geladen. Zuerst anmelden, dann 'Laden' klicken.",
				cls: "setting-item-description",
			});
			return;
		}

		const selected = new Set(this.plugin.settings.selectedCalendarIds);

		container.createEl("p", {
			text: selected.size === 0
				? "Alle Kalender werden synchronisiert."
				: `${selected.size} Kalender ausgewählt.`,
			cls: "setting-item-description",
		});

		for (const cal of this.calendars) {
			new Setting(container)
				.setName(cal.name + (cal.isDefault ? " ★" : ""))
				.addToggle(t => t
					.setValue(selected.size === 0 || selected.has(cal.id))
					.onChange(async (on) => {
						// Wenn alle an waren (leeres Array) und einer wird abgewählt:
						// alle anderen explizit aktivieren
						const cur = new Set(this.plugin.settings.selectedCalendarIds);
						if (cur.size === 0) {
							// Fülle mit allen außer diesem
							this.calendars.forEach(c => { if (c.id !== cal.id) cur.add(c.id); });
						} else if (on) {
							cur.add(cal.id);
						} else {
							cur.delete(cal.id);
						}
						// Wenn alle aktiv: zurück zu leerem Array (= alle)
						if (cur.size === this.calendars.length) cur.clear();
						this.plugin.settings.selectedCalendarIds = Array.from(cur);
						await this.plugin.saveSettings();
						this.renderCalendarSection(container);
					}));
		}
	}

	private async loadCalendars(container: HTMLElement) {
		try {
			this.calendars = await this.plugin.graph.getCalendars();
			this.renderCalendarSection(container);
		} catch (e) {
			new Notice(`Kalender laden fehlgeschlagen: ${e}`);
		}
	}
}
