/**
 * Standalone Auth-Test — läuft ohne Obsidian.
 * Startet lokalen HTTP-Server, gibt Auth-URL aus, wartet auf Code,
 * tauscht Code gegen Token, ruft /me/calendars auf.
 *
 * Usage: npx tsx test-auth.ts
 */

import * as http from "http";
import * as https from "https";
import * as net from "net";
import * as crypto from "crypto";

const CLIENT_ID = "aa4ae0d0-02b3-4a4a-8859-f99395bb753d";
const TENANT_ID = "organizations";
const TOKEN_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const GRAPH_URL = "https://graph.microsoft.com/v1.0/me/calendars?$select=id,name,isDefaultCalendar&$top=10";

function base64url(buf: Buffer): string {
	return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function generatePKCE(): { verifier: string; challenge: string } {
	const verifier = base64url(crypto.randomBytes(32));
	const challenge = base64url(crypto.createHash("sha256").update(verifier).digest());
	return { verifier, challenge };
}

function nodePost(url: string, body: string): Promise<Record<string, unknown>> {
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
			res.on("data", (c: Buffer) => { data += c.toString(); });
			res.on("end", () => {
				try { resolve(JSON.parse(data)); }
				catch { reject(new Error(`Invalid JSON: ${data.substring(0, 300)}`)); }
			});
		});
		req.on("error", reject);
		req.write(bodyBuf);
		req.end();
	});
}

function nodeGet(url: string, token: string): Promise<Record<string, unknown>> {
	return new Promise((resolve, reject) => {
		const urlObj = new URL(url);
		const req = https.request({
			hostname: urlObj.hostname,
			path: urlObj.pathname + urlObj.search,
			method: "GET",
			headers: { Authorization: `Bearer ${token}` },
		}, (res) => {
			let data = "";
			res.on("data", (c: Buffer) => { data += c.toString(); });
			res.on("end", () => {
				try { resolve(JSON.parse(data)); }
				catch { reject(new Error(`Invalid JSON: ${data.substring(0, 300)}`)); }
			});
		});
		req.on("error", reject);
		req.end();
	});
}

async function startServer(): Promise<{ port: number; waitForCode: () => Promise<string> }> {
	return new Promise((resolve) => {
		let resolveCode: (code: string) => void;
		let rejectCode: (e: Error) => void;
		const codePromise = new Promise<string>((res, rej) => { resolveCode = res; rejectCode = rej; });

		const server = http.createServer((req, res) => {
			const url = new URL(req.url ?? "/", "http://localhost");
			const code = url.searchParams.get("code");
			const error = url.searchParams.get("error");
			if (code) {
				res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
				res.end("<html><body><h2>✅ Auth erfolgreich — Fenster schließen</h2></body></html>");
				server.close();
				resolveCode(code);
			} else if (error) {
				res.writeHead(400);
				res.end(`Fehler: ${error}`);
				server.close();
				rejectCode(new Error(error));
			} else {
				res.writeHead(404); res.end();
			}
		});

		server.listen(0, "localhost", () => {
			const port = (server.address() as net.AddressInfo).port;
			resolve({ port, waitForCode: () => codePromise });
		});
	});
}

async function main() {
	const { verifier, challenge } = generatePKCE();
	const { port, waitForCode } = await startServer();
	const redirectUri = `http://localhost:${port}`;

	const params = new URLSearchParams({
		client_id: CLIENT_ID,
		response_type: "code",
		redirect_uri: redirectUri,
		scope: "Calendars.Read offline_access",
		code_challenge: challenge,
		code_challenge_method: "S256",
		response_mode: "query",
	});

	const authUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?${params}`;

	console.log("\n=== MS365 Auth Test ===");
	console.log("Öffne diese URL im Browser:\n");
	console.log(authUrl);
	console.log("\nWarte auf Callback...");

	const code = await waitForCode();
	console.log("✅ Code erhalten, tausche gegen Token...");

	const body = new URLSearchParams({
		client_id: CLIENT_ID,
		grant_type: "authorization_code",
		code,
		redirect_uri: redirectUri,
		code_verifier: verifier,
		scope: "Calendars.Read offline_access",
	});

	const tokenData = await nodePost(TOKEN_URL, body.toString()) as {
		access_token?: string; expires_in?: number; error?: string; error_description?: string;
	};

	if (tokenData.error) {
		console.error("❌ Token Exchange Fehler:", tokenData.error_description || tokenData.error);
		process.exit(1);
	}

	console.log(`✅ Token erhalten (läuft in ${tokenData.expires_in}s ab)`);
	console.log("Rufe /me/calendars ab...");

	const calendars = await nodeGet(GRAPH_URL, tokenData.access_token!) as {
		value?: { id: string; name: string; isDefaultCalendar: boolean }[];
		error?: { message: string };
	};

	if (calendars.error) {
		console.error("❌ Graph Fehler:", calendars.error.message);
		process.exit(1);
	}

	console.log("\n✅ Kalender gefunden:");
	(calendars.value ?? []).forEach(c => {
		console.log(`  ${c.isDefaultCalendar ? "★" : " "} ${c.name} (${c.id.substring(0, 20)}...)`);
	});
	console.log("\n✅ Auth-Flow funktioniert vollständig.");
}

main().catch(e => { console.error("FATAL:", e); process.exit(1); });
