// Sets one or more people's passwords, hashed server-side. The browser
// never stores or transmits anything except the one new plaintext value
// at the moment someone types it — never a hash, never anyone else's.
//
// NOTE: there is no session/role check here yet — see README security
// notes. Anyone who can call this endpoint can currently set anyone's
// password. That's the same trust boundary the app already has on the
// client (only Master/P&O see the People UI that calls this), just not
// yet enforced a second time on the server. Flagging clearly rather than
// leaving it undocumented.
import { Redis } from "@upstash/redis";
import { hashPassword } from "../lib/hash.js";

const REDIS_URL =
  process.env.KV_REST_API_URL ||
  process.env.KV_URL ||
  process.env.UPSTASH_REDIS_REST_URL ||
  process.env.UPSTASH_REDIS_URL ||
  process.env.REDIS_REST_API_URL ||
  process.env.REDIS_URL ||
  process.env.STORAGE_KV_REST_API_URL ||
  process.env.STORAGE_KV_URL;
const REDIS_TOKEN =
  process.env.KV_REST_API_TOKEN ||
  process.env.KV_TOKEN ||
  process.env.UPSTASH_REDIS_REST_TOKEN ||
  process.env.UPSTASH_REDIS_TOKEN ||
  process.env.REDIS_REST_API_TOKEN ||
  process.env.REDIS_TOKEN ||
  process.env.STORAGE_KV_REST_API_TOKEN ||
  process.env.STORAGE_KV_TOKEN;

const redis = REDIS_URL && REDIS_TOKEN ? new Redis({ url: REDIS_URL, token: REDIS_TOKEN }) : null;

function envDiagnostics() {
  return Object.keys(process.env).filter((k) => /REDIS|KV_|UPSTASH/i.test(k));
}

const ROSTER_KEY = "gva-roster-v10";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const expectedSecret = process.env.API_SECRET;
  if (expectedSecret) {
    if (req.headers["x-api-secret"] !== expectedSecret) {
      return res.status(401).json({ error: "Unauthorized" });
    }
  }

  // updates: [{ id, newPassword }]
  const { updates } = req.body || {};
  if (!Array.isArray(updates) || !updates.length) {
    return res.status(400).json({ error: "updates array required" });
  }

  if (!redis) {
    return res.status(500).json({
      error: "Redis is not configured — no URL/token env var matched.",
      envVarsFound: envDiagnostics(),
    });
  }

  try {
    const raw = await redis.get(ROSTER_KEY);
    const roster = raw ? JSON.parse(typeof raw === "string" ? raw : JSON.stringify(raw)) : null;
    if (!roster || !Array.isArray(roster)) {
      return res.status(503).json({ error: "Roster not seeded yet" });
    }

    let applied = 0;
    const byId = new Map(roster.map((p, i) => [p.id, i]));
    for (const { id, newPassword } of updates) {
      if (!id || !newPassword) continue;
      const idx = byId.get(id);
      if (idx == null) continue;
      roster[idx] = { ...roster[idx], password: hashPassword(newPassword) };
      applied++;
    }

    await redis.set(ROSTER_KEY, JSON.stringify(roster));
    return res.status(200).json({ applied });
  } catch (err) {
    console.error("set-passwords handler error", err);
    return res.status(500).json({ error: "Failed to set passwords", detail: String(err?.message || err) });
  }
}
