// Generic key-value proxy backing src/storageShim.js. Backed by Upstash
// Redis — Vercel's own "Vercel KV" product is deprecated; the current
// path is Vercel Marketplace -> add a Redis (Upstash) integration ->
// connect it to this project. That injects the REST URL/token below
// under one of a couple of possible env var names depending on how the
// integration was added, so this checks both rather than guessing wrong.

import { Redis } from "@upstash/redis";

const REDIS_URL =
  process.env.KV_REST_API_URL ||
  process.env.UPSTASH_REDIS_REST_URL ||
  process.env.REDIS_REST_API_URL ||
  process.env.STORAGE_KV_REST_API_URL;
const REDIS_TOKEN =
  process.env.KV_REST_API_TOKEN ||
  process.env.UPSTASH_REDIS_REST_TOKEN ||
  process.env.REDIS_REST_API_TOKEN ||
  process.env.STORAGE_KV_REST_API_TOKEN;

const redis = REDIS_URL && REDIS_TOKEN ? new Redis({ url: REDIS_URL, token: REDIS_TOKEN }) : null;

function envDiagnostics() {
  return Object.keys(process.env).filter((k) => /REDIS|KV_|UPSTASH/i.test(k));
}

// NOTE ON SECURITY: this endpoint currently only checks a single shared
// secret (API_SECRET), not who the caller actually is. That's enough to
// stop random internet traffic from finding this URL and reading/writing
// your data, but it is NOT per-user authorization — anyone who has the
// app open can read or write any key, including other people's data.
// Real authorization (e.g. only admins can write the roster) needs to
// move server-side before this holds sensitive data at scale.

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  if (process.env.API_SECRET) {
    const provided = req.headers["x-api-secret"];
    if (provided !== process.env.API_SECRET) {
      res.status(401).json({ error: "Unauthorized" });
      return;
    }
  }

  if (!redis) {
    return res.status(500).json({
      error: "Redis is not configured — no URL/token env var matched.",
      envVarsFound: envDiagnostics(),
    });
  }

  const { op, key, value, prefix } = req.body || {};

  try {
    if (op === "get") {
      if (!key) return res.status(400).json({ error: "key required" });
      const v = await redis.get(key);
      return res.status(200).json({ value: v ?? null });
    }

    if (op === "set") {
      if (!key) return res.status(400).json({ error: "key required" });
      await redis.set(key, value);
      return res.status(200).json({ ok: true });
    }

    if (op === "delete") {
      if (!key) return res.status(400).json({ error: "key required" });
      await redis.del(key);
      return res.status(200).json({ ok: true });
    }

    if (op === "list") {
      const pattern = prefix ? `${prefix}*` : "*";
      const keys = await redis.keys(pattern);
      return res.status(200).json({ keys });
    }

    return res.status(400).json({ error: `Unknown op: ${op}` });
  } catch (err) {
    console.error("kv handler error", err);
    return res.status(500).json({ error: "Storage operation failed", detail: String(err?.message || err) });
  }
}
