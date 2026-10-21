// AES-GCM + PBKDF2-SHA256. The browser (index.html) uses the identical parameters.
import { webcrypto } from "node:crypto";
const subtle = webcrypto.subtle;
export const ITER = 250000;
const b64 = (u8) => Buffer.from(u8).toString("base64");
const unb64 = (s) => new Uint8Array(Buffer.from(s, "base64"));

async function key(pass, salt, usage) {
  const base = await subtle.importKey("raw", new TextEncoder().encode(pass.trim()), "PBKDF2", false, ["deriveKey"]);
  return subtle.deriveKey({ name: "PBKDF2", salt, iterations: ITER, hash: "SHA-256" }, base, { name: "AES-GCM", length: 256 }, false, usage);
}
export async function encryptJSON(obj, pass) {
  const salt = webcrypto.getRandomValues(new Uint8Array(16));
  const iv = webcrypto.getRandomValues(new Uint8Array(12));
  const k = await key(pass, salt, ["encrypt"]);
  const ct = await subtle.encrypt({ name: "AES-GCM", iv }, k, new TextEncoder().encode(JSON.stringify(obj)));
  return { v: 1, it: ITER, salt: b64(salt), iv: b64(iv), ct: b64(new Uint8Array(ct)) };
}
export async function decryptJSON(env, pass) {
  const k = await key(pass, unb64(env.salt), ["decrypt"]);
  const pt = await subtle.decrypt({ name: "AES-GCM", iv: unb64(env.iv) }, k, unb64(env.ct));
  return JSON.parse(new TextDecoder().decode(pt));
}
