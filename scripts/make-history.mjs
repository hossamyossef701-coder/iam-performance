// Run this once, on your own Mac, to encrypt the back-record into history.enc.json.
// Nothing here touches the network. history-private.json stays on your Mac — .gitignore
// keeps it out of the repository.
//
//   node scripts/make-history.mjs "your passphrase here"
//
import fs from "node:fs";
import { encryptJSON } from "./crypto.mjs";

const pass = (process.argv[2] || process.env.DATA_PASSPHRASE || "").trim();
if (!pass) { console.error('Usage: node scripts/make-history.mjs "your passphrase here"'); process.exit(2); }
if (!fs.existsSync("history-private.json")) { console.error("history-private.json is not in this folder."); process.exit(2); }

const h = JSON.parse(fs.readFileSync("history-private.json", "utf8"));
for (const k of ["SEED", "HIST", "ROSTER_SEED", "GC_SEED"]) {
  if (!h[k]) { console.error(`history-private.json is missing "${k}".`); process.exit(2); }
}
fs.writeFileSync("history.enc.json", JSON.stringify(await encryptJSON(h, pass)));
console.log(`history.enc.json written (${Object.keys(h.HIST).length} months of history). Use this same passphrase for the DATA_PASSPHRASE secret and in the dashboard.`);
