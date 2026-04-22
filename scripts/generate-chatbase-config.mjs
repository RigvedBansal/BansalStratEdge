import {
  cpSync,
  existsSync,
  mkdirSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { resolve } from "node:path";

const defaults = {
  botId: "wDye2cLkm_hH2e4lyjq_F",
  host: "www.chatbase.co",
};

const config = {
  botId: process.env.NEXT_PUBLIC_CHATBOT_ID || defaults.botId,
  host: process.env.NEXT_PUBLIC_CHATBASE_HOST || defaults.host,
};

const output = `window.CHATBASE_CONFIG = ${JSON.stringify(config, null, 2)};
window.CHATBASE_BOT_ID = window.CHATBASE_BOT_ID || window.CHATBASE_CONFIG.botId;
window.CHATBASE_HOST = window.CHATBASE_HOST || window.CHATBASE_CONFIG.host;
`;

const rootDir = process.cwd();
const publicDir = resolve(rootDir, "public");

const filesToCopy = [
  "index.html",
  "blogs.html",
  "founders.html",
  "cfo-capital-efficiency-checklist.html",
  "journal-ai-control-tower.html",
  "journal-board-narratives.html",
  "journal-capital-confidence.html",
  "journal-cfo-dashboards.html",
  "journal-daily-confidence.html",
  "journal-growth-architecture.html",
  "journal-treasury-mandate.html",
  "styles.css",
  "script.js",
  "speed-insights.js",
  "favicon.svg",
];

rmSync(publicDir, { recursive: true, force: true });
mkdirSync(publicDir, { recursive: true });

writeFileSync(resolve(rootDir, "chatbase-config.js"), output);
writeFileSync(resolve(publicDir, "chatbase-config.js"), output);

for (const file of filesToCopy) {
  cpSync(resolve(rootDir, file), resolve(publicDir, file));
}

const assetsDir = resolve(rootDir, "assets");

if (existsSync(assetsDir) && statSync(assetsDir).isDirectory()) {
  cpSync(assetsDir, resolve(publicDir, "assets"), { recursive: true });
}

console.log(
  `Generated public build with Chatbase bot ID ${config.botId} and host ${config.host}.`
);
