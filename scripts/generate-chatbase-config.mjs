import {
  cpSync,
  existsSync,
  mkdirSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { extname, resolve } from "node:path";

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

const rootStaticExtensions = new Set([
  ".html",
  ".css",
  ".js",
  ".svg",
  ".txt",
  ".xml",
  ".webmanifest",
]);

const filesToCopy = readdirSync(rootDir).filter((file) => {
  const filePath = resolve(rootDir, file);

  if (!statSync(filePath).isFile()) {
    return false;
  }

  if (file === "chatbase-config.js") {
    return false;
  }

  return rootStaticExtensions.has(extname(file));
});

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
