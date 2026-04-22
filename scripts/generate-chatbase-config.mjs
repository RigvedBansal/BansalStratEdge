import { writeFileSync } from "node:fs";
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

writeFileSync(resolve(process.cwd(), "chatbase-config.js"), output);

console.log(
  `Generated chatbase-config.js with bot ID ${config.botId} and host ${config.host}.`
);
