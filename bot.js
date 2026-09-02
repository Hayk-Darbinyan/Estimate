import express from "express";
import { Telegraf } from "telegraf";
import ExcelJS from "exceljs";
import crypto from "crypto";
import dotenv from "dotenv";

dotenv.config();

const BOT_TOKEN = process.env.BOT_TOKEN;
const PORT = process.env.PORT || 3000;

// Render sets this automatically for every web service. Falls back to a
// manually-set WEBHOOK_URL for local testing with a tunnel (e.g. ngrok).
const PUBLIC_URL = process.env.RENDER_EXTERNAL_URL || process.env.WEBHOOK_URL;

if (!BOT_TOKEN) {
  console.error(
    "BOT_TOKEN environment variable is not set. Get a token from @BotFather and set it before running.",
  );
  process.exit(1);
}

if (!PUBLIC_URL) {
  console.error(
    "No public URL found. On Render this comes from RENDER_EXTERNAL_URL automatically. " +
      "For local/manual runs, set WEBHOOK_URL to a publicly reachable https URL (e.g. an ngrok tunnel).",
  );
  process.exit(1);
}

const THRESHOLD = 5;

// A random-ish but stable path so random internet traffic can't hit your
// webhook endpoint and pretend to be Telegram. Derived from the bot token
// so it doesn't need its own env var, but isn't guessable without the token.
const WEBHOOK_PATH = `/telegraf/${crypto
  .createHash("sha256")
  .update(BOT_TOKEN)
  .digest("hex")
  .slice(0, 32)}`;

// Telegram can also send a secret header we can verify on every request,
// as extra protection against spoofed webhook calls.
const WEBHOOK_SECRET = crypto
  .createHash("sha256")
  .update(`${BOT_TOKEN}:secret`)
  .digest("hex")
  .slice(0, 32);

// Column letters -> 1-indexed column numbers (ExcelJS uses 1-indexed columns)
const COL = {
  A: 1, // Գնման ամսաթիվ
  B: 2, // Հեռախոսահամար
  C: 3, // Գնորդ
  D: 4, // Գնած մոդել
  E: 5, // Սպասարկող
  F: 6, // Սպասարկման գնահատական
  G: 7, // Գիտելիքի գնահատական
  I: 9, // Մեկնաբանություն
};

function toNumber(value) {
  if (value === null || value === undefined) return null;

  // ExcelJS may give { result: ... } for formula cells
  if (typeof value === "object" && value !== null && "result" in value) {
    return toNumber(value.result);
  }

  if (typeof value === "number") return value;

  if (typeof value === "string") {
    const trimmed = value.trim().replace(",", ".");
    if (trimmed === "") return null;
    const num = Number(trimmed);
    return Number.isNaN(num) ? null : num;
  }

  return null;
}

function formatDate(date) {
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = date.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

function formatValue(value) {
  if (value === null || value === undefined || value === "") return "-";

  if (value instanceof Date) return formatDate(value);

  if (typeof value === "object" && value !== null) {
    if ("result" in value) return formatValue(value.result);
    if ("text" in value) return String(value.text);
  }

  return String(value).trim();
}

function buildMessage(row) {
  return (
    `*1. Գնման ամսաթիվ* ${formatValue(row.A)}\n` +
    `*2. Հեռախոսահամար* ${formatValue(row.B)}\n` +
    `*3. Գնորդ* ${formatValue(row.C)}\n` +
    `*4. Գնած մոդել* ${formatValue(row.D)}\n` +
    `*5. Սպասարկող* ${formatValue(row.E)}\n` +
    `*6. Սպասարկման գնահատական* ${formatValue(row.F)}\n` +
    `*7. Գիտելիքի գնահատական* ${formatValue(row.G)}\n` +
    `*8. Մեկնաբանություն* ${formatValue(row.I)}`
  );
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

const bot = new Telegraf(BOT_TOKEN);

bot.start((ctx) => {
  ctx.reply(
    "Ուղարկեք .xlsx ֆայլ, ես կստուգեմ F և G սյուները։\n\n" +
      "Send me an .xlsx file and I'll check columns F and G. " +
      `Any row where F or G has a numeric value lower than ${THRESHOLD} ` +
      "will be sent back to you as a separate message.",
  );
});

bot.on("document", async (ctx) => {
  const doc = ctx.message.document;
  const fileName = (doc.file_name || "").toLowerCase();

  if (!fileName.endsWith(".xlsx")) {
    await ctx.reply("Please upload a valid .xlsx (Excel) file.");
    return;
  }

  await ctx.reply("Processing your file, please wait...");

  try {
    const fileLink = await ctx.telegram.getFileLink(doc.file_id);
    const response = await fetch(fileLink.href ?? fileLink);
    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.worksheets[workbook.worksheets.length - 1];

    let matches = 0;
    let rowsChecked = 0;

    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
      const excelRow = sheet.getRow(rowNumber);

      // Skip fully empty rows
      if (!excelRow.hasValues) continue;

      rowsChecked++;

      const get = (letter) => excelRow.getCell(COL[letter]).value;

      const fRaw = get("F");
      const gRaw = get("G");
      const fVal = toNumber(fRaw);
      const gVal = toNumber(gRaw);
      const isDash = (v) => String(v ?? "").trim() === "-";
      const isLow =
        (fVal !== null && fVal < THRESHOLD) ||
        (gVal !== null && gVal < THRESHOLD) ||
        isDash(fRaw) ||
        isDash(gRaw);

      if (isLow) {
        const rowValues = {};
        for (const letter of Object.keys(COL)) {
          rowValues[letter] = get(letter);
        }
        await ctx.reply(buildMessage(rowValues), {
          parse_mode: "Markdown",
        });
        matches++;
        // small delay to be gentle on Telegram's rate limits
        await sleep(150);
      }
    }

    if (matches === 0) {
      await ctx.reply(
        `Checked ${rowsChecked} rows. No rows found with F or G below ${THRESHOLD}.`,
      );
    } else {
      await ctx.reply(
        `Done. Checked ${rowsChecked} rows, found ${matches} matching row(s) above.`,
      );
    }
  } catch (err) {
    console.error(err);
    await ctx.reply(
      `Sorry, something went wrong while reading the file: ${err.message}`,
    );
  }
});

bot.catch((err, ctx) => {
  console.error(`Telegraf error for update ${ctx.updateType}:`, err);
});

// ---------------------------------------------------------------------
// Express server: this is what makes the app a valid Render Web Service.
// Render requires the process to bind to process.env.PORT and answer
// HTTP requests — that's how it knows the service is alive, and it's also
// what lets Telegram's webhook calls (or an external pinger) wake it up.
// ---------------------------------------------------------------------
const app = express();

// Telegraf needs the raw JSON body of incoming updates.
app.use(express.json());

// Health check / keep-alive endpoint. Render's own health checks hit this,
// and you can optionally point an external uptime pinger (see README) at
// it to reduce how often the instance goes to sleep.
app.get("/", (req, res) => {
  res.status(200).send("Bot is running.");
});

// Telegram webhook endpoint. Only requests carrying the correct secret
// header are treated as genuine Telegram traffic.
app.post(WEBHOOK_PATH, (req, res, next) => {
  const incomingSecret = req.header("X-Telegram-Bot-Api-Secret-Token");
  if (incomingSecret !== WEBHOOK_SECRET) {
    return res.sendStatus(401);
  }
  return next();
});
app.use(bot.webhookCallback(WEBHOOK_PATH));

async function main() {
  const webhookUrl = `${PUBLIC_URL.replace(/\/+$/, "")}${WEBHOOK_PATH}`;

  await bot.telegram.setWebhook(webhookUrl, {
    secret_token: WEBHOOK_SECRET,
  });

  const info = await bot.telegram.getWebhookInfo();
  console.log("Webhook set to:", info.url);

  app.listen(PORT, () => {
    console.log(`Server listening on port ${PORT}`);
  });
}

main().catch((err) => {
  console.error("Failed to start bot:", err);
  process.exit(1);
});

// Graceful shutdown
process.once("SIGINT", () => process.exit(0));
process.once("SIGTERM", () => process.exit(0));
