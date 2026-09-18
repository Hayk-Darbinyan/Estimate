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

const CHECK_COMMENTS = "Մեկնաբանություն";
const DATE_GROUP_SIZE = 5;

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
  I: 9, // Մեկնաբանություն
};

function toDate(value) {
  if (value === null || value === undefined) return null;

  if (typeof value === "object" && value !== null) {
    if ("result" in value) return toDate(value.result);
    if ("text" in value) return toDate(value.text);
    if (value instanceof Date) return new Date(value.getTime());
  }

  if (value instanceof Date) return new Date(value.getTime());

  if (typeof value === "number") {
    const asDate = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
    return Number.isNaN(asDate.getTime()) ? null : asDate;
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;
    const asDate = new Date(trimmed);
    return Number.isNaN(asDate.getTime()) ? null : asDate;
  }

  return null;
}

function dateKey(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
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

function hasContent(value) {
  if (value === null || value === undefined) return false;

  if (typeof value === "object") {
    if ("result" in value) return hasContent(value.result);
    if ("text" in value) return hasContent(value.text);
  }

  return String(value).trim() !== "";
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function buildMessage(row) {
  return (
    `<b>1. Գնման ամսաթիվ:</b> ${escapeHtml(formatValue(row.A))}\n` +
    `<b>2. Հեռախոսահամար:</b> ${escapeHtml(formatValue(row.B))}\n` +
    `<b>3. Գնորդ:</b> ${escapeHtml(formatValue(row.C))}\n` +
    `<b>4. Գնած մոդել:</b> ${escapeHtml(formatValue(row.D))}\n` +
    `<b>5. Սպասարկող:</b> ${escapeHtml(formatValue(row.E))}\n` +
    `<b>6. Մեկնաբանություն:</b> ${escapeHtml(formatValue(row.I))}`
  );
}

function monthShortLabel(date) {
  return new Intl.DateTimeFormat("en-US", { month: "short" }).format(date);
}

function formatRangeLabel(startDate, endDate) {
  const sameMonth =
    startDate.getMonth() === endDate.getMonth() &&
    startDate.getFullYear() === endDate.getFullYear();

  if (sameMonth) {
    return `${startDate.getDate()}-${endDate.getDate()} ${monthShortLabel(startDate)} ${startDate.getFullYear()}`;
  }

  return `${startDate.getDate()} ${monthShortLabel(startDate)} - ${endDate.getDate()} ${monthShortLabel(endDate)} ${endDate.getFullYear()}`;
}

function buildDateRangeButtons(records) {
  const uniqueDates = [...new Set(records.map((record) => dateKey(record.A)))].sort(
    (a, b) => new Date(a) - new Date(b),
  );

  const groups = [];
  for (let index = 0; index < uniqueDates.length; index += DATE_GROUP_SIZE) {
    const slice = uniqueDates.slice(index, index + DATE_GROUP_SIZE);
    const startDate = new Date(slice[0]);
    const endDate = new Date(slice[slice.length - 1]);

    groups.push({
      start: slice[0],
      end: slice[slice.length - 1],
      label: formatRangeLabel(startDate, endDate),
    });
  }

  return groups
    .map((group) => [
      {
        text: group.label,
        callback_data: `date_range:${group.start}:${group.end}`,
      },
    ])
    .concat([[{ text: "Check All", callback_data: "check_all" }]]);
}

async function sendFilteredRecords(ctx, records) {
  if (records.length === 0) {
    await ctx.answerCbQuery("No matching records found for this date range.");
    return;
  }

  await ctx.answerCbQuery();

  for (const record of records) {
    await ctx.reply(buildMessage(record), {
      parse_mode: "HTML",
    });
    await sleep(150);
  }

  await ctx.reply(
    `Completed. Found ${records.length} matching record(s) for the selected date range.`,
  );
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

const bot = new Telegraf(BOT_TOKEN);
const selectedModeByChat = new Map();
const commentRecordsByChat = new Map();

const modeKeyboard = {
  reply_markup: {
    keyboard: [[CHECK_COMMENTS]],
    resize_keyboard: true,
  },
};

bot.start((ctx) => {
  ctx.reply(
    "Սեղմեք «Մեկնաբանություն» կոճակը, ապա ուղարկեք .xlsx ֆայլ։",
    modeKeyboard,
  );
});

bot.hears(CHECK_COMMENTS, (ctx) => {
  selectedModeByChat.set(ctx.chat.id, "comments");
  return ctx.reply(
    "Ընտրված է մեկնաբանությունների ստուգումը։ Ուղարկեք .xlsx ֆայլ։",
  );
});

bot.on("document", async (ctx) => {
  const mode = selectedModeByChat.get(ctx.chat.id);
  if (mode !== "comments") return;

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

    const records = [];

    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
      const excelRow = sheet.getRow(rowNumber);
      if (!excelRow.hasValues) continue;

      const dateValue = excelRow.getCell(COL.A).value;
      const commentValue = excelRow.getCell(COL.I).value;

      if (!hasContent(commentValue)) continue;

      const parsedDate = toDate(dateValue);
      if (!parsedDate) continue;

      const row = {};
      for (const [letter, columnIndex] of Object.entries(COL)) {
        row[letter] = excelRow.getCell(columnIndex).value;
      }

      row.A = parsedDate;
      records.push(row);
    }

    commentRecordsByChat.set(ctx.chat.id, records);

    if (records.length === 0) {
      await ctx.reply("Checked the uploaded file. No rows with comments in column I were found.");
      return;
    }

    await ctx.reply("Choose a date range to filter the comment results:", {
      reply_markup: {
        inline_keyboard: buildDateRangeButtons(records),
      },
    });
  } catch (err) {
    console.error(err);
    await ctx.reply(
      `Sorry, something went wrong while reading the file: ${err.message}`,
    );
  }
});

bot.action("check_all", async (ctx) => {
  const records = commentRecordsByChat.get(ctx.chat.id) || [];
  await sendFilteredRecords(ctx, records);
});

bot.action(/^date_range:([0-9-]+):([0-9-]+)$/, async (ctx, next) => {
  const startKey = ctx.match[1];
  const endKey = ctx.match[2];

  const records = commentRecordsByChat.get(ctx.chat.id) || [];
  const filtered = records.filter((record) => {
    const date = dateKey(record.A instanceof Date ? record.A : toDate(record.A));
    return date >= startKey && date <= endKey;
  });

  await sendFilteredRecords(ctx, filtered);
  return next();
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
