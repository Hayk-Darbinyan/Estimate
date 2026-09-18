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

function monthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function monthLabel(date) {
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    year: "numeric",
  }).format(date);
}

function getUniqueMonths(records) {
  const months = [...new Set(records.map((record) => monthKey(record.A)))];
  return months.sort();
}

function buildMonthCalendar(state) {
  const availableDates = new Set(state.records.map((record) => dateKey(record.A)));
  const [year, month] = state.currentMonthKey.split("-").map(Number);
  const monthDate = new Date(year, month - 1, 1);
  const firstDay = new Date(year, month - 1, 1);
  const lastDay = new Date(year, month, 0);
  const dayNames = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"];

  const rows = [
    [
      { text: "⬅️", callback_data: `calendar_nav:prev:${state.currentMonthKey}` },
      { text: monthLabel(monthDate), callback_data: "calendar_month_label" },
      { text: "➡️", callback_data: `calendar_nav:next:${state.currentMonthKey}` },
    ],
    ...dayNames.map((day) => [{ text: day, callback_data: "calendar_day_label" }]),
  ];

  const startOffset = firstDay.getDay();
  const blanksBefore = Array.from({ length: startOffset }, () => ({ text: " ", callback_data: "calendar_empty" }));
  const datesInMonth = [];

  for (let day = 1; day <= lastDay.getDate(); day++) {
    const date = new Date(year, month - 1, day);
    const key = dateKey(date);
    const exists = availableDates.has(key);
    const isSelectable = exists && (!state.startDate || state.step === "end" ? true : true);

    if (exists) {
      const isBeforeStart =
        state.step === "end" && state.startDate && date < new Date(state.startDate.getTime());

      const callback = isBeforeStart ? "calendar_disabled" : `calendar:${state.step}:${key}`;
      const text = String(day);
      datesInMonth.push({ text, callback_data: callback, disabled: isBeforeStart });
    } else {
      datesInMonth.push({ text: " ", callback_data: "calendar_empty" });
    }
  }

  const cells = [...blanksBefore, ...datesInMonth];
  while (cells.length % 7 !== 0) {
    cells.push({ text: " ", callback_data: "calendar_empty" });
  }

  for (let index = 0; index < cells.length; index += 7) {
    rows.push(cells.slice(index, index + 7));
  }

  rows.push([{ text: "Check All", callback_data: "check_all" }]);
  return { inline_keyboard: rows };
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
const calendarStateByChat = new Map();

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

    const months = getUniqueMonths(records);
    calendarStateByChat.set(ctx.chat.id, {
      records,
      startDate: null,
      endDate: null,
      step: "start",
      currentMonthKey: months[0],
    });

    await ctx.reply("Select start date:", {
      reply_markup: buildMonthCalendar(calendarStateByChat.get(ctx.chat.id)),
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

bot.action(/^calendar_nav:(prev|next):([0-9]{4}-[0-9]{2})$/, async (ctx) => {
  const state = calendarStateByChat.get(ctx.chat.id);
  if (!state) return;

  const months = getUniqueMonths(state.records);
  const currentIndex = months.indexOf(state.currentMonthKey);
  let targetIndex = currentIndex;

  if (ctx.match[1] === "prev") {
    targetIndex = Math.max(0, currentIndex - 1);
  } else {
    targetIndex = Math.min(months.length - 1, currentIndex + 1);
  }

  state.currentMonthKey = months[targetIndex];
  await ctx.editMessageText(
    state.step === "start" ? "Select start date:" : "Select end date:",
    { reply_markup: buildMonthCalendar(state) },
  );
});

bot.action(/^calendar:(start|end):([0-9]{4}-[0-9]{2}-[0-9]{2})$/, async (ctx) => {
  const state = calendarStateByChat.get(ctx.chat.id);
  if (!state) return;

  const selectedDate = new Date(`${ctx.match[2]}T00:00:00`);

  if (ctx.match[1] === "start") {
    state.startDate = selectedDate;
    state.step = "end";
    state.currentMonthKey = monthKey(selectedDate);
    await ctx.editMessageText("Select end date:", {
      reply_markup: buildMonthCalendar(state),
    });
    return;
  }

  state.endDate = selectedDate;
  const records = commentRecordsByChat.get(ctx.chat.id) || [];
  const startKey = dateKey(state.startDate);
  const endKey = dateKey(state.endDate);
  const filtered = records.filter((record) => {
    const recordKey = dateKey(record.A instanceof Date ? record.A : toDate(record.A));
    return recordKey >= startKey && recordKey <= endKey;
  });

  calendarStateByChat.delete(ctx.chat.id);
  await sendFilteredRecords(ctx, filtered);
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
