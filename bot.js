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

// Parses M/D/YYYY  (the old format, e.g. 9/14/2026)
function normalizeDateString(value) {
  return String(value)
    .trim()
    .replace(/\s+/g, "")
    .replace(/[\u2024\u3002\uFF0E]/g, ".")
    .replace(/[\uFF0F]/g, "/");
}

function parseMDYString(value) {
  if (value === null || value === undefined) return null;

  const trimmed = normalizeDateString(value);
  const match = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!match) return null;

  const month = Number(match[1]);
  const day = Number(match[2]);
  const year = Number(match[3]);

  if (month < 1 || month > 12 || day < 1 || day > 31) return null;

  const parsedDate = new Date(year, month - 1, day);
  if (
    parsedDate.getFullYear() !== year ||
    parsedDate.getMonth() !== month - 1 ||
    parsedDate.getDate() !== day
  ) {
    return null;
  }

  return parsedDate;
}

// Parses DD.MM.YYYY  (the new format, e.g. 14.09.2026)
// Spaces are stripped first so "11. 09. 26"-style values also parse correctly.
function parseDMYString(value) {
  if (value === null || value === undefined) return null;

  const trimmed = normalizeDateString(value);
  const match = trimmed.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})$/);
  if (!match) return null;

  const day = Number(match[1]);
  const month = Number(match[2]);
  // Expand 2-digit years: 00-49 → 2000-2049, 50-99 → 1950-1999
  const rawYear = Number(match[3]);
  const year = rawYear < 100 ? (rawYear < 50 ? 2000 + rawYear : 1900 + rawYear) : rawYear;

  if (month < 1 || month > 12 || day < 1 || day > 31) return null;

  const parsedDate = new Date(year, month - 1, day);
  if (
    parsedDate.getFullYear() !== year ||
    parsedDate.getMonth() !== month - 1 ||
    parsedDate.getDate() !== day
  ) {
    return null;
  }

  return parsedDate;
}

function toDate(value) {
  if (value === null || value === undefined) return null;

  if (typeof value === "object" && value !== null) {
    if ("result" in value) return toDate(value.result);
    if ("text" in value) return toDate(value.text);
    if (value instanceof Date) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (value instanceof Date) return new Date(value.getFullYear(), value.getMonth(), value.getDate());

  if (typeof value === "number") {
    const asDate = new Date(1899, 11, 30);
    asDate.setDate(asDate.getDate() + value);
    return Number.isNaN(asDate.getTime()) ? null : asDate;
  }

  if (typeof value === "string") {
    const trimmed = normalizeDateString(value);
    if (!trimmed) return null;

    // Try DD.MM.YYYY first (spaces stripped inside parser), then M/D/YYYY (old format)
    const dmy = parseDMYString(trimmed);
    if (dmy) return dmy;

    const mdy = parseMDYString(trimmed);
    if (mdy) return mdy;

    // Last resort: let JS parse it (handles ISO strings etc.)
    const asDate = new Date(trimmed);
    return Number.isNaN(asDate.getTime()) ? null : asDate;
  }

  return null;
}

function dateKey(date) {
  const dateObject = date instanceof Date ? date : toDate(date);
  if (!dateObject) return null;

  const y = dateObject.getFullYear();
  const m = String(dateObject.getMonth() + 1).padStart(2, "0");
  const d = String(dateObject.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function dateFromKey(dateKeyValue) {
  const [year, month, day] = dateKeyValue.split("-").map(Number);
  return new Date(year, month - 1, day);
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
  const fields = [
    ["1. Գնման ամսաթիվ", row.A],
    ["2. Գնորդ", row.B],
    ["3. Հեռախոսահամար", row.C],
    ["4. Գնում/սպասարկում", row.D],
    ["5. Սպասարկող", row.E],
    ["6. Սպասարկման գնահատական", row.F],
    ["7. Գիտելիքի գնահատական", row.G],
    ["8. Գնահատական", row.I],
    ["9. Մեկնաբանություն", row.H],
  ];

  return fields
    .filter(([, value]) => hasContent(value))
    .map(([label, value]) => `<b>${label}:</b> ${escapeHtml(formatValue(value))}`)
    .join("\n");
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

function buildDateRangeButtons(dateKeys) {
  const uniqueDates = [...new Set(dateKeys)].filter(Boolean).sort();

  if (uniqueDates.length === 0) {
    return [
      [{ text: "Check All", callback_data: "check_all" }],
      [{ text: "Custom Date Range", callback_data: "custom_date_range" }],
    ];
  }

  const groups = [];
  for (let i = 0; i < uniqueDates.length; i += DATE_GROUP_SIZE) {
    const chunkKeys = uniqueDates.slice(i, i + DATE_GROUP_SIZE);
    const groupStart = dateFromKey(chunkKeys[0]);
    const groupEnd = dateFromKey(chunkKeys[chunkKeys.length - 1]);

    groups.push({
      start: chunkKeys[0],
      end: chunkKeys[chunkKeys.length - 1],
      label: formatRangeLabel(groupStart, groupEnd),
    });
  }

  return groups
    .map((group) => [
      {
        text: group.label,
        callback_data: `date_range:${group.start}:${group.end}`,
      },
    ])
    .concat([
      [{ text: "Check All", callback_data: "check_all" }],
      [{ text: "Custom Date Range", callback_data: "custom_date_range" }],
    ]);
}

async function sendFilteredRecords(ctx, records) {
  const isCallback = Boolean(ctx.callbackQuery);

  if (records.length === 0) {
    if (isCallback) {
      await ctx.answerCbQuery("No matching records found for this date range.");
    } else {
      await ctx.reply("No matching records found for this date range.");
    }
    return;
  }

  if (isCallback) {
    await ctx.answerCbQuery();
  }

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
const commentRecordsByChat = new Map();
const pendingCustomRangeByChat = new Map();

bot.start((ctx) => {
  ctx.reply("Ուղարկեք .xlsx ֆայլը՝ մեկնաբանությունները ստուգելու համար։", {
    reply_markup: {
      remove_keyboard: true,
    },
  });
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

    const headerColumns = new Map();
    sheet.getRow(1).eachCell((cell, columnNumber) => {
      const header = String(cell.value ?? "").trim();
      if (header) headerColumns.set(header, columnNumber);
    });

    const dateColumn = headerColumns.get("Գնման ամսաթիվ");
    const commentColumn = headerColumns.get(CHECK_COMMENTS);

    const getOptionalCellValue = (excelRow, header) => {
      const column = headerColumns.get(header);
      return column ? excelRow.getCell(column).value : undefined;
    };

    const records = [];
    const allDateKeys = new Set();

    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
      const excelRow = sheet.getRow(rowNumber);
      if (!excelRow.hasValues) continue;

      const dateValue = dateColumn
        ? excelRow.getCell(dateColumn).value
        : undefined;
      const parsedDate = toDate(dateValue);

      const commentValue = commentColumn
        ? excelRow.getCell(commentColumn).value
        : undefined;

      if (hasContent(commentValue)) {
        console.log(
          `[COMMENT ROW] row=${rowNumber} A=${formatValue(dateValue)} I=${formatValue(commentValue)}`,
        );
      }

      if (!hasContent(commentValue)) continue;
      if (!parsedDate) continue;

      const row = {
        A: parsedDate,
        B: getOptionalCellValue(excelRow, "Գնորդ"),
        C: getOptionalCellValue(excelRow, "Հեռախոսահամար"),
        D: getOptionalCellValue(excelRow, "Գնում/սպասարկում"),
        E: getOptionalCellValue(excelRow, "Սպասարկող"),
        F: getOptionalCellValue(excelRow, "Սպասարկման գնահատական"),
        G: getOptionalCellValue(excelRow, "Գիտելիքի գնահատական"),
        H: commentValue,
        I: getOptionalCellValue(excelRow, "Գնահատական"),
      };

      records.push(row);
      allDateKeys.add(dateKey(parsedDate));
    }

    commentRecordsByChat.set(ctx.chat.id, records);

    if (records.length === 0) {
      await ctx.reply("Checked the uploaded file. No rows with comments were found.");
      return;
    }

    await ctx.reply("Choose a date range to filter the comment results:", {
      reply_markup: {
        inline_keyboard: buildDateRangeButtons(allDateKeys),
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

bot.action("custom_date_range", async (ctx) => {
  pendingCustomRangeByChat.set(ctx.chat.id, true);
  await ctx.answerCbQuery();
  await ctx.reply(
    "Please enter your custom date range in DD.MM.YYYY format, for example: 01.09.2026 - 07.09.2026",
  );
});

// Handles both DD.MM.YYYY - DD.MM.YYYY, M/D/YYYY - M/D/YYYY, and alternate dot characters like 15․09․2026
bot.hears(
  /^(\d{1,2}[\/\.[\u2024\u2027\u00B7\u2219]]\d{1,2}[\/\.[\u2024\u2027\u00B7\u2219]]\d{4})\s*-\s*(\d{1,2}[\/\.[\u2024\u2027\u00B7\u2219]]\d{1,2}[\/\.[\u2024\u2027\u00B7\u2219]]\d{4})$/i,
  async (ctx) => {
    if (!pendingCustomRangeByChat.get(ctx.chat.id)) {
      return;
    }

    pendingCustomRangeByChat.delete(ctx.chat.id);

    const startRaw = ctx.match[1];
    const endRaw = ctx.match[2];

    // Try DD.MM.YYYY first, then M/D/YYYY
    const startDate = parseDMYString(startRaw) || parseMDYString(startRaw);
    const endDate = parseDMYString(endRaw) || parseMDYString(endRaw);

    if (!startDate || !endDate) {
      await ctx.reply(
        "Invalid date format. Please use DD.MM.YYYY, for example: 01.09.2026 - 07.09.2026",
      );
      return;
    }

    const startKey = dateKey(startDate);
    const endKey = dateKey(endDate);

    if (startKey > endKey) {
      await ctx.reply("The start date cannot be after the end date. Please enter the range again.");
      return;
    }

    const records = commentRecordsByChat.get(ctx.chat.id) || [];
    const filtered = records.filter((record) => {
      const date = dateKey(record.A instanceof Date ? record.A : toDate(record.A));
      return date >= startKey && date <= endKey;
    });

    await sendFilteredRecords(ctx, filtered);
  },
);

bot.action(/^date_range:([0-9-]+):([0-9-]+)$/, async (ctx, next) => {
  const startKey = ctx.match[1];
  const endKey = ctx.match[2];

  const records = commentRecordsByChat.get(ctx.chat.id) || [];
  const filtered = records.filter((record) => {
    const recordDate = dateKey(record.A instanceof Date ? record.A : toDate(record.A));
    return recordDate >= startKey && recordDate <= endKey;
  });

  await sendFilteredRecords(ctx, filtered);
  return next();
});

bot.catch((err, ctx) => {
  console.error(`Telegraf error for update ${ctx.updateType}:`, err);
});

const app = express();

app.use(express.json());

app.get("/", (req, res) => {
  res.status(200).send("Bot is running.");
});

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

process.once("SIGINT", () => process.exit(0));
process.once("SIGTERM", () => process.exit(0));
