/**
 * Telegram bot: upload an .xlsx file, the bot checks columns F and G for
 * each row, and if either value is a number lower than 5, it sends that
 * row back to the user as a separate message.
 *
 * Requires Node.js 18+ (uses native fetch).
 *
 * Setup:
 *   1. npm install
 *   2. Get a bot token from @BotFather on Telegram.
 *   3. Set it as an environment variable:
 *        export BOT_TOKEN="123456:ABC-your-token-here"
 *      (Windows CMD: set BOT_TOKEN=123456:ABC-your-token-here)
 *   4. Run:
 *        npm start
 *
 * Usage:
 *   - Open a chat with your bot on Telegram.
 *   - Send /start to see instructions.
 *   - Send it an .xlsx file as a document.
 *   - The bot replies with one message per row where column F or G
 *     contains a number lower than 5.
 */

import { Telegraf } from "telegraf";
import ExcelJS from "exceljs";
import dotenv from "dotenv";

dotenv.config();
const BOT_TOKEN = process.env.BOT_TOKEN;

if (!BOT_TOKEN) {
  console.error(
    "BOT_TOKEN environment variable is not set. Get a token from @BotFather and set it before running."
  );
  process.exit(1);
}

const THRESHOLD = 5;

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
      "will be sent back to you as a separate message."
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
    const sheet = workbook.worksheets[0];

    let matches = 0;
    let rowsChecked = 0;

    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
      const excelRow = sheet.getRow(rowNumber);

      // Skip fully empty rows
      if (!excelRow.hasValues) continue;

      rowsChecked++;

      const get = (letter) => excelRow.getCell(COL[letter]).value;

      const fVal = toNumber(get("F"));
      const gVal = toNumber(get("G"));

      const isLow =
        (fVal !== null && fVal < THRESHOLD) ||
        (gVal !== null && gVal < THRESHOLD);

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
        `Checked ${rowsChecked} rows. No rows found with F or G below ${THRESHOLD}.`
      );
    } else {
      await ctx.reply(
        `Done. Checked ${rowsChecked} rows, found ${matches} matching row(s) above.`
      );
    }
  } catch (err) {
    console.error(err);
    await ctx.reply(
      `Sorry, something went wrong while reading the file: ${err.message}`
    );
  }
});

bot.launch().then(() => console.log("Bot is up and running..."));

// Enable graceful stop
process.once("SIGINT", () => bot.stop("SIGINT"));
process.once("SIGTERM", () => bot.stop("SIGTERM"));