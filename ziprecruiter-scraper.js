import axios from 'axios';
import * as cheerio from 'cheerio';
import XLSX from 'xlsx';
import fs from 'fs';
import FormData from 'form-data';
import { google } from 'googleapis';

// ==================== CẤU HÌNH ====================
const KEYWORDS = [
  "Analyst",
  "CFA",
  "CEO",
  "Data Science",
  "FP&A",
];

const LOCATIONS = [
  "United States",
  // Thêm location cụ thể nếu muốn, ví dụ:
  // "Los Angeles, CA",
  // "New York, NY",
];

const CONFIG = {
  maxPages:       3,
  maxPerKeyword:  10,
  outputFile:     "ZipRecruiter_Jobs.xlsx",
  delay:          3_000,
  concurrency:    3,
  fetchDetail:    true,
};

const MIN_SALARY_YEAR = 0;   // Đặt 0 = lấy tất cả job có salary; tăng lên ví dụ 80000 để lọc
const MIN_SALARY_HOUR = 0;

const SPREADSHEET_ID = '1vUcKAbDazlC_vFSjzty02Fdugu4Nw_jZsEG2k_wyxXY';
const SHEET_NAME     = 'Job ZipRecruiter';
const SHEET_GID      = '0';
// =====================================================

const now   = new Date();
const dd    = String(now.getDate()).padStart(2, '0');
const mm    = String(now.getMonth() + 1).padStart(2, '0');
const yyyy  = now.getFullYear();
const TODAY = `${dd}/${mm}/${yyyy}`;

// ==================== HELPERS ====================

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function parallelLimit(tasks, limit, delayMs = 1500) {
  const results = new Array(tasks.length);
  let idx = 0;
  async function worker() {
    while (idx < tasks.length) {
      const i = idx++;
      results[i] = await Promise.resolve().then(tasks[i]);
      if (idx < tasks.length) await sleep(delayMs);
    }
  }
  await Promise.all(Array.from({ length: limit }, worker));
  return results;
}

function dedup(jobs) {
  const seen = new Set();
  return jobs.filter(j => {
    const key = `${j.Title}|${j.Company}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function salaryQualifies(salaryText) {
  // Nếu MIN đều = 0 thì chỉ cần có salary là lấy
  if (!salaryText || salaryText === 'N/A') return false;
  if (MIN_SALARY_YEAR === 0 && MIN_SALARY_HOUR === 0) return true;

  const isHour = /hour|hr|\/hr/i.test(salaryText);
  const isYear = /year|yr|annual|\/yr/i.test(salaryText);
  const isWeek = /week|\/wk/i.test(salaryText);
  const isMon  = /month|\/mo/i.test(salaryText);
  const cleaned = salaryText.replace(/,/g, '');
  const nums = [...cleaned.matchAll(/(\d+(?:\.\d+)?)/g)]
    .map(m => parseFloat(m[1])).filter(n => n > 0);
  if (!nums.length) return false;
  const maxNum = Math.max(...nums);
  if (isHour) return maxNum >= MIN_SALARY_HOUR;
  if (isYear) return maxNum >= MIN_SALARY_YEAR;
  if (isWeek) return (maxNum * 52) >= MIN_SALARY_YEAR;
  if (isMon)  return (maxNum * 12) >= MIN_SALARY_YEAR;
  if (maxNum >= 1000) return maxNum >= MIN_SALARY_YEAR;
  return maxNum >= MIN_SALARY_HOUR;
}

async function scraperGet(url) {
  return axios.get('https://api.scraperapi.com/', {
    params: {
      api_key:      process.env.SCRAPER_API_KEY,
      url,
      country_code: 'us',
      render:       'true',
      keep_headers: 'true',
    },
    timeout: 120_000,
  });
}

// ==================== PARSE SALARY ====================

function parseSalary($, cardEl) {
  const card = cardEl ? $(cardEl) : $('body');
  const fullText = card.text().replace(/\s+/g, ' ');

  // Selectors ZipRecruiter thường dùng
  const selectors = [
    '[class*="salary"]',
    '[class*="Salary"]',
    '[data-testid*="salary"]',
    '.compensation',
    '[class*="compensation"]',
    '[class*="pay"]',
  ];
  for (const sel of selectors) {
    const t = card.find(sel).first().text().replace(/\s+/g, ' ').trim();
    if (t && t.includes('$')) return cleanSalary(t);
  }

  // Fallback: regex từ toàn bộ text của card
  const patterns = [
    /(\$[\d,]+(?:\.\d+)?\s*(?:[-–]\s*\$[\d,]+(?:\.\d+)?)?\s*(?:a year|an hour|per year|per hour|\/hr|\/year|\/yr|\/mo|a month))/i,
    /(\$\d{1,3}(?:,\d{3})*(?:\.\d+)?(?:\s*[–\-]\s*\$\d{1,3}(?:,\d{3})*(?:\.\d+)?)?(?:\s*\/\s*(?:hr|hour|yr|year|mo|month|week|wk))?)/i,
    /(\d{2,3}[kK]\s*[–\-]\s*\d{2,3}[kK])/,
    /(\d{5,6}\s*[–\-]\s*\d{5,6})/,
  ];
  for (const pat of patterns) {
    const m = fullText.match(pat);
    if (m && m[0].length >= 5) return cleanSalary(m[0]);
  }
  return '';
}

function cleanSalary(s) {
  return s.replace(/Full-time|Part-time|Permanent|Contract|Temporary/gi, '')
          .replace(/\+\d+/g, '').replace(/\s+/g, ' ').trim();
}

async function fetchDetailSalary(link) {
  try {
    const res = await scraperGet(link);
    const $ = cheerio.load(res.data);
    return parseSalary($, null);
  } catch { return ''; }
}

// ==================== SCRAPE ZIPRECRUITER ====================

async function scrapeKeywordLocation(keyword, location) {
  const jobs = [];

  for (let page = 1; page <= CONFIG.maxPages; page++) {
    const targetUrl =
      `https://www.ziprecruiter.com/jobs-search` +
      `?search=${encodeURIComponent(keyword)}` +
      `&location=${encodeURIComponent(location)}` +
      `&page=${page}`;

    try {
      console.log(`  📄 [${location}] "${keyword}" — trang ${page}...`);
      const res = await scraperGet(targetUrl);
      const $   = cheerio.load(res.data);

      // Debug: lưu HTML trang đầu để kiểm tra selector nếu cần
      if (page === 1) {
        const debugFile = `debug_zip_${keyword.replace(/\W/g, '_')}_${location.replace(/\W/g, '_')}.html`;
        fs.writeFileSync(debugFile, res.data);
      }

      // Các selector phổ biến của ZipRecruiter
      let cards = $('article[class*="job"]');
      if (!cards.length) cards = $('[data-testid="job-card"]');
      if (!cards.length) cards = $('div[class*="job_result"]');
      if (!cards.length) cards = $('li[class*="job_result"]');
      if (!cards.length) cards = $('[class*="jobList"] > li, [class*="jobs-list"] > li');

      console.log(`    📦 Cards tìm được: ${cards.length}`);
      if (cards.length === 0) {
        // Thử fallback: tìm tất cả thẻ có class chứa "job"
        const allJobLike = $('[class*="Job"], [class*="job-card"], [class*="JobCard"]');
        console.log(`    🔍 Fallback selector: ${allJobLike.length} phần tử`);
        break;
      }

      const cardList = [];
      cards.each((_, el) => {
        if (cardList.length >= CONFIG.maxPerKeyword) return false;
        const card = $(el);

        // Title
        const titleEl = card.find('h2 a, a[class*="job_link"], [class*="title"] a, h2[class*="title"]').first();
        const title = titleEl.text().trim() || card.find('h2').first().text().trim();
        if (!title || title.length < 3) return;

        // Company
        const company = card.find('[class*="company"], [data-testid*="company"], [class*="hiring"]')
          .first().text().trim() || 'N/A';

        // Location
        const loc = card.find('[class*="location"], [data-testid*="location"]')
          .first().text().trim() || location;

        // Link
        let href = titleEl.attr('href') || card.find('a[href*="/jobs/"], a[href*="/j/"]').first().attr('href') || '';
        if (!href) href = card.find('a').first().attr('href') || '';
        const link = href.startsWith('http') ? href.split('?')[0] : `https://www.ziprecruiter.com${href.split('?')[0]}`;

        // Salary từ card
        const salary = parseSalary($, el);

        // Quick apply
        const quick = card.find('[class*="quick"], [class*="easy"], [class*="1-Click"]').length > 0;

        cardList.push({ title, company, loc, link, salary, quick });
      });

      // Fetch detail salary song song cho card chưa có salary
      if (CONFIG.fetchDetail) {
        const detailTasks = cardList
          .filter(c => !c.salary && c.link && c.link !== 'https://www.ziprecruiter.com')
          .map(c => async () => {
            c.salary = await fetchDetailSalary(c.link);
          });
        if (detailTasks.length) {
          console.log(`    🔎 Fetch detail salary cho ${detailTasks.length} card...`);
          await parallelLimit(detailTasks, 3, 1000);
        }
      }

      // Lọc theo salary
      const qualified = cardList
        .filter(c => salaryQualifies(c.salary))
        .map(c => ({
          Company:     c.company,
          Title:       c.title,
          Salary:      c.salary,
          Location:    c.loc,
          Link:        c.link,
          EasilyApply: c.quick ? 'Quick Apply' : 'Company Website',
          Keyword:     keyword,
          DateCrawled: TODAY,
          CrawledBy:   '',
          Source:      'ZipRecruiter',
        }));

      jobs.push(...qualified);
      console.log(`    💰 Có salary đủ điều kiện: ${jobs.length}`);

      if (cards.length < 10) break; // ít card = trang cuối
      await sleep(CONFIG.delay);

    } catch (err) {
      console.error(`  ❌ Lỗi trang ${page}:`, err.response?.status || err.message);
      await sleep(5_000);
    }
  }

  console.log(`  ✅ [${location}] "${keyword}" → ${jobs.length} jobs`);
  return jobs;
}

// ==================== GOOGLE SHEETS ====================

async function appendToGoogleSheet(jobs) {
  try {
    const serviceAccountJson = process.env.GDRIVE_SERVICE_ACCOUNT_JSON;
    if (!serviceAccountJson) { console.warn("⚠️  Thiếu GDRIVE_SERVICE_ACCOUNT_JSON — bỏ qua Google Sheets"); return; }

    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(serviceAccountJson),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const sheets = google.sheets({ version: 'v4', auth });
    const rows = jobs.map(j => [
      j.Company, j.Title, j.Salary, j.Location, j.Link,
      j.EasilyApply, j.Keyword, `'${j.DateCrawled}`, j.CrawledBy, j.Source,
    ]);
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range:         `${SHEET_NAME}!A:J`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: rows },
    });
    console.log(`✅ [Google Sheets] Đã append ${jobs.length} rows vào "${SHEET_NAME}"`);
  } catch (e) {
    console.error("❌ [Google Sheets] Lỗi:", e.message);
  }
}

// ==================== UPLOAD & NOTIFY ====================

async function uploadToCatbox(filePath) {
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      console.log(`📤 Upload lên Catbox (lần ${attempt})...`);
      const form = new FormData();
      form.append('reqtype', 'fileupload');
      form.append('time', '72h');
      form.append('fileToUpload', fs.createReadStream(filePath));
      const res  = await axios.post(
        'https://litterbox.catbox.moe/resources/internals/api.php',
        form,
        { headers: form.getHeaders(), timeout: 45_000, maxBodyLength: Infinity, maxContentLength: Infinity }
      );
      const link = res.data.trim();
      if (link.includes('https://')) {
        console.log(`✅ Catbox OK → ${link}`);
        return link;
      }
    } catch (e) {
      console.error(`❌ Catbox lần ${attempt}:`, e.message);
      if (attempt < 3) await sleep(5_000);
    }
  }
  return `https://github.com/${process.env.GITHUB_REPOSITORY ?? 'unknown'}/actions`;
}

async function sendToTeams(n, fileLink) {
  const url = process.env.TEAMS_WEBHOOK_URL;
  if (!url) return;
  try {
    await axios.post(url, {
      type: "AdaptiveCard", version: "1.4",
      body: [
        { type: "TextBlock", text: "🚀 JOB MỚI — ZIPRECRUITER US", weight: "Bolder", size: "Medium", color: "Accent" },
        { type: "FactSet", facts: [
          { title: "Nguồn:",   value: "ZipRecruiter" },
          { title: "Số job:",  value: `${n}` },
          { title: "Status:",  value: "✅ Đã ghi Google Sheets" },
          { title: "Ngày:",    value: TODAY },
        ]},
      ],
      actions: [
        { type: "Action.OpenUrl", title: "📊 Mở Google Sheet",
          url: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit?gid=${SHEET_GID}#gid=${SHEET_GID}` },
        { type: "Action.OpenUrl", title: "📥 Tải Excel", url: fileLink },
      ],
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    });
    console.log("✅ [Teams] Gửi thành công!");
  } catch (e) { console.error("❌ [Teams]:", e.message); }
}

async function sendTelegramMessage(text) {
  const { TELEGRAM_TOKEN: t, TELEGRAM_CHAT_ID: c } = process.env;
  if (!t || !c) return;
  try {
    await axios.post(`https://api.telegram.org/bot${t}/sendMessage`, {
      chat_id: c, text, parse_mode: 'HTML',
    });
  } catch (e) { console.error("❌ Telegram:", e.message); }
}

async function sendTelegramFile(filePath) {
  const { TELEGRAM_TOKEN: t, TELEGRAM_CHAT_ID: c } = process.env;
  if (!t || !c || !fs.existsSync(filePath)) return;
  const form = new FormData();
  form.append('chat_id', c);
  form.append('document', fs.createReadStream(filePath));
  try {
    await axios.post(`https://api.telegram.org/bot${t}/sendDocument`, form, { headers: form.getHeaders() });
    console.log("✅ [Telegram] File đã gửi!");
  } catch (e) { console.error("❌ Telegram File:", e.message); }
}

// ==================== MAIN ====================

async function runScraper() {
  console.log("🚀 ZipRecruiter Scraper — US");
  console.log(`📋 ${KEYWORDS.length} keywords × ${LOCATIONS.length} location | Concurrency: ${CONFIG.concurrency} | Sheet: "${SHEET_NAME}"\n`);

  if (!process.env.SCRAPER_API_KEY) {
    console.error("❌ Thiếu SCRAPER_API_KEY! Hãy set env variable.");
    process.exit(1);
  }

  // Tạo tất cả task (keyword × location) rồi chạy song song
  const allTasks = [];
  for (const kw of KEYWORDS) {
    for (const loc of LOCATIONS) {
      allTasks.push(() => scrapeKeywordLocation(kw, loc));
    }
  }

  console.log(`⚡ Chạy ${allTasks.length} task song song (${CONFIG.concurrency} cùng lúc)...\n`);
  const results = await parallelLimit(allTasks, CONFIG.concurrency);

  let allJobs = results.flat();
  allJobs = dedup(allJobs);

  console.log(`\n📦 Tổng unique jobs: ${allJobs.length}`);

  if (!allJobs.length) {
    await sendTelegramMessage("❌ ZipRecruiter: Không tìm được job nào có salary.");
    console.log("❌ Không có job nào. Kiểm tra debug HTML để xem selector có đúng không.");
    return;
  }

  // Ghi Google Sheets
  await appendToGoogleSheet(allJobs);

  // Tạo Excel
  const fileName = CONFIG.outputFile;
  const ws = XLSX.utils.json_to_sheet(allJobs);
  ws['!cols'] = Object.keys(allJobs[0]).map(k => ({
    wch: Math.min(60, Math.max(k.length + 2, ...allJobs.map(r => String(r[k] || '').length))),
  }));
  const lastCol = String.fromCharCode(64 + Object.keys(allJobs[0]).length);
  ws['!autofilter'] = { ref: `A1:${lastCol}1` };

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Jobs");
  XLSX.writeFile(wb, fileName);
  console.log(`📊 Đã lưu Excel → ${fileName}`);

  // Upload & thông báo
  const fileLink = await uploadToCatbox(fileName);
  await Promise.all([
    sendTelegramMessage(
      `✅ <b>ZipRecruiter US</b>\n` +
      `📦 <b>${allJobs.length} jobs</b> có salary\n` +
      `📊 <a href="https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit?gid=${SHEET_GID}#gid=${SHEET_GID}">Mở Google Sheet</a>\n` +
      `📎 <a href="${fileLink}">Tải Excel</a>`
    ),
    sendTelegramFile(fileName),
    sendToTeams(allJobs.length, fileLink),
  ]);

  console.log("🏁 Hoàn tất!");
}

runScraper().catch(console.error);