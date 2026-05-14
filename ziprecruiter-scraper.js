import axios from 'axios';
import * as cheerio from 'cheerio';
import XLSX from 'xlsx';
import fs from 'fs';
import FormData from 'form-data';

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
  // Thêm location cụ thể nếu muốn:
  // "Los Angeles, CA",
  // "New York, NY",
];

const CONFIG = {
  maxPages:      3,
  maxPerKeyword: 10,
  outputFile:    "ZipRecruiter_Jobs.xlsx",
  delay:         3_000,
  concurrency:   3,
  fetchDetail:   true,
};

// Đặt 0 = lấy tất cả job có salary. Tăng lên để lọc, ví dụ: 80000 / 40
const MIN_SALARY_YEAR = 0;
const MIN_SALARY_HOUR = 0;
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

      if (page === 1) {
        const debugFile = `debug_zip_${keyword.replace(/\W/g, '_')}_${location.replace(/\W/g, '_')}.html`;
        fs.writeFileSync(debugFile, res.data);
      }

      let cards = $('article[class*="job"]');
      if (!cards.length) cards = $('[data-testid="job-card"]');
      if (!cards.length) cards = $('div[class*="job_result"]');
      if (!cards.length) cards = $('li[class*="job_result"]');
      if (!cards.length) cards = $('[class*="jobList"] > li, [class*="jobs-list"] > li');

      console.log(`    📦 Cards tìm được: ${cards.length}`);
      if (cards.length === 0) {
        const allJobLike = $('[class*="Job"], [class*="job-card"], [class*="JobCard"]');
        console.log(`    🔍 Fallback selector: ${allJobLike.length} phần tử`);
        break;
      }

      const cardList = [];
      cards.each((_, el) => {
        if (cardList.length >= CONFIG.maxPerKeyword) return false;
        const card = $(el);

        const titleEl = card.find('h2 a, a[class*="job_link"], [class*="title"] a, h2[class*="title"]').first();
        const title = titleEl.text().trim() || card.find('h2').first().text().trim();
        if (!title || title.length < 3) return;

        const company = card.find('[class*="company"], [data-testid*="company"], [class*="hiring"]')
          .first().text().trim() || 'N/A';

        const loc = card.find('[class*="location"], [data-testid*="location"]')
          .first().text().trim() || location;

        let href = titleEl.attr('href') || card.find('a[href*="/jobs/"], a[href*="/j/"]').first().attr('href') || '';
        if (!href) href = card.find('a').first().attr('href') || '';
        const link = href.startsWith('http') ? href.split('?')[0] : `https://www.ziprecruiter.com${href.split('?')[0]}`;

        const salary = parseSalary($, el);
        const quick  = card.find('[class*="quick"], [class*="easy"], [class*="1-Click"]').length > 0;

        cardList.push({ title, company, loc, link, salary, quick });
      });

      if (CONFIG.fetchDetail) {
        const detailTasks = cardList
          .filter(c => !c.salary && c.link && c.link !== 'https://www.ziprecruiter.com')
          .map(c => async () => { c.salary = await fetchDetailSalary(c.link); });
        if (detailTasks.length) {
          console.log(`    🔎 Fetch detail salary cho ${detailTasks.length} card...`);
          await parallelLimit(detailTasks, 3, 1000);
        }
      }

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
          Source:      'ZipRecruiter',
        }));

      jobs.push(...qualified);
      console.log(`    💰 Có salary đủ điều kiện: ${jobs.length}`);

      if (cards.length < 10) break;
      await sleep(CONFIG.delay);

    } catch (err) {
      console.error(`  ❌ Lỗi trang ${page}:`, err.response?.status || err.message);
      await sleep(5_000);
    }
  }

  console.log(`  ✅ [${location}] "${keyword}" → ${jobs.length} jobs`);
  return jobs;
}

// ==================== UPLOAD ====================

async function uploadToCatbox(filePath) {
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      console.log(`📤 Upload lên Catbox (lần ${attempt})...`);
      const form = new FormData();
      form.append('reqtype', 'fileupload');
      form.append('time', '72h');
      form.append('fileToUpload', fs.createReadStream(filePath));
      const res = await axios.post(
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

// ==================== MS TEAMS ====================

async function sendToTeams(jobCount, fileLink) {
  const webhookUrl = process.env.TEAMS_WEBHOOK_URL;
  if (!webhookUrl) {
    console.warn("⚠️  Thiếu TEAMS_WEBHOOK_URL — bỏ qua thông báo Teams");
    return;
  }
  try {
    await axios.post(webhookUrl, {
      "@type":    "MessageCard",
      "@context": "http://schema.org/extensions",
      themeColor: "0076D7",
      summary:    "ZipRecruiter Job Scraper Report",
      sections: [{
        activityTitle:    "🎯 ZipRecruiter Job Scraper",
        activitySubtitle: `Kết quả ngày ${TODAY}`,
        facts: [
          { name: "Nguồn:",    value: "ZipRecruiter" },
          { name: "Số job:",   value: `${jobCount}` },
          { name: "Keywords:", value: KEYWORDS.join(", ") },
          { name: "Ngày:",     value: TODAY },
        ],
        text: fileLink ? `📥 **Tải file Excel:** [ZipRecruiter_Jobs.xlsx](${fileLink})` : "",
      }],
    });
    console.log("✅ [Teams] Gửi thông báo thành công!");
  } catch (e) {
    console.error("❌ [Teams] Lỗi:", e.message);
  }
}

// ==================== MAIN ====================

async function runScraper() {
  console.log("🚀 ZipRecruiter Scraper — US");
  console.log(`📋 ${KEYWORDS.length} keywords × ${LOCATIONS.length} location | Concurrency: ${CONFIG.concurrency}\n`);

  if (!process.env.SCRAPER_API_KEY) {
    console.error("❌ Thiếu SCRAPER_API_KEY! Hãy set env variable.");
    process.exit(1);
  }

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
    await sendToTeams(0, null);
    console.log("❌ Không có job nào. Kiểm tra debug HTML để xem selector có đúng không.");
    return;
  }

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

  // Upload Catbox rồi gửi Teams
  const fileLink = await uploadToCatbox(fileName);
  await sendToTeams(allJobs.length, fileLink);

  console.log("🏁 Hoàn tất!");
}

runScraper().catch(console.error);