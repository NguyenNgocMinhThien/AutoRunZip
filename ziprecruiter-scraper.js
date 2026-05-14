import { chromium } from 'playwright';
import XLSX from 'xlsx';
import fs from 'fs';
import FormData from 'form-data';
import axios from 'axios';

// ====================== CONFIG ======================
const KEYWORDS = [
  "Analyst",
  "CFA",
  "CEO",
  "Data Science",
  "FP&A",
];

const CONFIG = {
  pageSize: 20,          // ZipRecruiter shows ~20 jobs/page
  maxPages: 3,           // pages per keyword
  maxRetries: 3,
  retryDelay: 10_000,    // ms
  pageTimeout: 90_000,
  scrollDelay: 5_000,
  outputFile: "ZipRecruiter_Jobs.xlsx",
};

// ====================== HELPERS ======================

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function uploadToCatbox(filePath) {
  for (let attempt = 1; attempt <= CONFIG.maxRetries; attempt++) {
    try {
      console.log(`📤 Uploading to Catbox (attempt ${attempt})...`);
      const form = new FormData();
      form.append('reqtype', 'fileupload');
      form.append('time', '72h');
      form.append('fileToUpload', fs.createReadStream(filePath));

      const res = await axios.post(
        'https://litterbox.catbox.moe/resources/internals/api.php',
        form,
        {
          headers: form.getHeaders(),
          timeout: 45_000,
          maxBodyLength: Infinity,
          maxContentLength: Infinity,
        }
      );

      const link = res.data.trim();
      if (link.includes('https://')) {
        console.log(`✅ Catbox upload OK → ${link}`);
        return link;
      }
    } catch (err) {
      console.error(`❌ Catbox error (attempt ${attempt}):`, err.message);
      if (attempt < CONFIG.maxRetries) await sleep(5_000);
    }
  }
  console.error('❌ Catbox upload failed after all attempts');
  return null;
}

// ====================== NOTIFICATIONS ======================

async function sendTeamsAlert(jobCount, fileLink = null) {
  const webhookUrl = process.env.TEAMS_WEBHOOK_URL;
  if (!webhookUrl) return;
  try {
    await axios.post(webhookUrl, {
      "@type": "MessageCard",
      "@context": "http://schema.org/extensions",
      themeColor: "0076D7",
      summary: "ZipRecruiter Scraper Report",
      sections: [{
        activityTitle: "🎯 ZipRecruiter Scraper",
        activitySubtitle: "Salary Filter Mode",
        facts: [
          { name: "Jobs found:", value: `${jobCount}` },
        ],
        text: fileLink ? `🔗 Download: ${fileLink}` : "",
      }],
    });
    console.log("✅ Teams notification sent");
  } catch (err) {
    console.error("❌ Teams error:", err.message);
  }
}

async function sendTelegramMessage(text) {
  const token  = process.env.TELEGRAM_TOKEN;
  const chatId = process.env.TELEGRAM_CHAT_ID;
  if (!token || !chatId) return;
  try {
    await axios.post(`https://api.telegram.org/bot${token}/sendMessage`, {
      chat_id: chatId,
      text,
      parse_mode: 'HTML',
    });
  } catch (_) {}
}

async function sendTelegramFile(filePath) {
  const token  = process.env.TELEGRAM_TOKEN;
  const chatId = process.env.TELEGRAM_CHAT_ID;
  if (!token || !chatId || !fs.existsSync(filePath)) return;
  const form = new FormData();
  form.append('chat_id', chatId);
  form.append('document', fs.createReadStream(filePath));
  try {
    await axios.post(
      `https://api.telegram.org/bot${token}/sendDocument`,
      form,
      { headers: form.getHeaders() }
    );
  } catch (_) {}
}

// ====================== SCRAPE LOGIC ======================

/**
 * Extract job cards from the current ZipRecruiter search page.
 * Returns an array of job objects.
 */
async function extractJobs(page, keyword) {
  return page.evaluate((kw) => {
    const jobs = [];

    // ZipRecruiter job cards – selector may vary; we cast a wide net
    const cards = document.querySelectorAll(
      'article[class*="job_result"], ' +
      'div[class*="jobList"] article, ' +
      'div[data-testid="job-card"], ' +
      'div[class*="job-listing"], ' +
      'li[class*="job_result"]'
    );

    cards.forEach(card => {
      // ── Title ──────────────────────────────────────────
      const titleEl =
        card.querySelector('h2 a, h2, a[class*="job_link"], a[data-testid="job-title"]');
      const title = titleEl?.textContent?.trim();
      if (!title || title.length < 5) return;

      // ── Link ───────────────────────────────────────────
      const linkEl = card.querySelector('a[href*="/jobs/"], a[href*="ziprecruiter.com"]');
      const link   = linkEl ? linkEl.href.split('?')[0] : '';

      // ── Company ────────────────────────────────────────
      const companyEl = card.querySelector(
        'a[class*="company"], span[class*="company"], ' +
        '[data-testid="job-card-hiring-company"], ' +
        'div[class*="hiring_company"] a'
      );
      const company = companyEl?.textContent?.trim() || 'N/A';

      // ── Location ───────────────────────────────────────
      const locationEl = card.querySelector(
        'a[class*="location"], span[class*="location"], ' +
        '[data-testid="job-card-location"], ' +
        'div[class*="location"]'
      );
      const location = locationEl?.textContent?.trim() || 'N/A';

      // ── Salary ─────────────────────────────────────────
      const fullText = card.textContent.replace(/\s+/g, ' ');
      let salary = '';

      const salaryPatterns = [
        /(\$\d{1,3}(?:,\d{3})*(?:\.\d+)?(?:\s*[–\-]\s*\$\d{1,3}(?:,\d{3})*(?:\.\d+)?)?(?:\s*\/\s*(?:hr|hour|yr|year|mo|month|week|wk))?)/i,
        /(\d{2,3}[kK]\s*[–\-]\s*\d{2,3}[kK])/,
        /(\d{5,6}\s*[–\-]\s*\d{5,6})/,
        /USD\s*\d+/i,
      ];

      for (const pat of salaryPatterns) {
        const m = fullText.match(pat);
        if (m && m[0].length >= 4) { salary = m[0].trim(); break; }
      }

      if (!salary) return; // skip jobs without salary

      // ── Job type / posted ──────────────────────────────
      const typeEl   = card.querySelector('[data-testid="job-type"], span[class*="job_type"]');
      const postedEl = card.querySelector('[data-testid="job-posted"], span[class*="posted"]');

      jobs.push({
        Title:    title,
        Company:  company,
        Salary:   salary,
        Location: location,
        Type:     typeEl?.textContent?.trim()   || '',
        Posted:   postedEl?.textContent?.trim() || '',
        Link:     link,
        Keyword:  kw,
      });
    });

    return {
      total: cards.length,
      withSalary: jobs.length,
      jobs,
    };
  }, keyword);
}

// ====================== MAIN ======================

async function runScraper() {
  console.log('🚀 Starting ZipRecruiter Scraper...');

  const browser = await chromium.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });

  const allJobs = [];

  for (const keyword of KEYWORDS) {
    console.log(`\n🔍 Keyword: "${keyword}"`);

    for (let pageNum = 1; pageNum <= CONFIG.maxPages; pageNum++) {
      let attempt = 0;

      while (attempt < CONFIG.maxRetries) {
        attempt++;
        console.log(`  📄 Page ${pageNum} (attempt ${attempt})...`);

        let page;
        try {
          page = await browser.newPage();
          await page.setViewportSize({ width: 1920, height: 1080 });

          // ── Extra headers to reduce bot-detection ───────
          await page.setExtraHTTPHeaders({
            'Accept-Language': 'en-US,en;q=0.9',
            'User-Agent':
              'Mozilla/5.0 (Windows NT 10.0; Win64; x64) ' +
              'AppleWebKit/537.36 (KHTML, like Gecko) ' +
              'Chrome/124.0.0.0 Safari/537.36',
          });

          const url =
            `https://www.ziprecruiter.com/jobs-search` +
            `?search=${encodeURIComponent(keyword)}` +
            `&location=United+States` +
            `&page=${pageNum}`;

          try {
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: CONFIG.pageTimeout });
} catch (_) {
  // timeout on full load is ok — grab what rendered
}
await sleep(12_000);
await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
await sleep(6_000);

          const result = await extractJobs(page, keyword);

          console.log(`    🔗 Cards found: ${result.total} | 💰 With salary: ${result.withSalary}`);

          if (result.jobs.length > 0) {
            allJobs.push(...result.jobs);
            // Sample preview
            console.log('    📋 Sample:', result.jobs[0]);
          }

          await page.close();

          // If no jobs at all on this page, stop pagination
          if (result.total === 0) {
            console.log('    ⚠️  No cards found – stopping pagination for this keyword');
            pageNum = CONFIG.maxPages + 1;
          }
          break; // success → next page

        } catch (err) {
          console.error(`    ❌ Error: ${err.message}`);
          if (page) await page.close().catch(() => {});
          if (attempt < CONFIG.maxRetries) await sleep(CONFIG.retryDelay);
        }
      }
    }
  }

  await browser.close();

  // ── De-duplicate by Link ─────────────────────────────
  const seen = new Set();
  const uniqueJobs = allJobs.filter(j => {
    if (seen.has(j.Link)) return false;
    seen.add(j.Link);
    return true;
  });

  console.log(`\n📊 Total unique jobs with salary: ${uniqueJobs.length}`);

  if (uniqueJobs.length > 0) {
    // Write Excel
    const ws = XLSX.utils.json_to_sheet(uniqueJobs);

    // Auto-size columns
    const colWidths = Object.keys(uniqueJobs[0]).map(key => ({
      wch: Math.max(key.length, ...uniqueJobs.map(j => String(j[key] || '').length)) + 2,
    }));
    ws['!cols'] = colWidths;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Jobs');
    XLSX.writeFile(wb, CONFIG.outputFile);
    console.log(`✅ Saved → ${CONFIG.outputFile}`);

    const fileLink = await uploadToCatbox(CONFIG.outputFile);
    const msg = `✅ ZipRecruiter: Found <b>${uniqueJobs.length}</b> jobs with salary!\n${fileLink ?? ''}`;

    await sendTelegramMessage(msg);
    await sendTeamsAlert(uniqueJobs.length, fileLink);
    await sendTelegramFile(CONFIG.outputFile);
  } else {
    const msg = '❌ ZipRecruiter: No jobs with salary found.';
    await sendTelegramMessage(msg);
    await sendTeamsAlert(0);
    console.log(msg);
  }
}

runScraper().catch(console.error);