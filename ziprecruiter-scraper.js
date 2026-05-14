import axios from 'axios';
import * as cheerio from 'cheerio';
import XLSX from 'xlsx';
import fs from 'fs';
import FormData from 'form-data';

// ====================== CONFIG ======================
const KEYWORDS = [
  "Analyst",
  "CFA",
  "CEO",
  "Data Science",
  "FP&A",
];

const CONFIG = {
  maxPages: 3,
  outputFile: "ZipRecruiter_Jobs.xlsx",
  delay: 3_000,
};

// Rotate ScraperAPI keys to avoid rate limits
const SCRAPER_KEYS = [
  "a4e059153392eaaf06a9b3f4babc2efc",
  "88e0fb0f3e8ea9bfc74e7e2c3774290d",
  "e313e263da60619bd30790b5ac483258",
  "f580a0a1b7c259634183ee1d7e970e58",
  "589bf9aee2cbffd22f04fb6dc07592b5",
  "a4879f66e9b5689e762c4fba46410a93",
  "34bb3b67fd52766b99d09cc308f5b191",
  "333d6263d4aa5f91efef17959b9e81c3",
  "1a86bdd0e53fc24f2b9b785665791669",
  "b74599ffb915fb16504f57cf970aa295",
  "ebe2940b8632051391e2cac76cd53717",
  "5431837a953ae44ea80d10c59db57073",
  "c220c564cfe88eb4f84c97df9d638bbe",
  "390b7a6aa207a1a3adf1f115f71cc1b4",
  "ca99962eeb24c51bc6bea2dec4ae56f3",
  "e87b431c2b2125b579bd3e6e0cc53ba3",
  "2d5e00e400bc46ce0d9303b8c14c71d3",
  "551372c9b08825f8b25dd5aff9e14016",
  "b92ad18d7632d50ce53f1ff109e2ff61",
  "b4604b622136a416f18ac4d90f39f285",
  "0cfa5cdb440efe7c31b58544881dacd4",
  "f158c06591080cec5a89ec51f6b6b60b",
  "567fd5673dd864fa1b8a47d50e33077d",
  "5df046aac5ca69642ff2c41e3536e5a7",
  "4446ba7e9e8fbb3a30fdce699d134004",
];

let keyIndex = 0;
function getKey() {
  const key = SCRAPER_KEYS[keyIndex % SCRAPER_KEYS.length];
  keyIndex++;
  return key;
}

// ====================== HELPERS ======================

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

// Build ScraperAPI proxied URL for a target URL
function scraperUrl(targetUrl) {
  const key = getKey();
  return `http://api.scraperapi.com?api_key=${key}&url=${encodeURIComponent(targetUrl)}&render=false&country_code=us`;
}

// ====================== NOTIFICATIONS ======================

async function uploadToCatbox(filePath) {
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      console.log(`📤 Uploading to Catbox (attempt ${attempt})...`);
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
        console.log(`✅ Catbox upload OK → ${link}`);
        return link;
      }
    } catch (err) {
      console.error(`❌ Catbox error (attempt ${attempt}):`, err.message);
      if (attempt < 3) await sleep(5_000);
    }
  }
  return null;
}

async function sendTeamsAlert(jobCount, fileLink = null) {
  const webhookUrl = process.env.TEAMS_WEBHOOK_URL;
  if (!webhookUrl) return;
  try {
    await axios.post(webhookUrl, {
      "@type": "MessageCard",
      "@context": "http://schema.org/extensions",
      themeColor: "0076D7",
      summary: "Job Scraper Report",
      sections: [{
        activityTitle: "🎯 ZipRecruiter Job Scraper",
        activitySubtitle: "Salary Filter Mode",
        facts: [{ name: "Jobs found:", value: `${jobCount}` }],
        text: fileLink ? `🔗 Download: ${fileLink}` : "",
      }],
    });
    console.log("✅ Teams notification sent");
  } catch (err) {
    console.error("❌ Teams error:", err.message);
  }
}

async function sendTelegramMessage(text) {
  const token = process.env.TELEGRAM_TOKEN;
  const chatId = process.env.TELEGRAM_CHAT_ID;
  if (!token || !chatId) return;
  try {
    await axios.post(`https://api.telegram.org/bot${token}/sendMessage`, {
      chat_id: chatId, text, parse_mode: 'HTML',
    });
  } catch (_) {}
}

async function sendTelegramFile(filePath) {
  const token = process.env.TELEGRAM_TOKEN;
  const chatId = process.env.TELEGRAM_CHAT_ID;
  if (!token || !chatId || !fs.existsSync(filePath)) return;
  const form = new FormData();
  form.append('chat_id', chatId);
  form.append('document', fs.createReadStream(filePath));
  try {
    await axios.post(`https://api.telegram.org/bot${token}/sendDocument`, form, { headers: form.getHeaders() });
  } catch (_) {}
}

// ====================== SCRAPE: ZIPRECRUITER ======================

async function scrapeZipRecruiter(keyword) {
  const jobs = [];

  for (let page = 1; page <= CONFIG.maxPages; page++) {
    const targetUrl =
      `https://www.ziprecruiter.com/jobs-search` +
      `?search=${encodeURIComponent(keyword)}` +
      `&location=United+States&page=${page}`;

    try {
      console.log(`  📄 ZipRecruiter page ${page}...`);
      const res = await axios.get(scraperUrl(targetUrl), { timeout: 60_000 });
      const $ = cheerio.load(res.data);

      // Save debug for first page of first keyword
      if (page === 1 && jobs.length === 0) {
        fs.writeFileSync(`debug_zip_${keyword.replace(/\s/g, '_')}.html`, res.data);
      }

      // ZipRecruiter job cards
      const cards = $('article[class*="job"], div[class*="job_result"], li[class*="job_result"]');
      console.log(`    📦 Cards: ${cards.length}`);

      cards.each((_, el) => {
        const card = $(el);
        const title = card.find('h2 a, a[class*="job_link"]').first().text().trim();
        if (!title || title.length < 4) return;

        const company  = card.find('[class*="company"]').first().text().trim() || 'N/A';
        const location = card.find('[class*="location"]').first().text().trim() || 'N/A';
        const link     = card.find('a[href*="/jobs/"], a[href*="/j/"]').first().attr('href') || '';
        const fullLink = link.startsWith('http') ? link.split('?')[0] : `https://www.ziprecruiter.com${link.split('?')[0]}`;

        const fullText = card.text().replace(/\s+/g, ' ');
        let salary = '';
        const patterns = [
          /(\$\d{1,3}(?:,\d{3})*(?:\.\d+)?(?:\s*[–\-]\s*\$\d{1,3}(?:,\d{3})*(?:\.\d+)?)?(?:\s*\/\s*(?:hr|hour|yr|year|mo|month|week|wk))?)/i,
          /(\d{2,3}[kK]\s*[–\-]\s*\d{2,3}[kK])/,
          /(\d{5,6}\s*[–\-]\s*\d{5,6})/,
        ];
        for (const pat of patterns) {
          const m = fullText.match(pat);
          if (m && m[0].length >= 5) { salary = m[0].trim(); break; }
        }
        if (!salary) return;

        jobs.push({ Title: title, Company: company, Salary: salary, Location: location, Posted: '', Link: fullLink, Keyword: keyword, Source: 'ZipRecruiter' });
      });

      console.log(`    💰 With salary so far: ${jobs.length}`);
      if (cards.length === 0) break;
      await sleep(CONFIG.delay);

    } catch (err) {
      console.error(`  ❌ ZipRecruiter error page ${page}:`, err.response?.status || err.message);
      await sleep(5_000);
    }
  }

  return jobs;
}

// ====================== SCRAPE: INDEED ======================

async function scrapeIndeed(keyword) {
  const jobs = [];

  for (let page = 0; page < CONFIG.maxPages; page++) {
    const start = page * 15;
    const targetUrl =
      `https://www.indeed.com/jobs?q=${encodeURIComponent(keyword)}&l=United+States&start=${start}&limit=15&filter=0`;

    try {
      console.log(`  📄 Indeed page ${page + 1}...`);
      const res = await axios.get(scraperUrl(targetUrl), { timeout: 60_000 });
      const $ = cheerio.load(res.data);

      if (page === 0 && jobs.length === 0) {
        fs.writeFileSync(`debug_indeed_${keyword.replace(/\s/g, '_')}.html`, res.data);
      }

      const cards = $('div.job_seen_beacon, div.resultContent, div[class*="job_seen"]');
      console.log(`    📦 Cards: ${cards.length}`);

      cards.each((_, el) => {
        const card = $(el);
        const title = card.find('h2.jobTitle span[title], h2 a span, span[title]').first().text().trim()
          || card.find('h2 a').text().trim();
        if (!title || title.length < 3) return;

        const company  = card.find('[data-testid="company-name"], span.companyName').first().text().trim() || 'N/A';
        const location = card.find('[data-testid="text-location"], div.companyLocation').first().text().trim() || 'N/A';
        const posted   = card.find('[class*="date"]').first().text().trim();

        const salaryEl = card.find('[class*="salary"], [data-testid*="salary"], div.salaryOnly').first().text().trim();
        let salary = salaryEl && salaryEl.length > 3 ? salaryEl : '';

        if (!salary) {
          const fullText = card.text().replace(/\s+/g, ' ');
          const patterns = [
            /(\$\d{1,3}(?:,\d{3})*(?:\.\d+)?(?:\s*[–\-a-z]\s*\$\d{1,3}(?:,\d{3})*(?:\.\d+)?)?(?:\s*(?:a year|a month|an hour|\/hr|\/yr|per hour|per year))?)/i,
            /(\d{2,3}[kK]\s*[–\-]\s*\d{2,3}[kK])/,
          ];
          for (const pat of patterns) {
            const m = fullText.match(pat);
            if (m && m[0].length >= 5) { salary = m[0].trim(); break; }
          }
        }
        if (!salary) return;

        const href = card.find('h2 a').first().attr('href') || '';
        const link = href.startsWith('http') ? href.split('?')[0] : `https://www.indeed.com${href.split('?')[0]}`;

        jobs.push({ Title: title, Company: company, Salary: salary, Location: location, Posted: posted, Link: link, Keyword: keyword, Source: 'Indeed' });
      });

      console.log(`    💰 With salary so far: ${jobs.length}`);
      if (cards.length === 0) break;
      await sleep(CONFIG.delay);

    } catch (err) {
      console.error(`  ❌ Indeed error page ${page + 1}:`, err.response?.status || err.message);
      await sleep(5_000);
    }
  }

  return jobs;
}

// ====================== MAIN ======================

async function runScraper() {
  console.log('🚀 Starting Job Scraper (ScraperAPI + ZipRecruiter & Indeed)...');

  const allJobs = [];

  for (const keyword of KEYWORDS) {
    console.log(`\n🔍 Keyword: "${keyword}"`);

    const [zipJobs, indeedJobs] = await Promise.all([
      scrapeZipRecruiter(keyword),
      scrapeIndeed(keyword),
    ]);

    const total = zipJobs.length + indeedJobs.length;
    allJobs.push(...zipJobs, ...indeedJobs);
    console.log(`  ✅ "${keyword}" → ${zipJobs.length} ZipRecruiter + ${indeedJobs.length} Indeed = ${total} jobs`);
    await sleep(2_000);
  }

  // De-duplicate by link
  const seen = new Set();
  const uniqueJobs = allJobs.filter(j => {
    if (seen.has(j.Link)) return false;
    seen.add(j.Link);
    return true;
  });

  console.log(`\n📊 Total unique jobs with salary: ${uniqueJobs.length}`);

  if (uniqueJobs.length > 0) {
    const ws = XLSX.utils.json_to_sheet(uniqueJobs);
    ws['!cols'] = Object.keys(uniqueJobs[0]).map(key => ({
      wch: Math.max(key.length, ...uniqueJobs.map(j => String(j[key] || '').length)) + 2,
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Jobs');
    XLSX.writeFile(wb, CONFIG.outputFile);
    console.log(`✅ Saved → ${CONFIG.outputFile}`);

    const fileLink = await uploadToCatbox(CONFIG.outputFile);
    await sendTelegramMessage(`✅ Found <b>${uniqueJobs.length}</b> jobs with salary!\n${fileLink ?? ''}`);
    await sendTeamsAlert(uniqueJobs.length, fileLink);
    await sendTelegramFile(CONFIG.outputFile);
  } else {
    await sendTelegramMessage('❌ No jobs with salary found.');
    await sendTeamsAlert(0);
    console.log('❌ No jobs with salary found.');
  }
}

runScraper().catch(console.error);

