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
  maxPages: 3,          // Indeed shows 15 jobs/page
  outputFile: "ZipRecruiter_Jobs.xlsx",
  delay: 4_000,         // ms between requests
};

// ====================== HELPERS ======================

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

// Rotate user agents to avoid blocks
const USER_AGENTS = [
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0',
  'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
];
function randomUA() {
  return USER_AGENTS[Math.floor(Math.random() * USER_AGENTS.length)];
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
      summary: "Indeed Scraper Report",
      sections: [{
        activityTitle: "🎯 Indeed Job Scraper",
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

// ====================== SCRAPE LOGIC ======================

async function fetchIndeedJobs(keyword) {
  const jobs = [];

  for (let page = 0; page < CONFIG.maxPages; page++) {
    const start = page * 15; // Indeed paginates by 15
    const url = `https://www.indeed.com/jobs?q=${encodeURIComponent(keyword)}&l=United+States&start=${start}&limit=15&filter=0`;

    try {
      console.log(`  📄 Page ${page + 1} (start=${start})...`);

      const res = await axios.get(url, {
        headers: {
          'User-Agent': randomUA(),
          'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
          'Accept-Language': 'en-US,en;q=0.9',
          'Accept-Encoding': 'gzip, deflate, br',
          'Referer': 'https://www.indeed.com/',
          'DNT': '1',
        },
        timeout: 30_000,
      });

      const $ = cheerio.load(res.data);

      // Save debug HTML for first keyword first page only
      if (page === 0 && jobs.length === 0) {
        fs.writeFileSync(`debug_${keyword.replace(/\s/g, '_')}.html`, res.data);
        console.log(`    💾 Debug HTML saved`);
      }

      // Indeed job cards
      const cards = $('div.job_seen_beacon, div[class*="jobsearch-ResultsList"] > li, div.resultContent');
      console.log(`    📦 Cards found: ${cards.length}`);

      cards.each((_, el) => {
        const card = $(el);

        // Title
        const title = card.find('h2.jobTitle span[title], h2.jobTitle a span, span[title]').first().text().trim()
          || card.find('h2 a').text().trim();
        if (!title || title.length < 3) return;

        // Company
        const company = card.find('[data-testid="company-name"], span.companyName, [class*="companyName"]').first().text().trim() || 'N/A';

        // Location
        const location = card.find('[data-testid="text-location"], div.companyLocation, [class*="companyLocation"]').first().text().trim() || 'N/A';

        // Salary — Indeed often shows it in a dedicated element
        const salaryEl = card.find(
          '[class*="salary"], [data-testid*="salary"], ' +
          'div.metadata.salary-snippet-container, ' +
          'div.salaryOnly, span.salaryText'
        ).first().text().trim();

        let salary = '';
        if (salaryEl && salaryEl.length > 3) {
          salary = salaryEl;
        } else {
          // Fallback: regex on full card text
          const fullText = card.text().replace(/\s+/g, ' ');
          const patterns = [
            /(\$\d{1,3}(?:,\d{3})*(?:\.\d+)?(?:\s*[–\-a-z]\s*\$\d{1,3}(?:,\d{3})*(?:\.\d+)?)?(?:\s*(?:a year|a month|an hour|\/hr|\/yr|\/mo|per hour|per year))?)/i,
            /(\d{2,3}[kK]\s*[–\-]\s*\d{2,3}[kK])/,
            /(\d{5,6}\s*[–\-]\s*\d{5,6})/,
          ];
          for (const pat of patterns) {
            const m = fullText.match(pat);
            if (m && m[0].length >= 5) { salary = m[0].trim(); break; }
          }
        }

        if (!salary) return; // skip jobs without salary

        // Link
        const linkEl = card.find('h2 a, a[id^="job_"]').first();
        const href = linkEl.attr('href') || '';
        const link = href.startsWith('http') ? href.split('?')[0] : `https://www.indeed.com${href.split('?')[0]}`;

        // Posted
        const posted = card.find('[class*="date"], span.date').first().text().trim();

        jobs.push({ Title: title, Company: company, Salary: salary, Location: location, Posted: posted, Link: link, Keyword: keyword });
      });

      console.log(`    💰 With salary so far: ${jobs.length}`);

      if (cards.length === 0) {
        console.log(`    ⚠️  No cards – stopping pagination`);
        break;
      }

      await sleep(CONFIG.delay);

    } catch (err) {
      console.error(`  ❌ Error page ${page + 1}:`, err.response?.status || err.message);
      await sleep(5_000);
    }
  }

  return jobs;
}

// ====================== MAIN ======================

async function runScraper() {
  console.log('🚀 Starting Indeed Job Scraper (salary filter)...');

  const allJobs = [];

  for (const keyword of KEYWORDS) {
    console.log(`\n🔍 Keyword: "${keyword}"`);
    const jobs = await fetchIndeedJobs(keyword);
    allJobs.push(...jobs);
    console.log(`  ✅ "${keyword}" → ${jobs.length} jobs with salary`);
    await sleep(3_000);
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
    await sendTelegramMessage(`✅ Indeed: Found <b>${uniqueJobs.length}</b> jobs with salary!\n${fileLink ?? ''}`);
    await sendTeamsAlert(uniqueJobs.length, fileLink);
    await sendTelegramFile(CONFIG.outputFile);
  } else {
    await sendTelegramMessage('❌ Indeed: No jobs with salary found.');
    await sendTeamsAlert(0);
    console.log('❌ No jobs with salary found.');
  }
}

runScraper().catch(console.error);