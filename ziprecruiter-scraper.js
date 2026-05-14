import axios from 'axios';
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
  jobsPerPage: 20,
  outputFile: "ZipRecruiter_Jobs.xlsx",
};

// ZipRecruiter public job search API
const API_URL = "https://api.ziprecruiter.com/jobs/v1";
const API_KEY = "aunzHt4sMnGNzMEBeCR5KhOWsGfG4gip"; // public key embedded in their web app

// ====================== HELPERS ======================

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

function extractSalary(job) {
  // Try structured salary fields first
  if (job.salary_min && job.salary_max) {
    const min = Number(job.salary_min);
    const max = Number(job.salary_max);
    if (min > 0 && max > 0) {
      const fmt = n => `$${n.toLocaleString('en-US')}`;
      const interval = job.salary_interval || '';
      const label = interval === 'hour' ? '/hr' : interval === 'week' ? '/wk' : interval === 'month' ? '/mo' : '/yr';
      return `${fmt(min)} - ${fmt(max)}${label}`;
    }
  }
  if (job.salary_min && Number(job.salary_min) > 0) {
    return `$${Number(job.salary_min).toLocaleString('en-US')}+`;
  }

  // Fallback: regex on job snippet / description
  const text = [job.snippet, job.job_description, job.name].filter(Boolean).join(' ');
  const patterns = [
    /(\$\d{1,3}(?:,\d{3})*(?:\.\d+)?(?:\s*[–\-]\s*\$\d{1,3}(?:,\d{3})*(?:\.\d+)?)?(?:\s*\/\s*(?:hr|hour|yr|year|mo|month|week|wk))?)/i,
    /(\d{2,3}[kK]\s*[–\-]\s*\d{2,3}[kK])/,
    /(\d{5,6}\s*[–\-]\s*\d{5,6})/,
  ];
  for (const pat of patterns) {
    const m = text.match(pat);
    if (m) return m[0].trim();
  }
  return '';
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
      summary: "ZipRecruiter Scraper Report",
      sections: [{
        activityTitle: "🎯 ZipRecruiter Scraper",
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

async function fetchJobsForKeyword(keyword) {
  const jobs = [];

  for (let page = 1; page <= CONFIG.maxPages; page++) {
    try {
      console.log(`  📄 Page ${page}...`);

      const params = {
        search:   keyword,
        location: 'United States',
        radius_miles: 5000,
        page,
        jobs_per_page: CONFIG.jobsPerPage,
        api_key: API_KEY,
      };

      const res = await axios.get(API_URL, {
        params,
        headers: {
          'Accept': 'application/json',
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        },
        timeout: 30_000,
      });

      const data = res.data;
      const jobList = data.jobs || [];

      console.log(`    📦 Jobs returned: ${jobList.length}`);

      if (jobList.length === 0) {
        console.log(`    ⚠️  No more jobs – stopping pagination`);
        break;
      }

      for (const job of jobList) {
        const salary = extractSalary(job);
        if (!salary) continue; // skip jobs without salary

        jobs.push({
          Title:    job.name || 'N/A',
          Company:  job.hiring_company?.name || job.source || 'N/A',
          Salary:   salary,
          Location: [job.city, job.state].filter(Boolean).join(', ') || job.location || 'N/A',
          Type:     job.employment_type || '',
          Posted:   job.posted_time_friendly || job.date_posted || '',
          Link:     job.url || job.job_url || '',
          Keyword:  keyword,
        });
      }

      console.log(`    💰 With salary: ${jobs.length} total so far`);

      // Respect rate limits
      await sleep(2_000);

    } catch (err) {
      console.error(`  ❌ API error page ${page}:`, err.response?.status, err.message);
      // If 401/403, API key may be stale — stop trying
      if (err.response?.status === 401 || err.response?.status === 403) break;
      await sleep(5_000);
    }
  }

  return jobs;
}

// ====================== MAIN ======================

async function runScraper() {
  console.log('🚀 Starting ZipRecruiter Scraper (API mode)...');

  const allJobs = [];

  for (const keyword of KEYWORDS) {
    console.log(`\n🔍 Keyword: "${keyword}"`);
    const jobs = await fetchJobsForKeyword(keyword);
    allJobs.push(...jobs);
    console.log(`  ✅ "${keyword}" → ${jobs.length} jobs with salary`);
    await sleep(3_000);
  }

  // De-duplicate by Link
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
    await sendTelegramMessage(`✅ ZipRecruiter: Found <b>${uniqueJobs.length}</b> jobs with salary!\n${fileLink ?? ''}`);
    await sendTeamsAlert(uniqueJobs.length, fileLink);
    await sendTelegramFile(CONFIG.outputFile);
  } else {
    await sendTelegramMessage('❌ ZipRecruiter: No jobs with salary found.');
    await sendTeamsAlert(0);
    console.log('❌ No jobs with salary found.');
  }
}

runScraper().catch(console.error);