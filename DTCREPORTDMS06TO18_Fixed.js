const fs = require('fs');
const path = require('path');

// Freeze the Bangkok report window before any GPS work.
const BANGKOK_OFFSET_MS = 7 * 60 * 60 * 1000;
const toBangkokClock = date => new Date(date.getTime() + BANGKOK_OFFSET_MS);

function validateReportDate(value) {
    const parsed = new Date(`${value}T00:00:00.000Z`);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value) || !Number.isFinite(parsed.getTime()) ||
        parsed.toISOString().slice(0, 10) !== value) {
        throw new Error('REPORT_DATE must be a valid YYYY-MM-DD date.');
    }
    return value;
}

async function getScheduledRunCreatedAt(env, fetchRun) {
    const repo = env.GITHUB_REPOSITORY || '';
    const runId = env.GITHUB_RUN_ID || '';
    const token = (env.GH_TOKEN || env.GITHUB_TOKEN || '').trim();
    if (!/^[\w.-]+\/[\w.-]+$/.test(repo) || !/^\d+$/.test(runId) || !token) {
        throw new Error('GitHub run metadata or GH_TOKEN missing; check workflow permissions (actions: read).');
    }
    const url = new URL(`${(env.GITHUB_API_URL || 'https://api.github.com').replace(/\/+$/, '')}/repos/${repo}/actions/runs/${runId}`);
    if (url.protocol !== 'https:' || url.username || url.password) throw new Error('Invalid GitHub API URL.');
    const response = await fetchRun(url, {
        headers: { Accept: 'application/vnd.github+json', Authorization: `Bearer ${token}` },
        redirect: 'error',
        signal: AbortSignal.timeout(15000),
    });
    if (!response.ok) throw new Error(`GitHub API HTTP ${response.status}; check actions: read or run manually with report_date.`);
    const run = await response.json();
    if (String(run.id) !== runId || run.event !== 'schedule' || typeof run.created_at !== 'string' ||
        !run.created_at.endsWith('Z') || !Number.isFinite(Date.parse(run.created_at))) {
        throw new Error('Invalid GitHub scheduled run metadata; run manually with report_date.');
    }
    return new Date(run.created_at);
}

async function resolveReportWindow({ startedAt = new Date(), env = process.env, fetchRun = globalThis.fetch } = {}) {
    const explicit = (env.REPORT_DATE || '').trim();
    let reportDate;
    let source;
    let createdAt;
    if (explicit) {
        reportDate = validateReportDate(explicit);
        source = 'REPORT_DATE';
    } else {
        if (env.GITHUB_ACTIONS === 'true') {
            if (env.GITHUB_EVENT_NAME !== 'schedule') {
                throw new Error('Manual and other GitHub triggers must supply REPORT_DATE (YYYY-MM-DD).');
            }
            createdAt = await getScheduledRunCreatedAt(env, fetchRun);
            source = 'github-run-created-at';
        } else {
            source = 'local-06-cutoff';
        }
        const clock = toBangkokClock(createdAt || startedAt);
        const hour = clock.getUTCHours();
        if (createdAt && hour >= 6 && hour < 18) {
            throw new Error('GitHub created the run between 06:00 and 17:59 Bangkok time; run manually with report_date to avoid selecting the wrong shift.');
        }
        if (hour < 6) clock.setUTCDate(clock.getUTCDate() - 1);
        reportDate = clock.toISOString().slice(0, 10);
    }
    if (startedAt.getTime() < Date.parse(`${reportDate}T18:00:00+07:00`)) {
        throw new Error(`The shift for ${reportDate} has not ended (18:00 Bangkok time).`);
    }
    return Object.freeze({ reportDate, source, createdAt, startDateTime: `${reportDate} 06:00`, endDateTime: `${reportDate} 18:00` });
}

function logReportWindow(window, startedAt) {
    const thai = date => `${toBangkokClock(date).toISOString().slice(0, 19)}+07:00`;
    console.log(`[Date] Script started: ${thai(startedAt)}`);
    if (window.createdAt) console.log(`[Date] GitHub run created: ${thai(window.createdAt)} (not nominal cron time)`);
    console.log(`[Date] Report window: ${window.startDateTime} to ${window.endDateTime} (Asia/Bangkok; source: ${window.source})`);
}

// ฟังก์ชันแปลง HTML เป็น Excel (.xlsx) แบบจัดเต็ม
async function convertHtmlToExcel(sourcePath, destPath) {
    try {
        const ExcelJS = require('exceljs');
        const { JSDOM } = require('jsdom');
        console.log(`   Converting HTML-XLS to Beautiful XLSX (ExcelJS)...`);
        
        const htmlContent = fs.readFileSync(sourcePath, 'utf-8');
        const dom = new JSDOM(htmlContent);
        const table = dom.window.document.querySelector('table');
        
        if (!table) throw new Error('No table found in downloaded file');

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('DMS Report');

        const rows = Array.from(table.querySelectorAll('tr'));
        
        rows.forEach((row, rowIndex) => {
            const cells = Array.from(row.querySelectorAll('td, th'));
            const rowData = cells.map(cell => {
                // ลบ HTML tags และช่องว่างส่วนเกิน
                return cell.textContent.replace(/<[^>]*>/g, '').trim();
            });
            
            const excelRow = worksheet.addRow(rowData);

            // จัดรูปแบบ (Styling)
            excelRow.eachCell((cell, colNumber) => {
                // เส้นขอบ (Borders)
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                
                // การจัดวาง (Alignment)
                cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

                // ถ้าเป็น Header
                if (rowIndex === 0 || cells[colNumber-1].tagName === 'TH') {
                    cell.font = { bold: true };
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFD3D3D3' } // สีเทาอ่อน
                    };
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                }
            });
        });

        // Auto-fit Column Width
        worksheet.columns.forEach(column => {
            let maxLength = 0;
            if (column.values) {
                column.values.forEach(val => {
                    const len = val ? String(val).length : 0;
                    if (len > maxLength) maxLength = len;
                });
            }
            column.width = Math.min(Math.max(maxLength + 2, 10), 50);
        });

        await workbook.xlsx.writeFile(destPath);
        console.log(`   Conversion success: ${destPath}`);
        return true;

    } catch (e) {
        console.error(`   Conversion failed: ${e.message}`);
        return false;
    }
}

(async () => {
    const startedAt = new Date();
    let reportWindow;
    try {
        reportWindow = await resolveReportWindow({ startedAt });
    } catch (error) {
        console.error(`[Date] Stopped before GPS login: ${error.message}`);
        process.exitCode = 1;
        return;
    }
    logReportWindow(reportWindow, startedAt);
    if (process.argv.includes('--check-report-date')) return;
    const { startDateTime, endDateTime } = reportWindow;
    const puppeteer = require('puppeteer');
    const nodemailer = require('nodemailer');

    // --- รับค่าจาก Secrets ---
    const { DTC_USERNAME, DTC_PASSWORD, EMAIL_USER, EMAIL_PASS, EMAIL_TO } = process.env;

    if (!DTC_USERNAME || !DTC_PASSWORD || !EMAIL_USER || !EMAIL_PASS || !EMAIL_TO) {
        console.error('Error: Missing required secrets.');
        process.exit(1);
    }

    console.log('Launching browser...');
    const downloadPath = path.resolve('./downloads');
    
    // เคลียร์โฟลเดอร์เก่า (Force Clean)
    if (fs.existsSync(downloadPath)) {
        try { fs.rmSync(downloadPath, { recursive: true, force: true }); } catch(e) {}
    }
    fs.mkdirSync(downloadPath);

    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
    });
    
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(300000);
    page.setDefaultTimeout(300000);
    await page.emulateTimezone('Asia/Bangkok');

    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });
    await page.setViewport({ width: 1920, height: 1080 });

    try {
        // Step 1: Login
        console.log('1️⃣ Step 1: Login...');
        await page.goto('https://gps.dtc.co.th/ultimate/index.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#txtname', { visible: true, timeout: 90000 });
        await page.type('#txtname', DTC_USERNAME);
        await page.type('#txtpass', DTC_PASSWORD);
        await Promise.all([
            page.evaluate(() => document.getElementById('btnLogin').click()),
            page.waitForFunction(() => !document.querySelector('#txtname'), { timeout: 90000 })
        ]);
        console.log('✅ Login Success');

        // Step 2: Navigate
        console.log('2️⃣ Step 2: Go to Report Page...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_other_status.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true, timeout: 60000 });

        // Step 2.5: Select Truck "ทั้งหมด"
        console.log('   Selecting Truck "ทั้งหมด"...');
        await page.waitForSelector('#ddl_truck', { visible: true });
        await page.waitForFunction(() => document.getElementById('ddl_truck').options.length > 0);
        await page.evaluate(() => {
            const select = document.getElementById('ddl_truck');
            for (let opt of select.options) {
                if (opt.text.includes('ทั้งหมด') || opt.text.toLowerCase().includes('all')) {
                    select.value = opt.value;
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    break;
                }
            }
        });

        // Step 2.6: Select Report Types
        console.log('   Selecting 3 Report Types...');
        try {
            await page.waitForSelector('#s2id_ddlharsh', { visible: true });
            const keywords = ["ความง่วงระดับ 1", "ความง่วงระดับ 2", "หาว"];
            for (const kw of keywords) {
                await page.click('#s2id_ddlharsh'); 
                await new Promise(r => setTimeout(r, 500));
                const input = await page.$('#s2id_ddlharsh input') || await page.$('.select2-input');
                if (input) {
                    await input.type(kw);
                    await new Promise(r => setTimeout(r, 1000));
                    await page.keyboard.press('Enter');
                    console.log(`      Selected: "${kw}"`);
                    await new Promise(r => setTimeout(r, 500));
                }
            }
        } catch (e) { console.log('❌ Error selecting reports:', e.message); }

        // Step 3: Apply the report window selected before login.
        console.log('3️⃣ Step 3: Setting Date Range 06:00 - 18:00...');
        console.log(`      Range: ${startDateTime} to ${endDateTime}`);

        await page.evaluate(() => document.getElementById('date9').value = '');
        await page.type('#date9', startDateTime);
        await page.evaluate(() => document.getElementById('date10').value = '');
        await page.type('#date10', endDateTime);

        console.log('   Clicking Search & Waiting 120s...');
        await page.click('td:nth-of-type(5) > span');
        await new Promise(r => setTimeout(r, 120000)); 

        // Step 4: Export
        console.log('4️⃣ Step 4: Clicking Export/Excel...');
        
        // Force Clean again before download
        try {
            if (fs.existsSync(downloadPath)) {
                fs.rmSync(downloadPath, { recursive: true, force: true });
                fs.mkdirSync(downloadPath);
            }
        } catch(e) {}

        await page.waitForSelector('#btnexport', { visible: true });
        await page.evaluate(() => document.querySelector('#btnexport').click());
        console.log('   Waiting for download (30s)...');
        await new Promise(r => setTimeout(r, 30000));

        // Step 5: Convert & Email
        console.log('5️⃣ Step 5: Preparing Email...');
        const files = fs.readdirSync(downloadPath).filter(f => !f.startsWith('.'));
        if (files.length > 0) {
            const recentFile = files.map(f => ({ 
                name: f, 
                time: fs.statSync(path.join(downloadPath, f)).mtime.getTime() 
            })).sort((a, b) => b.time - a.time)[0];

            const originalPath = path.join(downloadPath, recentFile.name);
            const newFileName = recentFile.name.replace(/\.xls$/, '') + '.xlsx';
            const newFilePath = path.join(downloadPath, newFileName);
            
            // แปลงไฟล์เป็น XLSX ด้วย ExcelJS
            const converted = await convertHtmlToExcel(originalPath, newFilePath);
            
            const fileToSend = converted ? newFilePath : originalPath;
            const nameToSend = converted ? newFileName : recentFile.name;
            const subject = `${recentFile.name} รอบ 06:00 ถึง 18:00`;

            const transporter = nodemailer.createTransport({
                service: 'gmail',
                auth: { user: EMAIL_USER, pass: EMAIL_PASS }
            });

            console.log(`   Sending email to: ${EMAIL_TO}`);
            await transporter.sendMail({
                from: `"DTC DMS Reporter" <${EMAIL_USER}>`,
                to: EMAIL_TO,
                subject: subject,
                text: 'ถึง ผู้เกี่ยวข้อง\nรายงาน DMS กะเช้า (06:00 - 18:00)\nด้วยความนับถือ\nBot Report',
                attachments: [{ filename: nameToSend, path: fileToSend }]
            });
            console.log('📧 Email Sent!');
            
            try { fs.unlinkSync(originalPath); } catch(e){}
            if(converted) try { fs.unlinkSync(newFilePath); } catch(e){}
        } else {
            console.error('❌ No file found!');
            process.exit(1);
        }

    } catch (err) {
        console.error('❌ Error:', err);
        await page.screenshot({ path: 'error.png' });
        process.exit(1);
    } finally {
        await browser.close();
    }
})();
