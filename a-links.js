const puppeteer = require('puppeteer');
const fs = require('fs');
const axios = require('axios');
const xlsx = require('xlsx');

function getISTTimestamp() {
    const date = new Date();
    const offset = 5.5 * 60; // IST is UTC+5:30
    const utc = date.getTime() + (date.getTimezoneOffset() * 60000);
    const istDate = new Date(utc + (offset * 60000));
    return istDate.toISOString().replace('T', ' ').slice(0, -1); // Format: YYYY-MM-DD HH:MM:SS
}

async function checkLinks(url) {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto(url);

    const links = await page.evaluate(() => {
        return Array.from(document.querySelectorAll('a')).map(a => a.href);
    });

    const brokenLinks = [];
    const timestamp = getISTTimestamp(); // IST Timestamp
    
    for (const link of links) {
        try {
            const response = await axios.get(link);
            if (response.status >= 400) {
                brokenLinks.push({ link, status: response.status, timestamp });
            }
        } catch (error) {
            brokenLinks.push({ link, status: error.response ? error.response.status : 'Network Error', timestamp });
        }
    }

    await browser.close();

    if (brokenLinks.length > 0) {
        await appendToExcel(brokenLinks);
        await sendAlertToTeams(brokenLinks, links.length);
    }

    return brokenLinks;
}

async function appendToExcel(brokenLinks) {
    const filePath = 'broken_links.xlsx';
    let workbook;
    
    // Check if the Excel file exists
    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath);
    } else {
        workbook = xlsx.utils.book_new();
    }

    const sheetName = 'Broken Links';
    let worksheet = workbook.Sheets[sheetName];

    if (!worksheet) {
        worksheet = xlsx.utils.aoa_to_sheet([['Link', 'HTTP Status Code', 'Timestamp']]);
        xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
    }

    brokenLinks.forEach(brokenLink => {
        const newRow = [brokenLink.link, brokenLink.status, brokenLink.timestamp];
        xlsx.utils.sheet_add_aoa(worksheet, [newRow], { origin: -1 });
    });
    try {
        xlsx.writeFile(workbook, filePath);
        console.log('Updated Excel!');
    } catch (error) {
        console.log('Error Updating Excel Sheet!', error);
    }
}

async function sendAlertToTeams(brokenLinks, totalLinks) {
    if (brokenLinks.length === 0) return;
    
    // HTML table for broken links
    let htmlTableRows = brokenLinks.map(link => {
        return `<tr><td>${link.link}</td><td>${link.status}</td><td>${link.timestamp}</td></tr>`;
    }).join('');

    // HTML summary
    const totalBrokenLinks = brokenLinks.length;
    const totalLinksCount = totalLinks;
    
    const htmlSummary = `
        <h1 style="color: red;">Red Alert: Broken Links Detected!</h1>
        <h2>Summary</h2>
        <table border="1" style="width:100%; border-collapse: collapse;">
            <thead>
                <tr>
                    <th>Total Links</th>
                    <th>Total Broken Links</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>${totalLinksCount}</td>
                    <td>${totalBrokenLinks}</td>
                </tr>
            </tbody>
        </table>
        <h2>Broken Links Details</h2>
        <table border="1" style="width:100%; border-collapse: collapse;">
            <thead>
                <tr>
                    <th>Link</th>
                    <th>HTTP Status Code</th>
                    <th>Timestamp (IST)</th>
                </tr>
            </thead>
            <tbody>
                ${htmlTableRows}
            </tbody>
        </table>
    `;

    // Webhook URL for Microsoft Teams
    const webhookUrl = 'https://prod-05.westus.logic.azure.com:443/workflows/13167b57b66840b99fc8a954bbd04c35/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=oERk_ZOvXfBPXZlU_FZRSy7SXUbtYoMnJl5gEQtpvMo';

    const message = {
        "type": "message",
        "attachments": [{
            "contentType": "text",
            "content": htmlSummary
        }]
    };

    try {
        // Send the alert to the Teams channel
        await axios.post(webhookUrl, message, {
            headers: {
                'Content-Type': 'application/json',
                'User-Agent': 'axios/0.21.1'
            }
        });

        console.log('Sent Alert to Teams!');
    } catch (error) {
        console.log('Error Sending Alert in Teams!', error);
    }
}

// Example usage
checkLinks('http://localhost:3000/').then(brokenLinks => {
    console.log('Broken Links:', brokenLinks);
}).catch(error => {
    console.error('Error:', error);
});
