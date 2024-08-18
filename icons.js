const puppeteer = require('puppeteer');
const xlsx = require('xlsx');
const axios = require('axios');

const THRESHOLD = 1000; // Threshold time for loading icons in ms

(async () => {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const validExcelName = `ValidIcons_${timestamp}.xlsx`;
    const slowIconsExcelName = `SlowIcons_${timestamp}.xlsx`;
    const invalidIconsExcelName = `InvalidIcons_${timestamp}.xlsx`;

    // Initialize data arrays
    const validIcons = [];
    const slowIcons = [];
    const invalidIcons = [];

    // Launch the browser
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    const url = 'http://localhost:3000/icons.html'; // Replace with your target URL
    await page.goto(url, { waitUntil: 'domcontentloaded' });

    // Validate and measure load times of different icons
    await validateBootstrapIcons(page);
    await validateGoogleMaterialIcons(page);
    await validateSVGIcons(page);

    // Write results to Excel files
    appendToExcel(validIcons, validExcelName);
    appendToExcel(slowIcons, slowIconsExcelName);
    appendToExcel(invalidIcons, invalidIconsExcelName);

    // Send alerts for slow and invalid icons
    await sendAlertToTeams('Slow Icons Alert', slowIcons);
    await sendAlertToTeams('Invalid Icons Alert', invalidIcons);

    // Send summary alert
    const summaryMessage = `
        Total Number of Icons: ${validIcons.length + slowIcons.length + invalidIcons.length}
        Total Loaded Icons: ${validIcons.length}
        Number of Icons Didn't Load in Threshold Time: ${slowIcons.length}
        Number of Invalid Icons: ${invalidIcons.length}
        Number of Icons Didn't Load: ${invalidIcons.length}
    `;
    await sendSummaryAlertToTeams('Icon Load Summary', summaryMessage);

    // Close the browser
    await browser.close();

    // Function to validate and measure load time of Bootstrap icons
    async function validateBootstrapIcons(page) {
        const icons = await page.$$('[class*="bi-"]');
        for (let icon of icons) {
            await validateIcon(icon, 'Bootstrap');
        }
    }

    // Function to validate and measure load time of Google Material icons
    async function validateGoogleMaterialIcons(page) {
        const icons = await page.$$('[class*="material-icons"]');
        for (let icon of icons) {
            await validateIcon(icon, 'Google Material');
        }
    }

    // Function to validate and measure load time of SVG tag icons
    async function validateSVGIcons(page) {
        const icons = await page.$$('svg');
        for (let icon of icons) {
            await validateIcon(icon, 'SVG');
        }
    }

    // Function to validate individual icons and measure load time
    async function validateIcon(icon, iconType) {
        const xpath = await page.evaluate(el => {
            let xpath = '';
            for (; el && el.nodeType == 1; el = el.parentNode) {
                let index = Array.from(el.parentNode.children).indexOf(el) + 1;
                xpath = '/' + el.tagName.toLowerCase() + '[' + index + ']' + xpath;
            }
            return xpath;
        }, icon);

        try {
            const startTime = Date.now();
            await icon.evaluate(node => node.complete);
            const loadTime = Date.now() - startTime;

            if (loadTime > THRESHOLD) {
                slowIcons.push({ type: iconType, xpath, loadTime });
            } else {
                validIcons.push({ type: iconType, xpath, loadTime });
            }
        } catch (error) {
            invalidIcons.push({ type: iconType, xpath, loadTime: 'N/A', error: error.message });
        }
    }

    // Function to append results to an Excel file
    function appendToExcel(dataArray, fileName) {
        const workbook = xlsx.utils.book_new();
        const worksheet = xlsx.utils.json_to_sheet(dataArray);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Icons');
        xlsx.writeFile(workbook, fileName);
        console.log(`${fileName} saved successfully.`);
    }

    // Function to send alerts to Microsoft Teams using axios
    async function sendAlertToTeams(alertTitle, data) {
        if (data.length === 0) return; // No need to send alert if data is empty

        try {
            const webhookUrl = 'https://prod-05.westus.logic.azure.com:443/workflows/13167b57b66840b99fc8a954bbd04c35/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=oERk_ZOvXfBPXZlU_FZRSy7SXUbtYoMnJl5gEQtpvMo'; // Replace with your webhook URL
            const tableRows = data.map(item =>
                `<tr><td>${item.type}</td><td>${item.xpath}</td><td>${item.loadTime}</td></tr>`
            ).join('');

            const htmlTable = `
                <table border="1" style="border-collapse: collapse;">
                    <thead>
                        <tr>
                            <th>Type</th>
                            <th>XPath</th>
                            <th>Load Time (ms)</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${tableRows}
                    </tbody>
                </table>
            `;

            const message = {"type": "message",
                "attachments": [{
                "contentType": "text",
                "content": htmlTable
            }]};
    
            // Send the alert to the Teams channel
            await axios.post(webhookUrl, message,{
                headers: {
                'Content-Type': 'application/json',
                'User-Agent': 'axios/0.21.1'
                }
            });        

            console.log(`${alertTitle} sent to Teams`);
        } catch (error) {
            console.error(`Failed to send ${alertTitle} to Teams:`, error);
        }
    }

    // Function to send summary alert to Microsoft Teams
    async function sendSummaryAlertToTeams(alertTitle, summaryMessage) {
        try {
            const webhookUrl = 'https://prod-05.westus.logic.azure.com:443/workflows/13167b57b66840b99fc8a954bbd04c35/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=oERk_ZOvXfBPXZlU_FZRSy7SXUbtYoMnJl5gEQtpvMo'; // Replace with your webhook URL
            const message = {"type": "message",
                "attachments": [{
                "contentType": "text",
                "content": summaryMessage
            }]};
    
            // Send the alert to the Teams channel
            await axios.post(webhookUrl, message,{
                headers: {
                'Content-Type': 'application/json',
                'User-Agent': 'axios/0.21.1'
                }
            });        
    
            console.log(`${alertTitle} sent to Teams`);
        } catch (error) {
            console.error(`Failed to send summary alert to Teams:`, error);
        }
    }
})();
