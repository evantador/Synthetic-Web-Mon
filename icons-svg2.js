const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const axios = require('axios');

(async () => {
    const url = 'https://www.bajajfinserv.in/';//'http://localhost:3000/icons.html'; // replace with your target URL
    const savePath = './screenshots/'; // folder to save screenshots
    const threshold = 1000; // threshold in milliseconds for slow loading
    const validIconsFilePath = './valid_icons.xlsx'; // path to store valid icons Excel sheet
    const slowIconsFilePath = './slow_icons.xlsx'; // path to store slow-loading icons Excel sheet
    const invalidIconsFilePath = './invalid_icons.xlsx'; // path to store invalid icons Excel sheet
    const teamsWebhookUrl = 'https://prod-05.westus.logic.azure.com:443/workflows/13167b57b66840b99fc8a954bbd04c35/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=oERk_ZOvXfBPXZlU_FZRSy7SXUbtYoMnJl5gEQtpvMo'; // replace with your Teams webhook URL

    // Create the screenshots folder if it doesn't exist
    if (!fs.existsSync(savePath)) {
        fs.mkdirSync(savePath, { recursive: true });
    }

    let validIcons = [];
    let slowLoadingIcons = [];
    let invalidIcons = [];
    let totalIcons = 0;
    let belowThreshold = 0;
    let aboveThreshold = 0;

    const browser = await puppeteer.launch();
    const page = await browser.newPage();

    await page.goto(url, { waitUntil: 'networkidle2' });

    // Find all SVG elements
    const svgElements = await page.$$('svg');

    if (svgElements.length === 0) {
        console.log('No SVG elements found.');
    } else {
        totalIcons = svgElements.length;
        console.log(`Found ${svgElements.length} SVG elements. Measuring load times and taking screenshots...`);

        for (let i = 0; i < svgElements.length; i++) {
            const element = svgElements[i];

            // Check if the SVG contains any child elements like path, rect, circle, or use
            const hasContent = await page.evaluate(el => {
                return el.querySelector('path, rect, circle, ellipse, line, polygon, polyline, use') !== null;
            }, element);

            if (!hasContent) {
                console.log(`Skipping SVG element ${i + 1}: No rendering content inside.`);
                
                // Get the XPath of the element
                const xpath = await page.evaluate(el => {
                    let getXPath = function(element) {
                        if (element.id !== '') {
                            return 'id("' + element.id + '")';
                        }
                        if (element === document.body) {
                            return element.tagName.toLowerCase();
                        }
                        let ix = 0;
                        let siblings = element.parentNode.childNodes;
                        for (let i = 0; i < siblings.length; i++) {
                            let sibling = siblings[i];
                            if (sibling === element) {
                                return getXPath(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix + 1) + ']';
                            }
                            if (sibling.nodeType === 1 && sibling.tagName === element.tagName) {
                                ix++;
                            }
                        }
                    };
                    return getXPath(el);
                }, element);

                // Add to invalid icons list
                invalidIcons.push({ index: i + 1, xpath });
                continue;
            }

            // Measure load time
            const startTime = new Date().getTime();
            await page.evaluate(el => el.complete, element);
            const endTime = new Date().getTime();
            const loadTime = endTime - startTime;

            // Get the bounding box of the element to ensure it has dimensions
            const boundingBox = await element.boundingBox();

            if (boundingBox && boundingBox.width > 0 && boundingBox.height > 0) {
                const screenshotPath = path.join(savePath, `svg-icon-${i + 1}.png`);
                
                // Take screenshot of the SVG icon
                await element.screenshot({ path: screenshotPath });

                console.log(`Screenshot saved: ${screenshotPath}`);

                // Get the XPath of the element
                const xpath = await page.evaluate(el => {
                    let getXPath = function(element) {
                        if (element.id !== '') {
                            return 'id("' + element.id + '")';
                        }
                        if (element === document.body) {
                            return element.tagName.toLowerCase();
                        }
                        let ix = 0;
                        let siblings = element.parentNode.childNodes;
                        for (let i = 0; i < siblings.length; i++) {
                            let sibling = siblings[i];
                            if (sibling === element) {
                                return getXPath(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix + 1) + ']';
                            }
                            if (sibling.nodeType === 1 && sibling.tagName === element.tagName) {
                                ix++;
                            }
                        }
                    };
                    return getXPath(el);
                }, element);

                // Append to valid icons
                validIcons.push({ index: i + 1, xpath, loadTime });

                // Categorize based on load time
                if (loadTime > threshold) {
                    aboveThreshold++;
                    slowLoadingIcons.push({ index: i + 1, xpath, loadTime });
                } else {
                    belowThreshold++;
                }

            } else {
                console.log(`Skipping SVG element ${i + 1}: Element has no width or height.`);

                // Get the XPath of the element
                const xpath = await page.evaluate(el => {
                    let getXPath = function(element) {
                        if (element.id !== '') {
                            return 'id("' + element.id + '")';
                        }
                        if (element === document.body) {
                            return element.tagName.toLowerCase();
                        }
                        let ix = 0;
                        let siblings = element.parentNode.childNodes;
                        for (let i = 0; i < siblings.length; i++) {
                            let sibling = siblings[i];
                            if (sibling === element) {
                                return getXPath(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix + 1) + ']';
                            }
                            if (sibling.nodeType === 1 && sibling.tagName === element.tagName) {
                                ix++;
                            }
                        }
                    };
                    return getXPath(el);
                }, element);

                // Add to invalid icons list
                invalidIcons.push({ index: i + 1, xpath });
            }
        }
    }

    await browser.close();

    // Write valid SVG icons to Excel
    if (validIcons.length > 0) {
        const wb = XLSX.utils.book_new();
        const wsData = [['Index', 'XPath', 'Load Time (ms)']];

        validIcons.forEach(icon => {
            wsData.push([icon.index, icon.xpath, icon.loadTime]);
        });

        const ws = XLSX.utils.aoa_to_sheet(wsData);
        XLSX.utils.book_append_sheet(wb, ws, 'Valid Icons');
        XLSX.writeFile(wb, validIconsFilePath);
    }

    // Write slow-loading SVG icons to Excel
    if (slowLoadingIcons.length > 0) {
        const wb = XLSX.utils.book_new();
        const wsData = [['Index', 'XPath', 'Load Time (ms)']];

        slowLoadingIcons.forEach(icon => {
            wsData.push([icon.index, icon.xpath, icon.loadTime]);
        });

        const ws = XLSX.utils.aoa_to_sheet(wsData);
        XLSX.utils.book_append_sheet(wb, ws, 'Slow Loading Icons');
        XLSX.writeFile(wb, slowIconsFilePath);

        // Send alert to Microsoft Teams
        let tableRows = slowLoadingIcons.map(icon => `
            <tr>
                <td>${icon.index}</td>
                <td>${icon.xpath}</td>
                <td>${icon.loadTime}ms</td>
            </tr>
        `).join('');

        let htmlTable = `
            <table border="1">
                <thead>
                    <tr>
                        <th>Index</th>
                        <th>XPath</th>
                        <th>Load Time</th>
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
        await axios.post(teamsWebhookUrl, message,{
            headers: {
            'Content-Type': 'application/json',
            'User-Agent': 'axios/0.21.1'
            }
        });
    }

    // Write invalid SVG icons to Excel
    if (invalidIcons.length > 0) {
        const wb = XLSX.utils.book_new();
        const wsData = [['Index', 'XPath']];

        invalidIcons.forEach(icon => {
            wsData.push([icon.index, icon.xpath]);
        });

        const ws = XLSX.utils.aoa_to_sheet(wsData);
        XLSX.utils.book_append_sheet(wb, ws, 'Invalid Icons');
        XLSX.writeFile(wb, invalidIconsFilePath);

        // Send alert to Microsoft Teams
        let tableRows = invalidIcons.map(icon => `
            <tr>
                <td>${icon.index}</td>
                <td>${icon.xpath}</td>
            </tr>
        `).join('');

        let htmlTable = `
            <table border="1">
                <thead>
                    <tr>
                        <th>Index</th>
                        <th>XPath</th>
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
        await axios.post(teamsWebhookUrl, message,{
            headers: {
            'Content-Type': 'application/json',
            'User-Agent': 'axios/0.21.1'
            }
        });
    }

    // Send summary report to Microsoft Teams
    let summaryTable = `
        <table border="1">
            <thead>
                <tr>
                    <th>Total Icons</th>
                    <th>Valid Icons</th>
                    <th>Invalid Icons</th>
                    <th>Icons Below Threshold</th>
                    <th>Icons Above Threshold</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>${totalIcons}</td>
                    <td>${validIcons.length}</td>
                    <td>${invalidIcons.length}</td>
                    <td>${belowThreshold}</td>
                    <td>${aboveThreshold}</td>
                </tr>
            </tbody>
        </table>
    `;

    const message = {"type": "message",
        "attachments": [{
        "contentType": "text",
        "content": summaryTable
    }]};

    // Send the alert to the Teams channel
    await axios.post(teamsWebhookUrl, message,{
        headers: {
        'Content-Type': 'application/json',
        'User-Agent': 'axios/0.21.1'
        }
    });

    console.log('Summary sent to Microsoft Teams.');
})();
