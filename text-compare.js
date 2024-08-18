const puppeteer = require('puppeteer');
const fs = require('fs');
const xlsx = require('xlsx');
const axios = require('axios');

(async () => {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const newExcelName = `WebText_${timestamp}.xlsx`;

    // Launch the browser and open a new page
    const browser = await puppeteer.launch();
    const page = await browser.newPage();

    // Go to the specified webpage
    const url = 'http://localhost:3000/'; // Replace with your target URL
    await page.goto(url, { waitUntil: 'domcontentloaded' });

    // Extract all text content and their corresponding XPaths, including nested elements
    const extractedData = await page.evaluate(() => {
        function getXPath(element) {
            let xpath = '';
            for (; element && element.nodeType === 1; element = element.parentNode) {
                let id = Array.from(element.parentNode.children).indexOf(element) + 1;
                xpath = '/' + element.tagName.toLowerCase() + '[' + id + ']' + xpath;
            }
            return xpath;
        }

        const textNodes = [];
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);

        while (walker.nextNode()) {
            const node = walker.currentNode;
            const text = node.nodeValue.trim();

            if (text.length > 0) {
                const parentElement = node.parentElement;
                textNodes.push({
                    text: text,
                    xpath: getXPath(parentElement)
                });
            }
        }

        return textNodes;
    });

    // Save the extracted text and XPaths to a new Excel file
    const workbookNew = xlsx.utils.book_new();
    const worksheetNew = xlsx.utils.json_to_sheet(extractedData);
    xlsx.utils.book_append_sheet(workbookNew, worksheetNew, 'Text and XPath');
    xlsx.writeFile(workbookNew, newExcelName);
    console.log(`New text and XPath data saved to ${newExcelName}`);

    // Load the original Excel file
    const originalWorkbook = xlsx.readFile('WebText.xlsx');
    const originalSheet = originalWorkbook.Sheets['Text and XPath'];
    const originalData = xlsx.utils.sheet_to_json(originalSheet);

    const dataAccuracyErrors = [];
    const missingElements = [];

    const extractedDataMap = new Map(extractedData.map(item => [item.xpath, item.text]));

    // Compare original data with the newly extracted data
    originalData.forEach(item => {
        const { text: originalText, xpath } = item;
        const currentText = extractedDataMap.get(xpath);

        if (currentText === undefined) {
            // If the XPath does not exist in the newly extracted data
            missingElements.push({ text: originalText, xpath });
        } else if (originalText !== currentText) {
            // If the text differs
            dataAccuracyErrors.push({
                originalText,
                currentText,
                xpath
            });
        }
    });

    // Save Data Accuracy Errors
    if (dataAccuracyErrors.length > 0) {
        const workbookErrors = xlsx.utils.book_new();
        const worksheetErrors = xlsx.utils.json_to_sheet(dataAccuracyErrors);
        xlsx.utils.book_append_sheet(workbookErrors, worksheetErrors, 'Data Accuracy Errors');
        xlsx.writeFile(workbookErrors, 'Data_accuracy_errors.xlsx');
        console.log('Data accuracy errors saved to Data_accuracy_errors.xlsx');

        // Send Teams Alert for Data Accuracy Errors
        await sendAlertToTeams('Data Accuracy Errors', dataAccuracyErrors);
    }

    // Save Missing Elements
    if (missingElements.length > 0) {
        const workbookMissing = xlsx.utils.book_new();
        const worksheetMissing = xlsx.utils.json_to_sheet(missingElements);
        xlsx.utils.book_append_sheet(workbookMissing, worksheetMissing, 'Missing Elements');
        xlsx.writeFile(workbookMissing, 'Missing_Elements.xlsx');
        console.log('Missing elements saved to Missing_Elements.xlsx');

        // Send Teams Alert for Missing Elements
        await sendAlertToTeams('Missing Elements', missingElements, true);
    }

    // Close the browser
    await browser.close();

    // Function to send an alert to a Microsoft Teams channel using axios
    async function sendAlertToTeams(alertType, data, isMissingElements = false) {
        try {
            const webhookUrl = 'https://prod-05.westus.logic.azure.com:443/workflows/13167b57b66840b99fc8a954bbd04c35/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=oERk_ZOvXfBPXZlU_FZRSy7SXUbtYoMnJl5gEQtpvMo'; // Replace with your Teams webhook URL

            if (isMissingElements) {
                // Format for missing elements alert in tabular format
                const tableRows = data.map(item => 
                    `<tr><td>${item.text || ''}</td><td>${item.xpath}</td></tr>`
                ).join('');

                const htmlTable = `
                    <table border="1" style="border-collapse: collapse;">
                        <thead>
                            <tr>
                                <th>Text</th>
                                <th>XPath</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${tableRows}
                        </tbody>
                    </table>
                `;

                const message = {
                    type: "message",
                    attachments: [{
                        contentType: "text/html",
                        content: `<h3>${alertType}</h3>${htmlTable}`
                    }]
                };

                await axios.post(webhookUrl, message, {
                    headers: {
                        'Content-Type': 'application/json',
                        'User-Agent': 'axios/0.21.1'
                    }
                });
                console.log(`${alertType} alert sent to Teams`);
            } else {
                // Format for data accuracy errors alert in tabular format
                const tableRows = data.map(item => 
                    `<tr><td>${item.originalText || ''}</td><td>${item.currentText || ''}</td><td>${item.xpath}</td></tr>`
                ).join('');

                const htmlTable = `
                    <table border="1" style="border-collapse: collapse;">
                        <thead>
                            <tr>
                                <th>Original Text</th>
                                <th>Current Text</th>
                                <th>XPath</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${tableRows}
                        </tbody>
                    </table>
                `;

                const message = {
                    type: "message",
                    attachments: [{
                        contentType: "text/html",
                        content: `<h3>${alertType}</h3>${htmlTable}`
                    }]
                };

                await axios.post(webhookUrl, message, {
                    headers: {
                        'Content-Type': 'application/json',
                        'User-Agent': 'axios/0.21.1'
                    }
                });
                console.log(`${alertType} alert sent to Teams`);
            }
        } catch (error) {
            console.error(`Failed to send ${alertType} alert to Teams:`, error);
        }
    }
})();
