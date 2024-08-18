const puppeteer = require('puppeteer');
const fs = require('fs');
const xlsx = require('xlsx');

(async () => {
    // Launch the browser and open a new page
    const browser = await puppeteer.launch();
    const page = await browser.newPage();

    // Go to the specified webpage
    const url = 'http://localhost:3000'; // Replace with your target URL
    await page.goto(url, { waitUntil: 'domcontentloaded' });

    // Extract all text content and their corresponding XPaths, ensuring the correct hierarchy
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

    // Save the extracted text to a text file
    const filePath = 'webpage_text.txt';
    const textContent = extractedData.map(item => item.text).join('\n\n');
    fs.writeFileSync(filePath, textContent);
    console.log(`Text extracted and saved to ${filePath}`);

    // Save the extracted text and XPaths to an Excel file
    const excelPath = `WebText.xlsx`;
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.json_to_sheet(extractedData);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Text and XPath');
    xlsx.writeFile(workbook, excelPath);
    console.log(`Text and XPath data saved to ${excelPath}`);

    // Close the browser
    await browser.close();
})();
