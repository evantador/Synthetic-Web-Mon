const puppeteer = require('puppeteer');
const fs = require('fs');
const axios = require('axios');
const xlsx = require('xlsx');

// Threshold for image load time in milliseconds
const loadTimeThreshold = 3000;

async function checkImages(url) {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto(url);

    const images = await page.evaluate(() => {
        const imgElements = Array.from(document.querySelectorAll('img, svg, [src]'));
        return imgElements.map(img => ({
            src: img.src || img.getAttribute('src'),
            xpath: getXPath(img)
        }));

        function getXPath(element) {
            let xpath = '';
            for (; element && element.nodeType == 1; element = element.parentNode) {
                let id = Array.from(element.parentNode.children).indexOf(element) + 1;
                xpath = '/' + element.tagName.toLowerCase() + '[' + id + ']' + xpath;
            }
            return xpath;
        }
    });

    const invalidImages = [];
    const slowLoadingImages = [];

    for (const image of images) {
        try {
            const start = Date.now();
            const response = await axios.get(image.src);
            const loadTime = Date.now() - start;

            if (response.status >= 400) {
                invalidImages.push({ src: image.src, xpath: image.xpath });
            } else if (loadTime > loadTimeThreshold) {
                slowLoadingImages.push({ src: image.src, loadTime, xpath: image.xpath });
            }
        } catch (error) {
            console.log(`Error fetching image ${image.src}:`, error.message);
            invalidImages.push({ src: image.src, xpath: image.xpath });
        }
    }

    await browser.close();

    if (invalidImages.length > 0) {
        await appendToExcel(invalidImages, 'invalid_images.xlsx', ['Image Link', 'XPath']);
        await sendAlertToTeams(invalidImages, 'Red Alert: Invalid Images Detected!');
    }

    if (slowLoadingImages.length > 0) {
        await appendToExcel(slowLoadingImages, 'slow_loading_images.xlsx', ['Source Link', 'Load Time (ms)', 'XPath']);
        await sendAlertToTeams(slowLoadingImages, 'Yellow Alert: Slow Loading Images Detected!');
    }
}

async function appendToExcel(data, filePath, headers) {
    try {
        let workbook;
        if (fs.existsSync(filePath)) {
            workbook = xlsx.readFile(filePath);
        } else {
            workbook = xlsx.utils.book_new();
        }

        const sheetName = 'Sheet1';
        let worksheet = workbook.Sheets[sheetName];

        if (!worksheet) {
            worksheet = xlsx.utils.aoa_to_sheet([headers]);
            xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
        }

        data.forEach(item => {
            const newRow = Object.values(item);
            xlsx.utils.sheet_add_aoa(worksheet, [newRow], { origin: -1 });
        });

        xlsx.writeFile(workbook, filePath);
        console.log(`Data appended to ${filePath}`);
    } catch (error) {
        console.log(`Error appending to Excel ${filePath}:`, error.message);
    }
}

async function sendAlertToTeams(data, alertTitle) {
    // Webhook URL for Microsoft Teams
    const webhookUrl = 'https://prod-05.westus.logic.azure.com:443/workflows/13167b57b66840b99fc8a954bbd04c35/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=oERk_ZOvXfBPXZlU_FZRSy7SXUbtYoMnJl5gEQtpvMo';

    try {
        let htmlTableRows = data.map(item => {
            return `<tr>${Object.values(item).map(val => `<td>${val}</td>`).join('')}</tr>`;
        }).join('');

        let htmlMessage = `
            <h1 style="color: red;">${alertTitle}</h1>
            <table border="1" style="width:100%; border-collapse: collapse;">
                <thead>
                    <tr>${Object.keys(data[0]).map(key => `<th>${key}</th>`).join('')}</tr>
                </thead>
                <tbody>
                    ${htmlTableRows}
                </tbody>
            </table>
        `;

        const message = {"type": "message",
            "attachments": [{
            "contentType": "text",
            "content": htmlMessage
        }]};
    
        
        // Send the alert to the Teams channel
        await axios.post(webhookUrl, message,{
            headers: {
            'Content-Type': 'application/json',
            'User-Agent': 'axios/0.21.1'
            }
        });

        console.log(`Alert sent to Teams: ${alertTitle}`);
    } catch (error) {
        console.log(`Error sending alert to Teams:`, error.message);
    }
}

// Example usage
checkImages('https://www.marutisuzukitruevalue.com/').then(() => {
    console.log('Image validation completed.');
}).catch(error => {
    console.error('Error during image validation:', error);
});
