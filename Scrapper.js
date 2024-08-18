const puppeteer = require('puppeteer');

(async () => {
  // Launch a new browser instance
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();

  // Navigate to a website
  await page.goto('http://localhost:3000/');

  // Take a screenshot
  await page.screenshot({ path: 'screenshot.png' });

  // Get the page title
  const title = await page.title();
  console.log('Page title:', title);

  // Close the browser
  await browser.close();
})();
