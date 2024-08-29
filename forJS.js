const ExcelJS = require('exceljs');
const { Builder, By, Key } = require('selenium-webdriver');

(async function googleSearchAutomation() {
    // Setup Selenium WebDriver for Chrome
    const driver = await new Builder().forBrowser('chrome').build();

    try {
        // Load the Excel workbook
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('4BeatsQ1.xlsx');
        
        // Get the current day of the week
        const currentDay = new Date().toLocaleDateString('en-US', { weekday: 'long' });
        
        // Load the correct sheet based on the current day
        const sheet = workbook.getWorksheet(currentDay);

        // Iterate over the rows in the sheet
        sheet.eachRow(async (row, rowNumber) => {
            const keyword = row.getCell(1).value;

            if (!keyword) return; // Skip if keyword is empty

            // Perform Google Search
            await driver.get('https://www.google.com');
            const searchBox = await driver.findElement(By.name('q'));
            await searchBox.sendKeys(keyword, Key.RETURN);

            // Wait for suggestions to load and extract the text
            const suggestions = await driver.findElements(By.css('li.sbct'));
            const suggestionsText = await Promise.all(suggestions.map(async (s) => await s.getText()));

            // Determine the longest and shortest suggestions
            const longestSuggestion = suggestionsText.reduce((a, b) => a.length > b.length ? a : b, '');
            const shortestSuggestion = suggestionsText.reduce((a, b) => a.length < b.length ? a : b, '');

            // Write the longest and shortest options back to the Excel sheet
            row.getCell(2).value = longestSuggestion;
            row.getCell(3).value = shortestSuggestion;
        });

        // Save the updated Excel file
        await workbook.xlsx.writeFile('4BeatsQ1.xlsx');
    } finally {
        // Close the WebDriver
        await driver.quit();
    }
})();
