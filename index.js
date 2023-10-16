const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const say = require('say');

(async () => {
  const browser = await puppeteer.launch({
    timeout: 10000,
  });
  const page = await browser.newPage();
  await page.setDefaultTimeout(5000000);

  await page.goto('http://result.biselahore.com/');

  await page.click('input[id="rdlistCourse_1"]');

  // put your starting roll number here
  const srn = ;
// put your ending roll number here
  const ern = ;


  let workbook;

  // Check if the Excel file already exists
  if (fs.existsSync('resut-2022.xlsx')) {
    // If it exists, open it
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('result-2022.xlsx');
  } else {
    // If it doesn't exist, create a new workbook
    workbook = new ExcelJS.Workbook();
    // Create a new worksheet
    workbook.addWorksheet('Result-2022');
  }

  // Get the worksheet by name
  const worksheet = workbook.getWorksheet('Result-2022');

  // Define the headers for the Excel sheet if it's a new worksheet
  if (worksheet.rowCount === 1) {
    const headers = [
      'Roll Number',
      'Name',
      'Father Name',
      'Institution',
      'Group',
      'SUBJECT',
      'Max Theory',
      'MAXIMUM MARKS PR',
      'MAX MARKS TOTAL',
      'MARKS OBTAINED TH1',
      'MARKS OBTAINED TH2',
      'MARKS OBTAINED PR',
      'MARKS OBTAINED TOTAL',
      'PERCENTILE SCORE',
      'REL. GRADE',
      'STATUS',
    ];

    worksheet.addRow(headers);

  }
  let crn = srn;

  while (crn <= ern) {   // Convert the roll number to a string
    const rollNumber = crn.toString();

    await page.type('#txtFormNo', rollNumber);

    await page.select('#ddlExamType', '2'); // Set value="1" for the dropdown
    await page.select('#ddlExamYear', '2022'); // Set value="2022" for the dropdown

    await page.click('#Button1');

    try {
      // Wait for the element #Name and #lblExamCenter
      await page.waitForSelector('#Name');
      await page.waitForSelector('#lblExamCenter');

      const name = await page.$eval('#Name', (element) => element.innerText.trim());
      const fatherName = await page.$eval('#lblFatherName', (element) => element.innerText.trim());
      const Group = await page.$eval('#lblGroup', (element) => element.innerText.trim());
      const institution = await page.$eval('#lblExamCenter', (element) => element.innerText.trim());

      const rows = await page.evaluate(() => {
        const rows = Array.from(document.querySelectorAll('#GridStudentData tr:not(:nth-child(1)):not(:nth-child(2)):not(:last-child)'));
        return rows.map((row, index) => {
          const columns = Array.from(row.querySelectorAll('td'));

          return columns.map((column) => column.innerText.trim());
        });
      });









      for (const row of rows) {

        worksheet.addRow([rollNumber, name, fatherName, institution, Group, ...row]);
      }
    } catch (error) {
      console.error(`Error for roll number ${rollNumber}: ${error.message}`);
    } finally {
      await page.goBack();
      await page.evaluate(() => {
        document.querySelector('#txtFormNo').value = '';
      });
    }
    crn++;
  }

  // Save the Excel file
  await workbook.xlsx.writeFile('result-2022.xlsx');

  await browser.close();


  say.speak(`target done.`);
})();
