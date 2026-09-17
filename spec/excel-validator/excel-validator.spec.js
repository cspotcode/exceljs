const path = require('path');
const util = require('util');

const ExcelJS = verquire('exceljs');

const {checkExcelValidity} = require('../utils/excel-validator');

const FIXTURES_DIR = path.resolve(__dirname, '..', '..');

// Loads `relPath` with exceljs, re-serializes it, and asserts the round-tripped output
// opens cleanly in real Excel (no repair/data-recovery needed).
async function checkRoundTrip(relPath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(FIXTURES_DIR, relPath));
  const buffer = await wb.xlsx.writeBuffer();
  const {valid, successfulMode, errors} = await checkExcelValidity(buffer, {
    baseName: path.basename(relPath),
  });
  const message =
      `Should open in "normal" mode. Opened in mode: ${util.inspect(successfulMode)}. ` +
      `(Errors from Excel: normal: ${errors.normal}, repair: ${errors.repair}, ` +
      `extractData: ${errors.extractData})`;
  expect(valid, message).to.equal(true);
}

describe('Excel Validator (round-trip files open cleanly in Excel)', function () {
  it('1519293514 - Krishnapatnam line-up export', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/1519293514-KRISHNAPATNAM_LINE_UP.xlsx');
  });

  it('1904 date system workbook', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/1904.xlsx');
  });

  it('bogus defined name', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/bogus-defined-name.xlsx');
  });

  it('chart sheet', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/chart-sheet.xlsx');
  });

  it('duplicate row test', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/duplicateRowTest.xlsx');
  });

  it('fibonacci formulas workbook', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/fibonacci.xlsx');
  });

  it('formulas workbook', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/formulas.xlsx');
  });

  it('gold standard workbook', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/gold.xlsx');
  });

  it('hancell-produced file', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/hancell-file.xlsx');
  });

  it('hidden-sheet test - Google Sheets export', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/hidden-test/google-sheets.xlsx');
  });

  it('huge workbook', async function () {
    this.timeout(60000);
    await checkRoundTrip('spec/integration/data/huge.xlsx');
  });

  it('images workbook', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/images.xlsx');
  });

  it('many-columns workbook', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/many-columns.xlsx');
  });

  it('shared string with escape', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/shared_string_with_escape.xlsx');
  });

  it('issue 1364 - Incorrect Worksheet Name on Streaming XLSX Reader', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-issue-1364.xlsx');
  });

  it('issue 1575 fixture', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-issue-1575.xlsx');
  });

  it('issue 257 - worksheet order is not respected', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-issue-257.xlsx');
  });

  it('issue 623 - borders for merged cell when rewriting a workbook', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-issue-623.xlsx');
  });

  it('issue 877 - hyperlink without text crashes on write', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-issue-877.xlsx');
  });

  it('issue 880 - malformed comment crashes on write', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-issue-880.xlsx');
  });

  it('issue 2585 - table ref corruption on load/write round-trip', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-issue-2585.xlsx');
  });

  it('pull request 1204 fixture', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-pr-1204.xlsx');
  });

  it('pull request 1220 - worksheet should not be undefined', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-pr-1220.xlsx');
  });

  it('pull request 567 - whole column defined names', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-pr-567.xlsx');
  });

  it('pull request 728 - read worksheet hidden state', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-pr-728.xlsx');
  });

  it('row styles workbook', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-row-styles.xlsx');
  });

  it('legacy comments fixture', async function () {
    this.timeout(30000);
    await checkRoundTrip('test/data/comments.xlsx');
  });

  it('legacy default-font fixture', async function () {
    this.timeout(30000);
    await checkRoundTrip('test/data/test-default-font.xlsx');
  });
});

describe.skip('Excel Validator (round-trip files open cleanly in Excel) - known failures', function () {
  it('date issue workbook', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/dateIssue.xlsx');
  });

  it('hidden-sheet test - LibreOffice Calc as Excel 2007-365', async function () {
    this.timeout(30000);
    await checkRoundTrip(
      'spec/integration/data/hidden-test/libre-calc-as-excel-2007-365.xlsx'
    );
  });

  it('hidden-sheet test - LibreOffice Calc as Office Open XML Spreadsheet', async function () {
    this.timeout(30000);
    await checkRoundTrip(
      'spec/integration/data/hidden-test/libre-calc-as-office-open-xml-spreadsheet.xlsx'
    );
  });

  it('issue 163 - Error while using xlsx readFile method', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-issue-163.xlsx');
  });

  it('issue 1669 - optional autofilter and custom autofilter on tables', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-issue-1669.xlsx');
  });

  it('issue 176 - Unexpected xml node in parseOpen', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-issue-176.xlsx');
  });

  it('issue 771 - dataValidation without type and with formula1/formula2', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-issue-771.xlsx');
  });

  it('issue 988 - table without autofilter model', async function () {
    this.timeout(30000);
    await checkRoundTrip('spec/integration/data/test-issue-988.xlsx');
  });
});
