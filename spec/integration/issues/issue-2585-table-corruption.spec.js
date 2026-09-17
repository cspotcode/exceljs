const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');
const TableXform = verquire('xlsx/xform/table/table-xform');

const FIXTURE = './spec/integration/data/test-issue-2585.xlsx';

describe('github issues', () => {
  it('issue 2585 - table ref/headerRow/totalsRow survive load/writeBuffer round-trip', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(FIXTURE);

    // initial parse
    const before = wb.getWorksheet('Sheet1').getTable('Table1');

    expect(before.ref).to.equal('A1:C3');
    expect(before.headerRow).to.equal(true);
    expect(before.totalsRow).to.equal(false);

    // round-trip parse
    const buffer = await wb.xlsx.writeBuffer();
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const after = wb2.getWorksheet('Sheet1').getTable('Table1');

    expect(after.ref).to.equal('A1:C3');
    expect(after.headerRow).to.equal(true);
    expect(after.totalsRow).to.equal(false);
    expect(after.model.columns.map(c => c.name)).to.deep.equal([
      'Header A',
      'Header B',
      'Header C',
    ]);

    // Validate table XML
    const zip = await JSZip.loadAsync(buffer);
    const tableStream = zip.files['xl/tables/table1.xml'].nodeStream();
    const rawTableModel = await new TableXform().parseStream(tableStream);
    expect(rawTableModel.tableRef).to.equal('A1:C3');
    expect(rawTableModel.totalsRowShown).to.equal('0');
  });
});
