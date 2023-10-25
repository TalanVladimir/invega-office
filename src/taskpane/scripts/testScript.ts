/* global console, Excel */

export const testScript = async () => {
  try {
    await Excel.run(async (context) => {
      const workbook: Excel.Workbook = context.workbook;
      const sheet: Excel.Worksheet = workbook.worksheets.getActiveWorksheet();

      const infoRow = sheet.getCell(0, 0);

      infoRow.values = [["pradžia"]];
      infoRow.format.fill.color = "yellow";
      await context.sync();

      const TableHeadersRange = sheet.getRange("A7:Z8");
      TableHeadersRange.load("values");

      await context.sync();
      const TableHeadersValues = await TableHeadersRange.values;

      if (
        TableHeadersValues[0][0] === "Įmokos ID" &&
        TableHeadersValues[0][24] === "Požymis" &&
        TableHeadersValues[1][0] === 1
      ) {
        infoRow.values = [["super"]];
        infoRow.format.fill.color = "green";

        const deleteRow = sheet.getCell(7, 0);
        deleteRow.getEntireRow().delete("Up");
      } else {
        infoRow.values = [["klaida"]];
        infoRow.format.fill.color = "red";
      }

      TableHeadersRange.format.autofitColumns();
      TableHeadersRange.untrack();

      await context.sync();

      let i = 9;
      let lastRow = 0;

      do {
        const itemRange = sheet.getCell(i, 0);
        itemRange.load("text");

        await context.sync();
        const itemValues = itemRange.text;
        if (itemValues[0][0] === "") {
          i = 0;
        } else {
          lastRow = i + 1;
          ++i;
        }

        itemRange.untrack();
        // console
      } while (i > 0);

      infoRow.values = [[lastRow]];
      infoRow.format.fill.color = "purple";

      const zzz = sheet.getUsedRange();

      infoRow.values = [[`getRow: ${zzz.getLastRow()}`]];
      infoRow.format.fill.color = "orange";
      await context.sync();
      // console.log(zzz.getLastRow());

      //   const range = workbook.getSelectedRange();

      //   await context.sync();
      //   range.getEntireRow().delete("Up");

      //   const values: string[][] = [["asd", "addd"]];

      //   await sheet.getRange("A3:B3").select();
      //   await context.sync();

      // Read the range address
      // await   range.load("address");
      //   range.values = values;
      //   await context.sync();

      // Update the fill color
      //   range.format.fill.color = "blue";

      //   await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
};
