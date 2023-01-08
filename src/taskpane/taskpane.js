import * as xlsx from "xlsx";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("hello this is excel");

    document.getElementById("run").onclick = processSheets;
  }
});

async function processSheets() {
  try {
    await Excel.run(async (context) => {
      console.log("processing");
      // get uploaded files
      const primaryFile = document.getElementById("primaryFile").files[0];
      const secondaryFile = document.getElementById("secondaryFile").files[0];

      // Return error if no 2 files uploaded
      if (!primaryFile && !secondaryFile) {
        return context
          .sync()
          .then(function () {
            Office.context.ui.displayDialogAsync(
              "https://localhost:3000/error.html",
              { height: 30, width: 20 },
              function (result) {
                var dialog = result.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (event) {
                  console.log(event.message);
                  dialog.close();
                });
              }
            );
          })
          .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
              console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
          });
      }

      // Compare two files columns and rows
      const primaryFileData = await readCSVfile(primaryFile);
      const secondaryFileData = await readCSVfile(secondaryFile);
      
      // fill sheets with data
      await fillSheetsWithData(primaryFileData, secondaryFileData, context);

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }

  async function readCSVfile(fileUploaded) {
    const dataObject = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = e.target.result;
        const workbook = xlsx.read(data, { type: "binary" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
        const dataObject = {};
        json.forEach((row) => {
          dataObject[row[0]] = row;
        });
        resolve(dataObject);
      };
      reader.onerror = (error) => reject(error);
      reader.readAsBinaryString(fileUploaded);
    });
    return dataObject;
  }

  async function fillSheetsWithData(primaryFileData, secondaryFileData, context) {
    // We will get the Primary Worksheet and fill it with primaryFileData and the same for secondaryFileData
    const primarySheet = context.workbook.worksheets.add("Primary Worksheet");
    const secondarySheet = context.workbook.worksheets.add("Secondary Worksheet");

    /******************************************** Primary Worksheet ***********************************************/
    const primaryHeaderRow = primaryFileData[Object.keys(primaryFileData)[Object.keys(primaryFileData).length - 1]];
    const primaryHeaderRange = primarySheet.getRangeByIndexes(0, 0, 1, primaryHeaderRow.length);
    primaryHeaderRange.values = [primaryHeaderRow];
    const primaryRows = Object.keys(primaryFileData).map((key) => primaryFileData[key]);
    primaryRows.pop();
    primaryRows.forEach((row) => {
      primarySheet.getRangeByIndexes(primaryRows.indexOf(row) + 1, 0, 1, row.length).values = [row];
    });

    /******************************************** Secondary Worksheet ********************************************/
    const secondaryHeaderRow =
      secondaryFileData[Object.keys(secondaryFileData)[Object.keys(secondaryFileData).length - 1]];
    const secondaryHeaderRange = secondarySheet.getRangeByIndexes(0, 0, 1, secondaryHeaderRow.length);
    secondaryHeaderRange.values = [secondaryHeaderRow];
    const secondaryRows = Object.keys(secondaryFileData).map((key) => secondaryFileData[key]);
    secondaryRows.pop();
    secondaryRows.forEach((row) => {
      secondarySheet.getRangeByIndexes(secondaryRows.indexOf(row) + 1, 0, 1, row.length).values = [row];
    });

    await context.sync();
  }
}
