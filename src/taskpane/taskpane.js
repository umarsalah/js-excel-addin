import * as xlsx from "xlsx";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("hello this is excel");

    document.getElementById("run").onclick = processSheets;
  }
});

function readBigCSVfile(fileUploaded) {
  console.log("reading big csv file ....");

  // read big csv file size with chunks
  var chunkSize = 1024 * 1024;

  var file = fileUploaded;

  var start = 0;
  var end = chunkSize;

  while (start < chunkSize) {
    var chunk = file.slice(start, end);

    var reader = new FileReader();
    reader.readAsBinaryString(chunk);

    reader.onload = async function (evt) {
      if (evt.target.readyState == FileReader.DONE) {
        var data = evt.target.result;
        const workbook = xlsx.read(data, { type: "binary" });

        // get the first column of the workbook and append it to a new file
        const firstSheet = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheet];
        const columnData = xlsx.utils.sheet_to_json(worksheet, {
          header: 1,
        });

        // get first column data in json
        const firstColumnData = columnData.map((row) => row[0]);
        console.log("firstColumnData", firstColumnData);
        
        // insert data to the running excel file
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          // clear old sheet cells data
          sheet.getUsedRange().clear();
          // set header of column firstColumnData[0]
          sheet.getRangeByIndexes(0, 0, 1, 1).values = [[firstColumnData[0]]];

          // set data of column firstColumnData.slice(1)
          console.log("Data length", firstColumnData.slice(1).length);
          firstColumnData.slice(1).forEach((value, index) => {
            sheet.getRangeByIndexes(index + 1, 0, 1, 1).values = [[value]];
          });

          await context.sync();
        });
      }
    };

    start += chunkSize;
    end = Math.min(end + chunkSize, file.size);
  }
}

async function processSheets() {
  try {
    await Excel.run(async (context) => {
      console.log("processing");
      // get uploaded file
      const fileSelector = document.getElementById("fileUpload").files[0];
      await readBigCSVfile(fileSelector);

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
