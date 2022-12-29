/* eslint-disable no-undef */
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

  while (start < file.size) {
    var chunk = file.slice(start, end);

    var reader = new FileReader();
    reader.readAsBinaryString(chunk);

    reader.onload = function (evt) {
      if (evt.target.readyState == FileReader.DONE) {
        var data = evt.target.result;
        console.log("data", data);
        var workbook = xlsx.read(data, { type: "binary" });
        // console.log("workbook", workbook);
        console.log("workbook.SheetNames", workbook.SheetNames);
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
      console.log("fileSelector", fileSelector);

      readBigCSVfile(fileSelector);

      // get current sheet values and store it in primaryValues array of objects
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      console.log("range", range);
      range.load("values");
      await context.sync();
      const primaryValues = range.values;

      // get uploaded sheet values and store it in secondaryValues array of objects
      const secondarySheet = context.workbook.worksheets.getItem("Sheet2");
      const secondaryRange = secondarySheet.getUsedRange();
      secondaryRange.load("values");
      await context.sync();
      const secondaryValues = secondaryRange.values;

      console.log("primaryValues", primaryValues);
      console.log("secondaryValues", secondaryValues);
    });
  } catch (error) {
    console.error(error);
  }
}
