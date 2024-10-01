<template>
  <div class="table-container">
    <!-- 버튼 클릭 이벤트-->
    <button
      style="cursor: pointer"
      type="button"
      id="excel_btn"
      @click="async () => await run()"
    >
      EXCEL
    </button>
    <img
      id="image-to-base64"
      ref="imageRef"
      src="data:image/png;base64,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"
      width="10%"
      crossorigin="anonymous"
    />
    <img ref="imageRef2" src="/image.png" width="20%" />
  </div>
</template>

<script setup>
import { ref } from "vue";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import html2canvas from "html2canvas";
const imageRef = ref(null);
const imageRef2 = ref(null);
const letsLearnExceljs = async () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("first sheet");

  worksheet.columns = [
    {
      header: "Id",
      key: "id",
      width: 10,
      style: {
        font: { size: 14 },
        numFmt: "@",
        alignment: { horizontal: "center" },
      },
    },
    {
      header: "Name",
      key: "name",
      width: 32,
      style: {
        font: { size: 14 },
        numFmt: "@",
        alignment: { horizontal: "center" },
      },
    },
    {
      header: "D.O.B.",
      key: "DOB",
      width: 20,
      outlineLevel: 1,
      style: {
        font: { size: 14 },
        numFmt: "YYYY-MM-DD",
        alignment: { horizontal: "center" },
      },
    },
    {
      header: "salary",
      key: "salary",
      width: 20,
      style: {
        font: { size: 14 },
        numFmt: "$#,##0",
        alignment: { horizontal: "right" },
      },
    },
  ];

  const data = [
    {
      id: 1,
      name: "Jamey",
      DOB: "2022-12-25",
      salary: 1000,
    },
    {
      salary: 2000,
      DOB: "2100-01-10",
      name: "Jimmy",
      id: 2,
    },
    {
      id: 3,
      name: "Jesus",
      DOB: "2000-12-25",
      salary: 1000000,
    },
  ];

  worksheet.insertRow(2, {
    id: 0,
    name: "Jenny",
    DOB: "2020-11-11",
    salary: 3000,
  });
  worksheet.insertRows(3, data);

  const headerStyle = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "cce6ff" },
  };

  const headerBorderStyle = {
    left: { style: "thin", color: { argb: "bfbfbf" } },
    right: { style: "thin", color: { argb: "bfbfbf" } },
  };

  for (let i = 1; i <= worksheet.columnCount; i++) {
    const headerEachCell = worksheet.getCell(`${String.fromCharCode(i + 64)}1`);
    headerEachCell.fill = headerStyle;
    headerEachCell.border = headerBorderStyle;
    headerEachCell.alignment = { horizontal: "center" };
  }
  worksheet.spliceRows(1, 0, [], [], []);
  worksheet.getCell("A2").value = "완성이다.";
  // 이미지 캡처 및 변환
  const canvas = await html2canvas(imageRef.value);
  const imageBase64 = canvas.toDataURL("image/png");
  const imageID = workbook.addImage({
    extension: "png",
    base64: imageBase64,
  });

  worksheet.addImage(imageID, {
    tl: { col: 2, row: 2 },
    ext: { width: 150, height: 100 },
  });

  const canvas2 = await html2canvas(imageRef2.value);
  let imageWidth = canvas2.width;
  let imageHeight = canvas2.height;
  const imageBase642 = canvas2.toDataURL("image/png");
  const imageID2 = workbook.addImage({
    extension: "png",
    base64: imageBase642,
  });
  const col = 4; // 열번호
  const row = 1; // 행번호
  while (imageHeight > 150 || imageWidth > 150) {
    imageHeight -= imageHeight / 2;
    imageWidth -= imageWidth / 2;
  }
  worksheet.addImage(imageID2, {
    tl: { col: col - 1, row: row - 1 },
    ext: { width: imageWidth, height: imageHeight },
    editAs: "oneCells",
  });
  worksheet.getRow(row).height = imageHeight * 0.7; // 이미지의 높이에 따른 행의 높이 조절
  // WORKSHEET2  ADDTABLE

  const worksheet2 = workbook.addWorksheet("second sheet");
  const arrayOfArraysData = data.map(({ id, name, DOB, salary }) => {
    return [id, name, DOB, salary];
  });

  const tableStartColumn = "C";
  const tableStartRow = "3";

  worksheet2.addTable({
    name: "letsMakeTable",
    ref: `${tableStartColumn}${tableStartRow}`,
    headerRow: true,
    totalsRow: true,
    style: {
      theme: "TableStyleLight7",
      showRowStripes: true,
    },
    columns: [
      { name: "id", filterButton: true },
      { name: "name", filterButton: true },
      { name: "D.O.B", filterButton: true },
      { name: "salary", filterButton: true },
    ],
    rows: arrayOfArraysData,
  });

  worksheet2.eachRow((row, rowNo) => {
    row.height = 18;
    row.eachCell((cell, colNo) => {
      if (cell.value || cell.value === 0) {
        const eachCell = row.getCell(colNo);
        eachCell.font = { size: 14 };

        if (colNo === 6) eachCell.numFmt = "$#,##0";
      }
    });
  });

  for (let i = 1; i <= worksheet2.columnCount; i++) {
    const eachColumn = worksheet2.getColumn(i);

    if (i === 6) {
      eachColumn.alignment = { horizontal: "right" };
    } else eachColumn.alignment = { horizontal: "center" };

    if (eachColumn.values.length !== 0) {
      eachColumn.width = 20;
    }
  }

  const tableColumnLength = Object.keys(data[0]).length;

  for (
    let i = +tableStartRow - 1;
    i < tableColumnLength + +tableStartRow - 1;
    i++
  ) {
    const headerEachCell = worksheet2.getCell(
      `${String.fromCharCode(i + 65)}${tableStartRow}`
    );
    headerEachCell.alignment = { horizontal: "center" };
  }

  return workbook;
};

const run = async () => {
  console.log("run");
  const workbook = await letsLearnExceljs();
  const mimeType = {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  };
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], mimeType);
  saveAs(blob, "testExcel.xlsx");
};
</script>

<style scoped>
.table-container {
  position: relative;
  display: inline-block;
}

.download-btn {
  position: absolute;
  top: 0;
  right: 0;
  padding: 8px 16px;
  background-color: #42b883;
  color: white;
  border: none;
  cursor: pointer;
  font-size: 14px;
  border-radius: 4px;
}

.download-btn:hover {
  background-color: #3a9e6c;
}

table {
  width: 100%;
  text-align: center;
  border-collapse: collapse;
}

th,
td {
  padding: 10px;
}
</style>
