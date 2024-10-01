<template>
  <div class="table-container">
    <button
      style="cursor: pointer"
      type="button"
      id="excel_btn"
      @click="downloadExcel()"
    >
      EXCEL
    </button>
    <table
      id="table-to-excel"
      style="border: 1px; border-color: white; border-style: solid"
      cellpadding="5"
      cellspacing="0"
    >
      <thead>
        <tr>
          <th colspan="3">공연요금</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <th>구분</th>
          <th>s석</th>
          <th>VIP</th>
        </tr>
        <tr>
          <td>성인</td>
          <td>30,000</td>
          <td><img src="/image.png" width="20%" /></td>
        </tr>
        <tr>
          <td>청소년</td>
          <td>40,000</td>
          <td>60,000</td>
        </tr>
        <tr>
          <td>소인</td>
          <td colspan="2">미취학 아동 일반 요금의 50%</td>
        </tr>
        <tr>
          <th colspan="3">공연시간</th>
        </tr>
        <tr>
          <td rowspan="2">시간</td>
          <td colspan="2">13:00시 ~ 15:00시</td>
        </tr>
        <tr>
          <td colspan="2">17:00시 ~ 19:00시</td>
        </tr>
      </tbody>
    </table>
  </div>
</template>

<script setup>
import { ref } from "vue";
import ExcelJS from "exceljs";

const downloadExcel = async () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("공연 요금");

  // 공연요금 병합
  worksheet.mergeCells("A1:C1");
  worksheet.getCell("A1").value = "공연요금";

  // 구분, s석, VIP 헤더 추가
  worksheet.addRow(["구분", "s석", "VIP"]);

  // 데이터 추가
  worksheet.addRow(["성인", "30,000"]);
  worksheet.getCell("C3").value = ""; // 이미지 자리를 비워둠

  worksheet.addRow(["청소년", "40,000", "60,000"]);
  worksheet.addRow(["소인", "미취학 아동 일반 요금의 50%"]);
  worksheet.mergeCells("B5:C5");

  worksheet.mergeCells("A6:C6");
  worksheet.getCell("A6").value = "공연시간";

  // 시간 병합
  worksheet.addRow(["시간", "13:00시 ~ 15:00시"]);
  worksheet.mergeCells("B7:C7");
  worksheet.addRow(["", "17:00시 ~ 19:00시"]);
  worksheet.mergeCells("B8:C8");
  worksheet.mergeCells("A7:A8");

  // 이미지 삽입
  const image = await fetch("/image.png");
  const imageBlob = await image.blob();
  const imageBuffer = await imageBlob.arrayBuffer();
  const imageId = workbook.addImage({
    buffer: imageBuffer,
    extension: "png",
  });

  // 이미지 추가할 셀 위치 지정
  worksheet.addImage(imageId, "C3:C3");

  // 엑셀 파일 생성 및 다운로드
  const uint8Array = await workbook.xlsx.writeBuffer();
  const blob = new Blob([uint8Array], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "table_to_excel.xlsx";
  link.click();
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
