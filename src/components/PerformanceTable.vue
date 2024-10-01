<template>
  <div class="table-container">
    <!-- 버튼 클릭 이벤트-->
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
          <td>
            <img src="/image.png" width="20%" />
          </td>
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
import * as XLSX from "xlsx"; // 라이브러리 import
const downloadExcel = () => {
  const workBook = XLSX.utils.book_new();
  const sheetData = XLSX.utils.table_to_sheet(
    document.getElementById("table-to-excel")
  );
  XLSX.utils.book_append_sheet(workBook, sheetData, "table_to_excel");
  XLSX.writeFile(workBook, "table_to_excel.xlsx");
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
