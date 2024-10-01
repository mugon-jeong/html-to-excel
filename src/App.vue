<script setup>
import PerformanceTable from "./components/PerformanceTable.vue";
import PerformanceTable2 from "./components/PerformanceTable2.vue";
import PerformanceTable3 from "./components/PerformanceTable3.vue";
</script>

<template>
  <div class="container">
    <PerformanceTable />
  </div>
  <div class="container">
    <PerformanceTable2 />
  </div>
  <div class="container">
    <PerformanceTable3 />
  </div>
</template>

<style scoped>
.container {
  display: flex;
  justify-content: center; /* 수평 중앙 정렬 */
  align-items: center; /* 수직 중앙 정렬 */
  height: 100vh; /* 화면 전체 높이 */
}

.logo {
  height: 6em;
  padding: 1.5em;
  will-change: filter;
  transition: filter 300ms;
}

.logo:hover {
  filter: drop-shadow(0 0 2em #646cffaa);
}

.logo.vue:hover {
  filter: drop-shadow(0 0 2em #42b883aa);
}
</style>
