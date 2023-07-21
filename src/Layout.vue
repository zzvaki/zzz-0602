<script setup lang="ts">
import { MenuOption, NLayout, NLayoutContent, NLayoutHeader, NMenu } from 'naive-ui';
import { h } from 'vue';
import { RouterLink, RouterView } from 'vue-router';

const menuOptions: MenuOption[] = [
  {
    key: "HandleMedicalRecords",
    label: () =>
      h(RouterLink, { to: { name: "HandleMedicalRecords" } }, { default: () => '正畸老顾客' }),
  },
  {
    key: "SheetSplit",
    label: () =>
      h(RouterLink, { to: { name: "SheetSplit" } }, { default: () => '表格分割' }),
  },
]
</script>

<template>
  <n-layout>
    <n-layout-header>
      <n-menu mode="horizontal" :options="menuOptions" />
    </n-layout-header>
    <n-layout-content content-style="padding: 24px;">
      <RouterView />
    </n-layout-content>
  </n-layout>
</template>

<style scoped></style>
