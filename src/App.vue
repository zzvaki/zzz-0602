<script setup lang="ts">
import { NConfigProvider, NDialogProvider, NMessageProvider } from 'naive-ui';
import HandleMedicalRecords from './components/HandleMedicalRecords.vue';
</script>

<template>
  <n-config-provider>
    <n-dialog-provider>
      <n-message-provider>
        <HandleMedicalRecords />
      </n-message-provider>
    </n-dialog-provider>
  </n-config-provider>
</template>

<style scoped></style>
