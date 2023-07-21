<script setup lang="ts">
import { NConfigProvider, NDialogProvider, NMessageProvider } from 'naive-ui';
// import HandleMedicalRecords from './components/HandleMedicalRecords.vue';
import Layout from './Layout.vue';


</script>

<template>
  <n-config-provider>
    <n-dialog-provider>
      <n-message-provider>
        <!-- <HandleMedicalRecords /> -->
        <Layout />
      </n-message-provider>
    </n-dialog-provider>
  </n-config-provider>
</template>

<style scoped></style>
