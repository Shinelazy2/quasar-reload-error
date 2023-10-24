<template>
  <q-page class="row items-center justify-evenly">
    <q-btn @click="convertDB" label="text to db"></q-btn>

  </q-page>
</template>

<script setup lang="ts">

const convertDB = async () => {
  return await window.api.readTxtFile('./txt.txt')
}

</script>
