<template>
  <div class="header">
    <div class="headerLogo">Excel TODEList</div>
    <div class="headerMenu">
    </div>
  </div>
</template>

<script>
export default {
  methods: {
    exportFile() {
      this.$$emit("exportFile")
    }
  }
}
</script>

<style>

</style>