<template>
  <div class="header">
    <div class="headerLogo">Excel TODOList</div>
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