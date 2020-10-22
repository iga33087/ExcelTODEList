<template>
  <div class="hello">
    <h1>{{ msg }}</h1>
    <p>
      For a guide and recipes on how to configure / customize this project,<br>
      check out the
      <a href="https://cli.vuejs.org" target="_blank" rel="noopener">vue-cli documentation</a>.
    </p>
    <h3>Installed CLI Plugins</h3>
    <ul>
      <li><a href="https://github.com/vuejs/vue-cli/tree/dev/packages/%40vue/cli-plugin-babel" target="_blank" rel="noopener">babel</a></li>
      <li><a href="https://github.com/vuejs/vue-cli/tree/dev/packages/%40vue/cli-plugin-router" target="_blank" rel="noopener">router</a></li>
      <li><a href="https://github.com/vuejs/vue-cli/tree/dev/packages/%40vue/cli-plugin-vuex" target="_blank" rel="noopener">vuex</a></li>
      <li><a href="https://github.com/vuejs/vue-cli/tree/dev/packages/%40vue/cli-plugin-eslint" target="_blank" rel="noopener">eslint</a></li>
    </ul>
    <h3>Essential Links</h3>
    <ul>
      <li><a href="https://vuejs.org" target="_blank" rel="noopener">Core Docs</a></li>
      <li><a href="https://forum.vuejs.org" target="_blank" rel="noopener">Forum</a></li>
      <li><a href="https://chat.vuejs.org" target="_blank" rel="noopener">Community Chat</a></li>
      <li><a href="https://twitter.com/vuejs" target="_blank" rel="noopener">Twitter</a></li>
      <li><a href="https://news.vuejs.org" target="_blank" rel="noopener">News</a></li>
    </ul>
    <h3>Ecosystem</h3>
    <ul>
      <li><a href="https://router.vuejs.org" target="_blank" rel="noopener">vue-router</a></li>
      <li><a href="https://vuex.vuejs.org" target="_blank" rel="noopener">vuex</a></li>
      <li><a href="https://github.com/vuejs/vue-devtools#vue-devtools" target="_blank" rel="noopener">vue-devtools</a></li>
      <li><a href="https://vue-loader.vuejs.org" target="_blank" rel="noopener">vue-loader</a></li>
      <li><a href="https://github.com/vuejs/awesome-vue" target="_blank" rel="noopener">awesome-vue</a></li>
    </ul>
    <input type="file" @change="exportFile">
    <table>
      <tr>
        <td>111</td>
        <td>222</td>
        <td>333</td>
      </tr>
      <tr>
        <td>444</td>
        <td>555</td>
        <td>666</td>
      </tr>
    </table>
  </div>
</template>

<script>
import xlsx from "xlsx"

export default {
  name: 'HelloWorld',
  props: {
    msg: String,
  },
  data() {
    return {
      url:""
    }
  },
  created() {
  },
  methods: {
    exportFile() {
      let arrayData = [
        ['name', 'age'],
        ['zhangsan', 20],
        ['lisi', 21],
        ['wangwu', 22],
        ['zhaoliu', 23],
        ['sunqi', 24],
      ];
      
      let jsonData = [{
        name: "zhangsan1",
        age: 30
      }, {
        name: "lisi1",
        age: 31
      }, {
        name: "wangwu1",
        age: 32
      }, {
        name: "zhaoliu1",
        age: 33
      }, {
        name: "sunqi1",
        age: 34
      }];
      
      let arrayWorkSheet = xlsx.utils.aoa_to_sheet(arrayData);
      let jsonWorkSheet = xlsx.utils.json_to_sheet(jsonData);

      let workBook = {
        SheetNames: ['arrayWorkSheet', 'jsonWorkSheet'],
        Sheets: {
          'arrayWorkSheet': arrayWorkSheet,
          'jsonWorkSheet': jsonWorkSheet,
        }
      };
      // 將workBook寫入檔案
      xlsx.writeFile(workBook, "./aa.xlsx");
    }
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped lang="scss">
h3 {
  margin: 40px 0 0;
}
ul {
  list-style-type: none;
  padding: 0;
}
li {
  display: inline-block;
  margin: 0 10px;
}
a {
  color: #42b983;
}
</style>
