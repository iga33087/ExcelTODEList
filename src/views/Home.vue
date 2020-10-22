<template>
  <div class="home">
    <!--<div class="tableBox" @dragover.prevent @drop.prevent="writeData">
      <div class="tableBoxTag">
        <div class="tableBoxTagItem" v-for="(item,index) in json" :class="{'tableBoxTagItemOnChange':index==changeTableNum}" :key="index" @click="changeTableNum=index">{{item.SheetNames}}<span @click.stop="delTable(index)">X</span></div>
        <div class="tableBoxTagItem" @click="addTable">+</div>
      </div>
      <table border=1 class="tableBoxList">
        <tr>
          <td>#</td>
          <td v-for="(item,index) in Object.keys(json[changeTableNum].data[0])" :key="index">
            {{item}}
            <span @click="delCell(item)">刪除</span>
          </td>
        </tr>
        <tr v-for="(item,index) in json[changeTableNum].data" :key="index">
          <td>{{index+1}}</td>
          <td v-for="(item2,index2) in Object.keys(item)" :key="index2">
            <input type="text" class="tableBoxListInput" v-model="item[item2]">
          </td>
        </tr>
      </table>
      <button @click="exportFile">OK</button>
      <button @click="addCell">增加一欄</button>
      <button @click="addRow">增加一列</button>
    </div>-->
    <XlsxTable :json="json" :changeTableNum="changeTableNum" @editChangeTableNum="editChangeTableNum" @addTable="addTable" @delTable="delTable" @delCell="delCell" @addCell="addCell" @addRow="addRow" @delRow="delRow" @exportFile="exportFile" @writeData="writeData"></XlsxTable>
  </div>
</template>

<script>
// @ is an alias to /src
import xlsx from "xlsx"
import XlsxTable from "@/components/XlsxTable"

export default {
  name: 'Home',
  components:{XlsxTable},
  data() {
    return {
      changeTableNum:0,
      fileName:"123",
      fileType:"xls",
      json:[
        {SheetNames:"工作表1",data:[
          {"標題":"測試標題","備註":"測試備註"},
        ]},
      ]
    }
  },
  methods: {
    async importFile(file) {
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            /* Parse data */
            const bstr = e.target.result;
            const wb = xlsx.read(bstr, {
                type: 'binary'
            });
            /* Get first worksheet */
            const wsname = wb.SheetNames[0];
            const ws = wb.Sheets[wsname];
            /* Convert array of arrays */
            const data = xlsx.utils.sheet_to_json(ws, {
                header: 1
            });
            console.log(data)
            let arr=[]
            for(let item of wb.SheetNames) {
              let itemJson=xlsx.utils.sheet_to_json(wb.Sheets[item], {
                header: 1
              });
              let obj={
                SheetNames:item,
                data:[]
              }
              for(let i=1;i<itemJson.length;i++) {
                let obj2= {}
                for(let j in itemJson[0]) {
                  obj2[itemJson[0][j]]=itemJson[i][j]
                }
                obj.data.push(obj2)
              }
              arr.push(obj)
            }
            resolve(arr)
        };
        reader.readAsBinaryString(file)
      })
    },
    async writeData(x) {
      this.json=x
    },
    exportFile(x) {
      //let arrayWorkSheet = xlsx.utils.aoa_to_sheet(x);
      let workBook=this.getWorkBook(x.data)
      /*let jsonWorkSheet = xlsx.utils.json_to_sheet(this.json[0].data);
      let workBook = {
        SheetNames: ['jsonWorkSheet'],
        Sheets: {jsonWorkSheet:jsonWorkSheet}
      };*/
      xlsx.writeFile(workBook, x.fileName);
    },
    getWorkBook(x) {
      let workBook={
        SheetNames:x.map(res=>res.SheetNames),
        Sheets:{}
      }
      for(let item of x) {
        workBook.Sheets[item.SheetNames]=xlsx.utils.json_to_sheet(item.data);
      }
      console.log(workBook)
      return workBook
    },
    editChangeTableNum(x) {
      this.changeTableNum=x
    },
    addTable(x) {
      let obj= {
        SheetNames:x,
        data:[
          {"標題":"範例標題","備註":"範例備註"},
        ]
      }
      this.json.push(obj)
    },
    delTable(x) {
      if(this.json.length==1) {
        alert("工作表不得全刪")
        return 0
      }
      this.json.splice(x,1)
      this.changeTableNum=0
      this.$forceUpdate()
    },
    delCell(x) {
      if(Object.keys(this.json[this.changeTableNum].data[0]).length==1) {
        alert("欄位不得全刪")
        return 0
      }
      for(let item of this.json[this.changeTableNum].data) {
        delete item[x];
      }
      this.$forceUpdate()
    },
    delRow(x) {
      if(this.json[this.changeTableNum].data.length==1) {
        alert("列不得全刪")
        return 0
      }
      this.json[this.changeTableNum].data.splice(x,1)
      this.$forceUpdate()
    },
    addCell(x) {
      for(let item of this.json[this.changeTableNum].data) {
        item[x]=""
      }
      this.$forceUpdate()
    },
    addRow() {
      if(!this.json[this.changeTableNum].data.length) return 0
      let obj={}
      for(let item of Object.keys(this.json[this.changeTableNum].data[0])) {
        obj[item]=""
      }
      this.json[this.changeTableNum].data.push(obj)
      console.log(this.json)
    }
  }
}
</script>
