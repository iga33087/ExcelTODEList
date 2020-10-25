<template>
  <div class="tableBox" @dragover.prevent @drop.prevent="writeData">
    <div class="tableBoxMenu">
      <div class="tableBoxMenuTitle">
        <div class="tableBoxMenuTitleLabel">檔案名稱</div>
        <input type="text" class="tableBoxMenuTitleInput" v-model="fileName">
      </div>
      <div class="tableBoxMenuButtonList">
        <div class="tableBoxMenuButtonListItem" @click="exportFile('.xlsx')">匯出XLSX檔</div>
        <div class="tableBoxMenuButtonListItem" @click="exportFile('.xls')">匯出XLS檔</div>
        <div class="tableBoxMenuButtonListItem" @click="exportFile('.ods')">匯出ODS檔</div>
        <div class="tableBoxMenuButtonListItem" @click="exportFile('.csv')">匯出CSV檔</div>
        <div class="tableBoxMenuButtonListItem" @click="$refs.file1.click()">匯入
          <input type="file" ref="file1" accept=".xlsx,.xls,.ods,.csv" @change="writeData" style="display:none;">
        </div>
      </div>
    </div>
    <div class="tableBoxTag">
      <div class="tableBoxTagItem" v-for="(item,index) in json" :class="{'tableBoxTagItemOnChange':index==changeTableNum}" :key="index" @click="editChangeTableNum(index)">{{item.SheetNames}}<span @click.stop="delTable(index)">X</span></div>
      <div class="tableBoxTagItem" @click="addTable">+</div>
    </div>
    <div class="tableBoxListBox">
      <table border=1 class="tableBoxList">
        <tr>
          <td>#</td>
          <td v-for="(item,index) in Object.keys(json[changeTableNum].data[0])" :key="index">
            <div class="tableBoxListHeader">
              <div class="tableBoxListHeaderTitle">{{item}}</div>
              <div class="tableBoxListHeaderX" @click="delCell(item)">X</div>
            </div>
          </td>
          <td>操作</td>
        </tr>
        <tr v-for="(item,index) in json[changeTableNum].data" :key="index">
          <td>{{index+1}}</td>
          <td v-for="(item2,index2) in Object.keys(item)" :key="index2">
            <input type="text" class="tableBoxListInput" v-model="item[item2]">
          </td>
          <td><button @click="delRow(index)">刪除</button></td>
        </tr>
      </table>
    </div>
    <div class="tableBoxCtrlMenu">
      <div class="tableBoxCtrlMenuItem" @click="addCell">增加一欄</div>
      <div class="tableBoxCtrlMenuItem" @click="addRow">增加一列</div>
    </div>
  </div>
</template>

<script>
import xlsx from "xlsx"

export default {
  data() {
    return {
      fileName:"TODOList"
    }
  },
  props: {
    json: {
      type:Array
    },
    changeTableNum: {
      type:Number,
      default:0
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
    async writeData(e) {
      console.log(e)
      let file=e.target.files||e.dataTransfer.files
      let data=await this.importFile(file[0])
      console.log(file[0])
      this.fileName=""
      for(let i=0;i<file[0].name.split(".").length-1;i++) {
        this.fileName+=file[0].name.split(".")[i]
      }
      this.$emit("writeData",data)
    },
    exportFile(x) {
      this.$emit("exportFile",{data:this.json,fileName:this.fileName+x})
      //let arrayWorkSheet = xlsx.utils.aoa_to_sheet(x);
      //let workBook=this.getWorkBook(this.json)
      /*let jsonWorkSheet = xlsx.utils.json_to_sheet(this.json[0].data);
      let workBook = {
        SheetNames: ['jsonWorkSheet'],
        Sheets: {jsonWorkSheet:jsonWorkSheet}
      };*/
      //xlsx.writeFile(workBook, this.fileName+"."+this.fileType);
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
      this.$emit("editChangeTableNum",x)
    },
    addTable() {
      let defaultTitle="工作表"+Number(this.json.length+1)
      let title = prompt("請輸入工作表名",defaultTitle);
      if(!title) return 0
      this.$emit("addTable",title)
    },
    delTable(x) {
      this.$emit("delTable",x)
    },
    delCell(x) {
      this.$emit("delCell",x)
      this.$forceUpdate()
    },
    delRow(x) {
      this.$emit("delRow",x)
      this.$forceUpdate()      
    },
    addCell() {
      let defaultTitle="欄位"+Number(Object.keys(this.json[this.changeTableNum].data[0]).length+1)
      let title = prompt("請輸入欄位名",defaultTitle);
      if(!title) return 0
      this.$emit("addCell",title)
      this.$forceUpdate()
    },
    addRow() {
      this.$emit("addRow")
    }
  }
}
</script>

<style>

</style>