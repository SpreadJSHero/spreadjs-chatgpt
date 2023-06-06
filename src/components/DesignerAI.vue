<script lang="ts">
import { defineComponent, ref } from 'vue'
import '@grapecity/spread-sheets/styles/gc.spread.sheets.excel2013white.css';
import "@grapecity/spread-sheets-designer/styles/gc.spread.sheets.designer.min.css";

import * as GC from "@grapecity/spread-sheets";
import "@grapecity/spread-sheets-pivot-addon";
import "@grapecity/spread-sheets-tablesheet";
import "@grapecity/spread-sheets-resources-zh";
import "@grapecity/spread-sheets-designer-resources-cn";
import * as DeisgnerGC from "@grapecity/spread-sheets-designer";
import GcSpreadDesigner from "@grapecity/spread-sheets-designer-vue"
GC.Spread.Common.CultureManager.culture("zh-cn");

import {openai} from './openai';
import './GPT_Functions'
import {config} from './GPT_Config'


let designer: DeisgnerGC.Spread.Sheets.Designer.Designer;
let chatDisabled = false;
let chatHistory: any[] = [{"role": "system", "content": "你是一个数据分析助手，我将给你提供数据。"},];
const chatText = ref("");
// let fbx: GC.Spread.Sheets.FormulaTextBox.FormulaTextBox;

const loading = ref(false)

const scrollbarMax = ref(0)

export default defineComponent({
  components: {
    "gc-spread-designer": GcSpreadDesigner
  },
  data: function () {
    return {
        styleInfo: { height: "100%", width: "75%", color: "black" },
        chatDisabled,
        chatHistory,
        chatText,
        scrollbarMax,
        loading,
    }
  },
  mounted(){
    },
  methods: {
    designerInitialized: function(value: DeisgnerGC.Spread.Sheets.Designer.Designer){
      designer = value;
      designer.setConfig(config);
      let spread = designer.getWorkbook() as GC.Spread.Sheets.Workbook;
      spread.options.allowDynamicArray = true;
      // fbx = new GC.Spread.Sheets.FormulaTextBox.FormulaTextBox(this.$refs.formulaBar as HTMLElement, {rangeSelectMode: true, absoluteReference: false});
      // fbx.workbook(designer.getWorkbook() as GC.Spread.Sheets.Workbook);
    },
    startChat: async function(){
      if(!this.chatText){
        return;
      }
      this.chatDisabled = true;
      this.loading = true;

      this.chatHistory.push({
        role: "user",
        content: this.chatText,
      })
      setTimeout(() => {
        this.scrollbarMax = (this.$refs.innerRef as HTMLDivElement)!.clientHeight - (this.$refs.innerRef as any)!.parentNode.parentNode!.clientHeight + 20;
        if(this.scrollbarMax > 0){
          (this.$refs.scrollbarRef as any)!.setScrollTop(this.scrollbarMax)
        }
      }, 100)
      try{
        const completion = await openai.createChatCompletion({
          model: "gpt-3.5-turbo",
          messages: this.chatHistory,
        });
        this.chatHistory.push({
          role: "assistant",
          content: completion.data.choices[0].message!.content,
        })
        console.log(completion);
        setTimeout(() => {
          this.scrollbarMax = (this.$refs.innerRef as HTMLDivElement)!.clientHeight - (this.$refs.innerRef as any)!.parentNode.parentNode!.clientHeight + 20;
          if(this.scrollbarMax > 0){
            (this.$refs.scrollbarRef as any)!.setScrollTop(this.scrollbarMax)
          }
        }, 100)
      }
      catch{
        
      }
      finally{
        this.chatText = "";
        this.chatDisabled = false;
        this.loading = false;
      }
    },
    getSelectionValue: function(){
      // if(!fbx.text()){
      //   return;
      // }
      let spread = designer.getWorkbook() as GC.Spread.Sheets.Workbook;
      let sheet = spread.getActiveSheet();
      // let node = GC.Spread.Sheets.CalcEngine.formulaToRanges(sheet, fbx.text(), 0, 0)[0] as any;
      // let range = node.ranges[0] as GC.Spread.Sheets.Range

      let range = sheet.getSelections()[0];
      let data = sheet.getCsv(range.row, range.col, range.rowCount, range.colCount, "\n", ",");
      this.chatText = data;
    }
  }
});

</script>

<template>
  <div class="container">
    <gc-spread-designer :styleInfo="styleInfo" @designerInitialized="designerInitialized"></gc-spread-designer>
    <div class="chatContainer"  v-loading="loading">
      <div class="chatRecord">
        <el-scrollbar ref="scrollbarRef" >
          <div ref="innerRef">
            <el-timeline>
              <el-timeline-item v-for="item in chatHistory.filter(c => c.role !== 'system')" center :timestamp="item.role" placement="top">
                <el-card>
                  <div style="white-space: pre-line;">{{ item.content }}</div>
                  <!-- <p>{{ item.time }}</p> -->
                </el-card>
              </el-timeline-item>
            </el-timeline>
          </div>
        </el-scrollbar>
      </div>
      <el-divider />
      <div class="messageContainer">
        <el-row>
          <!-- <el-col :span="12">
            <div ref="formulaBar" spellcheck="false" class="formulaBar"></div>
          </el-col>
          <el-col :span="2">
          </el-col> -->
          <el-col :span="8">
            <el-button @click="getSelectionValue" >获取数据</el-button>
          </el-col>
        </el-row>
        <el-row>
          <el-col :span="24">
            <div style="margin-top: 1rem;">
              <el-input
                v-model="chatText"
                :rows="5"
                type="textarea"
                placeholder="Please input"
              />
            </div>
          </el-col>
        </el-row>
        <el-row>
          <div style="margin-top: 1rem;">
            <el-button @click="startChat" :disabled="chatDisabled">发送</el-button>
          </div>
        </el-row>
      </div>
    </div>
  </div>
</template>

<style scoped>
  .container{
    height: 100%;
    display: flex;
  }
  .chatContainer{
    width: 30%;
    height: 100%;
    /* padding: 0 0 0 1rem; */
  }
  .chatRecord{
    /* background-color: aliceblue;  */
    height: CALC(100% - 300px);
    padding: 1rem 0 0 0;
  }
  .messageContainer{
    height: 250px;
    padding: 0 0 0 1rem;
  }
  .formulaBar{
    background-color: white;
    height: 100%;
  }
</style>

<style>
  .ribbon-button-item-text {
    color: #000;
  }
  .gc-range-select .range-input{
    background-color: white;
    color: #000;
  }
  .gc-sjs-designer-dialog .dialog-footer button{
    color: #000;
  }

  .ribbon-button-formulaGenerate{
      background-image: url("/icons/formulaGenerate.png")
  }
  .ribbon-button-formulaAnalyze{
      background-image: url("/icons/formulaAnalyze.png")
  }
  .ribbon-button-pivotTableSuggest{
      background-image: url("/icons/pivotTableSuggest.png")
  }
  .ribbon-button-formulaOptimize{
      background-image: url("/icons/formulaOptimize.png")
  }
  .ribbon-button-dataAnalyze{
      background-image: url("/icons/dataAnalyze.png")
  }
</style>