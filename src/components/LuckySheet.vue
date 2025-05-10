<template>
  <div style="position: absolute; top: 0">
    <input id="uploadBtn" type="file" @change="loadExcel" />

    <span>Or Load remote xlsx file:</span>

    <select v-model="selected" @change="selectExcel">
      <option disabled value="">Choose</option>
      <option v-for="option in options" :key="option.text" :value="option.value">
        {{ option.text }}
      </option>
    </select>

    <a href="javascript:void(0)" @click="downloadExcel">Download source xlsx file</a>
  </div>
  <div id="luckysheet"></div>
  <div v-show="isMaskShow" id="tip">Downloading</div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { exportExcel } from '../utils/export'
import { isFunction } from '../utils/is'
import LuckyExcel from 'luckyexcel'

const isMaskShow = ref(false)
const selected = ref('')
const jsonData = ref({}) // Not strictly needed for AI function but part of existing logic
const options = ref([
  { text: 'Money Manager.xlsx', value: 'https://minio.cnbabylon.com/public/luckysheet/money-manager-2.xlsx' },
  {
    text: 'Activity costs tracker.xlsx',
    value: 'https://minio.cnbabylon.com/public/luckysheet/Activity%20costs%20tracker.xlsx',
  },
  {
    text: 'House cleaning checklist.xlsx',
    value: 'https://minio.cnbabylon.com/public/luckysheet/House%20cleaning%20checklist.xlsx',
  },
  {
    text: 'Student assignment planner.xlsx',
    value: 'https://minio.cnbabylon.com/public/luckysheet/Student%20assignment%20planner.xlsx',
  },
  {
    text: 'Credit card tracker.xlsx',
    value: 'https://minio.cnbabylon.com/public/luckysheet/Credit%20card%20tracker.xlsx',
  },
  { text: 'Blue timesheet.xlsx', value: 'https://minio.cnbabylon.com/public/luckysheet/Blue%20timesheet.xlsx' },
  {
    text: 'Student calendar (Mon).xlsx',
    value: 'https://minio.cnbabylon.com/public/luckysheet/Student%20calendar%20%28Mon%29.xlsx',
  },
  {
    text: 'Blue mileage and expense report.xlsx',
    value: 'https://minio.cnbabylon.com/public/luckysheet/Blue%20mileage%20and%20expense%20report.xlsx',
  },
])

// Helper function to define common Luckysheet options with AI formula hook
const getLuckysheetOptions = (specificOptions = {}) => {
  return {
    container: 'luckysheet', //luckysheet is the container id
    lang: 'en', // Set language, e.g., 'en' or 'zh'
    allowUpdate: true, // Important for promises from AI function to update cell
    ...specificOptions, // Merge specific options like data, title
    hook: {
      workbookCreateAfter: function () {
        console.log('workbookCreateAfter hook triggered');
        // 1) Add metadata so =AI appears in autocomplete/help
        // Check if AI function is already registered to avoid duplicates
        if (window.luckysheet && window.luckysheet.formula && !window.luckysheet.formula.functionList.find(f => f.n === "AI")) {
          window.luckysheet.formula.functionList.push({
            n: "AI", // Function name
            t: "data", // Function type (data, math, text, etc.)
            d: "AI(prompt, [range]) – Generates text using a WebGPU LLM.", // Description
            a: "AI-powered assistant. Provide a prompt and optionally a cell range for context.", // Arguments description
            m: [1, 2] // [minParams, maxParams] (prompt is required, range is optional)
          });
          console.log('AI function metadata registered.');
        } else if (window.luckysheet && window.luckysheet.formula && window.luckysheet.formula.functionList.find(f => f.n === "AI")) {
          console.log('AI function metadata already registered.');
        }


        // 2) Define how AI(...) actually runs
        if (window.luckysheet && window.luckysheet.formula) {
            window.luckysheet.formula.functionImplement.AI = function (promptText, rangeData) {
            console.log('AI function called with prompt:', promptText, 'and rangeData:', rangeData);

            // Ensure your AI generation function (window.AIgenerator) is available
            if (typeof window.AIgenerator !== 'function') {
              console.error('window.AIgenerator is not defined or not a function.');
              return "#ERROR: AIgenerator not found!"; // Or "⌛ AI Model is loading..."
            }

            let contextString = "";
            if (rangeData !== undefined && rangeData !== null) {
              // Flatten the rangeData. It can be a single value, 1D array, or 2D array.
              const flatValues = Array.isArray(rangeData) ?
                rangeData.flat(Infinity) : // Flattens to any depth
                [rangeData]; // Ensure it's an array

              contextString = flatValues
                .filter(v => v !== null && v !== undefined && v.toString().trim() !== "") // Filter out null, undefined, empty strings
                .join(" "); // Join non-empty values with a space
            }

            const fullPrompt = contextString ? `${promptText} [CONTEXT: ${contextString}]` : promptText;
            console.log('Full prompt for AIgenerator:', fullPrompt);

            // Return a Promise. Luckysheet will wait for it to resolve and then update the cell.
            return window.AIgenerator(
              [
                { role: "system", content: "You are a helpful spreadsheet assistant." },
                { role: "user", content: fullPrompt }
              ],
              { max_new_tokens: 128 } // Example options, adjust as needed
            ).then(response => {
              console.log('AIgenerator response:', response);
              // Assuming response is an array with the first element having generated_text
              if (response && Array.isArray(response) && response.length > 0 && response[0] && typeof response[0].generated_text === 'string') {
                return response[0].generated_text;
              }
              console.error('Unexpected AIgenerator response format:', response);
              return "#ERROR: AI response format issue";
            }).catch(error => {
              console.error("AI Function Execution Error:", error);
              return "#ERROR: AI execution failed";
            });
          };
          console.log('AI function implementation registered.');
        } else {
            console.error('Luckysheet formula object not available for AI implementation.');
        }
      }
    }
  };
};

const loadExcel = (evt) => {
  const files = evt.target.files
  if (files == null || files.length == 0) {
    alert('No files wait for import')
    return
  }

  let name = files[0].name
  let suffixArr = name.split('.'),
    suffix = suffixArr[suffixArr.length - 1]
  if (suffix != 'xlsx') {
    alert('Currently only supports the import of xlsx files')
    return
  }
  LuckyExcel.transformExcelToLucky(files[0], function (exportJson, luckysheetfile) {
    if (exportJson.sheets == null || exportJson.sheets.length == 0) {
      alert('Failed to read the content of the excel file, currently does not support xls files!')
      return
    }
    console.log('exportJson for loadExcel', exportJson)
    jsonData.value = exportJson // Keep if jsonData is used elsewhere

    isFunction(window?.luckysheet?.destroy) && window.luckysheet.destroy()

    window.luckysheet.create(getLuckysheetOptions({
      showinfobar: false,
      data: exportJson.sheets,
      title: exportJson.info.name,
      userInfo: exportJson.info.name.creator,
    }));
  })
}

const selectExcel = (evt) => {
  const value = selected.value
  const name = evt.target.options[evt.target.selectedIndex].innerText

  if (value == '') {
    return
  }
  isMaskShow.value = true

  LuckyExcel.transformExcelToLuckyByUrl(value, name, (exportJson, luckysheetfile) => {
    if (exportJson.sheets == null || exportJson.sheets.length == 0) {
      alert('Failed to read the content of the excel file, currently does not support xls files!')
      isMaskShow.value = false;
      return
    }
    console.log('exportJson for selectExcel', exportJson)
    jsonData.value = exportJson // Keep if jsonData is used elsewhere
    isMaskShow.value = false

    isFunction(window?.luckysheet?.destroy) && window.luckysheet.destroy()

    window.luckysheet.create(getLuckysheetOptions({
      showinfobar: false,
      data: exportJson.sheets,
      title: exportJson.info.name,
      userInfo: exportJson.info.name.creator,
    }));
  })
}

const downloadExcel = () => {
  // Ensure luckysheet global object and getAllSheets method are available
  if (window.luckysheet && typeof window.luckysheet.getAllSheets === 'function') {
    exportExcel(window.luckysheet.getAllSheets(), 'DownloadedSheet')
  } else {
    alert('Luckysheet is not initialized or getAllSheets is not available.');
  }
}

// !!! create luckysheet after mounted
onMounted(() => {
  // Initial creation with minimal data, AI function will be hooked
  window.luckysheet.create(getLuckysheetOptions({
    data: [{ name: "Sheet1", color: "", status: "1", order: "0", celldata: [], config: {} }], // Basic sheet
    title: "AIxCel Demo"
  }));
})
</script>

<style scoped>
#luckysheet {
  margin: 0px;
  padding: 0px;
  position: absolute;
  width: 100%;
  left: 0px;
  top: 30px; /* Adjust if your layout needs more/less space for the controls */
  bottom: 0px;
}

#uploadBtn {
  font-size: 16px;
}

#tip {
  position: absolute;
  z-index: 1000000;
  left: 0px;
  top: 0px;
  bottom: 0px;
  right: 0px;
  background: rgba(255, 255, 255, 0.8);
  text-align: center;
  font-size: 40px;
  align-items: center;
  justify-content: center;
  display: flex;
}
</style>
