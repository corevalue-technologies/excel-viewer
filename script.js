let fileInput = document.getElementById('fileInput')
let viewer = document.getElementById('viewer')
let workBook = null
let excelGrid = null
let activeSheet = ''
let sheets = []
let excelButtons = null
let buttons = []

// <button v-for="(sheet, index) in sheets" :key="index" @click="showSheet(sheet)" :class="{'active': activeSheet === sheet}">{{sheet}}</button>

function getFile (e) {
    var reader = new FileReader()
    reader.readAsBinaryString(e.target.files[0])
    reader.onload = function () {
      showExcel(reader.result)
    }
    reader.onerror = function (error) {
      console.log('error', error)
    }
}

fileInput.addEventListener('change', getFile)


function showSheet (el) {
    let buttons = document.querySelectorAll('button')
    console.log(buttons)
    buttons.forEach(button => {
        console.log(button)
        button.classList.remove('active')
    })
    el.classList.add('active')
    var workSheet = workBook.Sheets[el.innerText]
    excelGrid.innerHTML = XLSX.utils.sheet_to_html(workSheet)
    activeSheet = el.innerText
}

function clearAll () {
    viewer.innerHTML = ""
    workBook = null
    excelGrid = null
    sheets = []
    excelButtons = null
    buttons = []
}

function showExcel (data) {
    clearAll()
    workBook = XLSX.read(data, {type: 'binary'})
    console.log(workBook)
    sheets = workBook.SheetNames
    workBook.SheetNames.forEach(function (sheetName) {
      // Get headers.
      var headers = []
      var sheet = workBook.Sheets[sheetName]
      var range = XLSX.utils.decode_range(sheet['!ref'])
      var C = range.s.r
      var R = range.s.r
      /* start in the first row */
      /* walk every column in the range */
      for (C = range.s.c; C <= range.e.c; ++C) {
        var cell = sheet[XLSX.utils.encode_cell({c: C, r: R})]
        /* find the cell in the first row */
        var hdr = 'NIPUN'
        if (cell && cell.t) {
          hdr = XLSX.utils.format_cell(cell)
        }
        headers.push(hdr)
      }
      // For each sheets, convert to json.
      var roa = XLSX.utils.sheet_to_json(workBook.Sheets[sheetName])
      if (roa.length > 0) {
        roa.forEach(function (row) {
          // Set empty cell to ''.
          headers.forEach(function (hd) {
            if (row[hd] === undefined) {
              row[hd] = ''
            }
          })
        })
      }
    })
    excelGrid = document.createElement('table')
    excelGrid.classList.add('table')
    excelGrid.classList.add('table-bordered')
    excelGrid.classList.add('table-responsive')
    excelGrid.classList.add('excel-table')
    excelButtons = document.createElement('div')
    excelButtons.classList.add('excelButtons')
    for (var i = 0; i < sheets.length; i++) {
      let button = document.createElement('button')
      button.classList.add('sheetBtn')
      button.innerText = sheets[i]
      button.addEventListener('click', (e) => {
        showSheet(e.target)
      })
      excelButtons.appendChild(button)
      buttons.push(button)
    }
    let container = document.createElement('div')
    container.classList.add('excel-container')
    container.appendChild(excelGrid)
    viewer.innerHTML = ""
    viewer.appendChild(container)
    viewer.appendChild(excelButtons)
    // self.excelGrid = canvasDatagrid({
    //   parentNode: document.getElementById('pdf-viewer'),
    //   data: []
    // })
    // self.excelGrid.style.width = '100%'
    // self.excelGrid.style.height = '100%'

    // self.excelGrid.style.gridBackgroundColor = 'white'
    // self.excelGrid.style.cellFont = '14px sans-serif'
    showSheet(buttons[0])
  }