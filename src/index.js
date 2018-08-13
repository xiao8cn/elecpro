const {ipcRenderer} = require('electron')
const docUtil = require('./word.js')
const pptUtil = require('./ppt.js')
const xlsxUtil = require('./xlsx.js')
const path = require('path')
const fs = require('fs')

let config = null

window.onload = function () {
  config = JSON.parse(fs.readFileSync(path.join(path.dirname(__dirname), "src/config.json"), "utf-8"))
  init()
}

/**
 * 页面初始化
 */
function init() {
  let forms = config.forms
  let dom = $('.forms .row')
  forms.forEach(item => {
    if (item.type === "text" || item.type === "number" || item.type === "date") {
      dom.append(`
        <div class="col-3 form-item">
          <label for="${item.type}">${item.label}</label>
          <input value="${item.default_value}" type="${item.type}" id="${item.key}" placeholder="${item.placeholder}">
        </div>
      `)
    } else if (item.type === "select") {
      let optionsHtml = ''
      item.options.forEach(op => {
        optionsHtml += `<option value="${op.value}">${op.label}</option>`
      })
      dom.append(`
        <div class="col-3 form-group">
          <label for="exampleFormControlSelect1">${item.label}</label>
          <select class="form-control" id="${item.key}">
            ${optionsHtml}
          </select>
        </div>
      `)
    }
  })
}

/**
 * 生成合同 click 事件
 */
$('#docInit').on('click', function() {
  let options = {}
  Array.from(document.querySelectorAll('form input'))
    .forEach(item => {
      options[item.id] = item.value
    })
    Array.from(document.querySelectorAll('form select'))
    .forEach(item => {
      options[item.id] = item.value
    })
  docUtil.shengcDocx(options)
    .then(() => {
      ipcRenderer.send('test', 'ping')
    })
})

/**
 * 生成ppt 事件
 */
$('#pptInit').on('click', function() {
  pptUtil.shengcPpt()
    .then(() => {
      ipcRenderer.send('ppt')
    })
})

/**
 * 生成xlsx 事件
 */
$('#xlsxInit').on('click', function() {
  xlsxUtil.shengcXlsx()
    .then(() => {
      ipcRenderer.send('xlsx')
    })
})