const {ipcRenderer} = require('electron')
const docUtil = require('./word.js')
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
    dom.append(`
      <div class="col form-item">
        <label for="${item.type}">${item.label}</label>
        <input type="${item.type}" id="${item.key}" placeholder="${item.placeholder}">
      </div>
    `)
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
  console.log(options)
  docUtil.shengcDocx(options)
    .then(() => {
      ipcRenderer.send('test', 'ping')
    })
})
