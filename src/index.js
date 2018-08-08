const {ipcRenderer} = require('electron')
const docUtil = require('./word.js')

$('#docInit').on('click', function() {
  let options = {}
  Array.from(document.querySelectorAll('form input'))
    .forEach(item => {
      options[item.id] = item.value
    })
  docUtil.shengcDocx(options)
    .then(() => {
      ipcRenderer.send('test', 'ping')
    })
})
