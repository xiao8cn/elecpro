const {ipcRenderer} = require('electron')
const docUtil = require('./word.js')

document.getElementById('docInit')
  .addEventListener('click', () => {
    let contractNo = document.getElementById('contractNo').value
    let purchaser = document.getElementById('purchaser').value
    let seller = document.getElementById('seller').value
    let options = {
      contractNo,
      seller,
      purchaser
    }
    docUtil.shengcDocx(options)
      .then(() => {
        ipcRenderer.send('test', 'ping')
      })
  })
