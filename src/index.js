const {ipcRenderer} = require('electron')
const docUtil = require('./word.js')

document.getElementById('docInit')
  .addEventListener('click', () => {
    let contractNo = document.getElementById('contractNo').value
    console.log(contractNo)
    let options = {
      contractNo
    }
    docUtil.shengcDocx(options)
      .then(() => {
        ipcRenderer.send('test', 'ping')
      })
  })
