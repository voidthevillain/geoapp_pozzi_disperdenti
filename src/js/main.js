import { btn_new, btn_open, btn_plus, newFile, showFileDialog, addRow } from './utils/dom.util'

function bindEventHandlers() {
  btn_new.onclick = newFile
  btn_open.onclick = showFileDialog
  btn_plus.onclick = addRow
}

function initApp() {
  bindEventHandlers()
}

initApp()