import { btn_new, btn_open, btn_plus, btn_minus, newFile, showFileDialog, addRow, removeRow } from './utils/dom.util'

function bindEventHandlers() {
  btn_new.onclick = newFile
  btn_open.onclick = showFileDialog
  btn_plus.onclick = addRow
  btn_minus.onclick = removeRow
}

function initApp() {
  bindEventHandlers()
}

initApp()