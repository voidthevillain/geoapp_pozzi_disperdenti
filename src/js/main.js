import { btn_new, btn_open, btn_save, btn_plus, btn_minus, btn_calculate, newFile, showFileDialog, saveFile, addRow, removeRow, calculate } from './utils/dom.util'

function bindEventHandlers() {
  btn_new.onclick = newFile
  btn_open.onclick = showFileDialog
  btn_save.onclick = saveFile
  btn_plus.onclick = addRow
  btn_minus.onclick = removeRow
  btn_calculate.onclick = calculate
}

function initApp() {
  bindEventHandlers()
}

initApp()