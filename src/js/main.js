import { btn_new, btn_open, btn_save, btn_print, btn_plus, btn_minus, btn_calculate, newFile, showFileDialog, saveFile, addRow, removeRow, calculate, printFile } from './utils/dom.util'

function bindEventHandlers() {
  btn_new.onclick = newFile
  btn_open.onclick = showFileDialog
  btn_save.onclick = saveFile
  btn_plus.onclick = addRow
  btn_minus.onclick = removeRow
  btn_calculate.onclick = calculate
  btn_print.onclick = printFile
}

function initApp() {
  bindEventHandlers()
}

initApp()