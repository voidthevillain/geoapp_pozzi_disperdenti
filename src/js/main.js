import { btn_new, btn_open, btn_save, btn_print, btn_calculate, newFile, showFileDialog, saveFile, calculate, printFile } from './utils/dom.util'

function bindEventHandlers() {
  btn_new.onclick = newFile
  btn_open.onclick = showFileDialog
  btn_save.onclick = saveFile
  btn_calculate.onclick = calculate
  btn_print.onclick = printFile
}

function initApp() {
  bindEventHandlers()
}

initApp()