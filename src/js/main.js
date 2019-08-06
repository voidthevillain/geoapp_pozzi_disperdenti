import { btn_new, btn_open, newFile, showFileDialog } from './utils/dom.util'

function bindEventHandlers() {
  btn_new.onclick = newFile
  btn_open.onclick = showFileDialog
}

function initApp() {
  bindEventHandlers()
}

initApp()