import { btn_new, newFile } from './utils/dom.util'

function bindEventHandlers() {
  btn_new.onclick = newFile
}

function initApp() {
  bindEventHandlers()
}

initApp()