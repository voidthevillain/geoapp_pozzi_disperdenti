import parse from './parse.util'
import { saveAs } from 'file-saver'

// Buttons
export const btn_new = document.querySelector('#btn-new')
export const btn_open = document.querySelector('#btn-open')
export const btn_save = document.querySelector('#btn-save')
export const btn_print = document.querySelector('#btn-print')
export const btn_help = document.querySelector('#btn-help')

export const btn_minus = document.querySelector('#btn-minus')
export const btn_plus = document.querySelector('#btn-plus')

export const btn_calculate = document.querySelector('#btn-calculate')

// File
const input_file = document.querySelector('#file-input')

// Inputs
const input_k = document.querySelector('#k')
const input_d = document.querySelector('#d')
const input_n = document.querySelector('#n')
const input_as = document.querySelector('#as')
const input_phi = document.querySelector('#phi')
const input_tc = document.querySelector('#tc')
const input_dp = document.querySelector('#dp')
const input_dt = document.querySelector('#dt')
const input_a = document.querySelector('#a')

const inputs = document.querySelectorAll('.input-value')

// Outputs
const output_hwmax = document.querySelector('#hwmax')

// Table
const table = document.querySelector('#table')

// Modals
// Save
const btn_save_modal = document.querySelector('#btn-save-modal')
const btn_close_save_modal = document.querySelector('#btn-close-save-modal')
const btn_confirm_save_modal = document.querySelector('#btn-confirm-save-modal')
const btn_ext_txt = document.querySelector('#btn-ext-txt')
const btn_ext_gap = document.querySelector('#btn-ext-gap')
const file_name = document.querySelector('#file-name')
const file_ext = document.querySelector('#file-ext')

// Alert
const btn_toggle_alert = document.querySelector('#btn-toggle-alert')
const btn_go_to_field = document.querySelector('#btn-go-to-field')
const alert_body = document.querySelector('#alert-modal-body')

// Data
export let K, D, N, As, phi, tc, DP, DT, a
export let T = []
export let Qp = []
export let hwmax

function initData() {
  K = parse(input_k.value)
  D = parse(input_d.value)
  N = parse(input_n.value)
  As = parse(input_as.value)
  phi = parse(input_phi.value)
  tc = parse(input_tc.value)
  DP = parse(input_dp.value)
  DT = parse(input_dt.value)
  a = parse(input_a.value)

  for (let i = 1; i < table.rows.length; i++) {
    T.push(table.rows[i].cells[1].innerText)
    Qp.push(table.rows[i].cells[2].innerText)
  }
}

// Event handlers
export function newFile() {
  return (() => {
    inputs.forEach(e => e.value = '')

    for (let i = 1; i < table.rows.length; i++) {
      table.rows[i].cells[1].innerText = '-'
      table.rows[i].cells[2].innerText = '-'
    }

    output_hwmax.value = ''
  })()
}

export function showFileDialog() {
  return (() => {
    input_file.click()

    input_file.onchange = e => {
      openFile(e)
    }
  })()
}

export function saveFile() {
  return (() => {
    btn_save_modal.click()

    btn_close_save_modal.onclick = () => {
      return
    }

    btn_ext_txt.onclick = () => {
      file_ext.innerHTML = '.txt'
    }

    btn_ext_gap.onclick = () => {
      file_ext.innerHTML = '.gap'
    }

    btn_confirm_save_modal.onclick = () => {
      let fileName = file_name.value
      let fileExt = file_ext.innerHTML

      let str = ''

      if (fileName === '') {
        file_name.focus()
      } else {
        inputs.forEach(e => {
          str += e.value + ';'
        })
        
        for (let i = 1; i < table.rows.length; i++) {
          str += table.rows[i].cells[1].innerText + ';'
          str += table.rows[i].cells[2].innerText + ';'
        }
      }

      let blob = new Blob([str], {
        type: 'text/plain; charset=utf-8'
      })
      saveAs(blob, fileName + fileExt)
    }
  })()
}

export function addRow() {
  return (() => {
    let rows = parse(table.rows[table.rows.length - 1].cells[0].innerText)

    let t_body = table.querySelector('tbody')
    t_body.appendChild(document.createElement('tr'))
    t_body.lastChild.classList.add('text-center')
    t_body.lastChild.innerHTML = `<td class="border-right">${rows + 1}</td><td class="border-right" contenteditable="true"></td><td contenteditable="true"></td>`
  })()
}

export function removeRow() {
  return (() => {
    let rows = parse(table.rows[table.rows.length - 1].cells[0].innerText)

    if (rows >= 2) {
      table.rows[table.rows.length - 1].remove()
    }
  })()
}