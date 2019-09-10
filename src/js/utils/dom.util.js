import parse from './parse.util'
import {
  saveAs
} from 'file-saver'
import {
  createDocument
} from './print.util'

// Buttons
export const btn_new = document.querySelector('#btn-new')
export const btn_open = document.querySelector('#btn-open')
export const btn_save = document.querySelector('#btn-save')
export const btn_print = document.querySelector('#btn-print')
export const btn_help = document.querySelector('#btn-help')
export const btn_calculate = document.querySelector('#btn-calculate')

// File
const input_file = document.querySelector('#file-input')

// Inputs
export const input_k = document.querySelector('#k')
export const input_d = document.querySelector('#d')
export const input_np = document.querySelector('#np')
export const input_as = document.querySelector('#as')
export const input_phi = document.querySelector('#phi')
export const input_tc = document.querySelector('#tc')
export const input_dp = document.querySelector('#dp')
export const input_dt = document.querySelector('#dt')
export const input_a = document.querySelector('#a')
export const input_n = document.querySelector('#n')

const inputs = document.querySelectorAll('.input-value')

// Outputs
const output_hwmax = document.querySelector('#hwmax')

// Table
export const table = document.querySelector('#table')

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
export let K, D, Np, As, phi, tc, DP, DT, a, n
export let T = []
export let Qp = []
export let Qf = []
export let DW = []
export let hw = []
export let hwmax


function initData() {
  K = parse(input_k.value)
  D = parse(input_d.value)
  Np = parse(input_np.value)
  As = parse(input_as.value)
  phi = parse(input_phi.value)
  tc = parse(input_tc.value)
  DP = parse(input_dp.value)
  DT = parse(input_dt.value)
  a = parse(input_a.value)
  n = parse(input_n.value)
}

// Event handlers
export function newFile() {
  return (() => {
    inputs.forEach(e => e.value = '')

    table.querySelector('thead').innerHTML = ``
    table.querySelector('tbody').innerHTML = ``

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

function openFile(e) {
  return (e => {
    let file = e.target.files[0]

    if (!file) return

    let reader = new FileReader()
    reader.onload = e => {
      let contents = e.target.result
      loadInput(contents)
    }
    reader.readAsText(file)

    function loadInput(contents) {
      let arr = contents.split(';')
      arr.pop()

      for (let i = 0; i <= 9; i++) {
        inputs[i].value = arr[i]
      }
    }
  })(e)
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
          str += table.rows[i].cells[0].innerText + ';'
          str += table.rows[i].cells[1].innerText + ';'
        }
      }

      let blob = new Blob([str], {
        type: 'text/plain; charset=utf-8'
      })
      saveAs(blob, fileName + fileExt)
    }
  })()
}

export function calculate() {
  return (() => {
    initData()

    let Af // Area efficace di drenaggio
    let Ap // Area effettiva pozzetto
    let Dw1
    let Qf = []
    let DW = []
    let hw = []
    let hwmax = 0

    let qc
    let t_length

    qc = As * phi * a * Math.pow(tc, n - 1) / 1000
    t_length = DP / DT

    T[0] = 0
    Qp[0] = 0

    table.querySelector('thead').innerHTML = `
    <th class="border-top-0 align-middle text-center" scope="scol">t<br>[h]</th>
    <th class="border-top-0 align-middle text-center" scope="col">Q<sub>p</sub><br>[m<sup>3</sup>/h]
    <th class="border-top-0 align-middle text-center" scope="col">Q<sub>f</sub><br>[m<sup>3</sup>/h]
    <th class="border-top-0 align-middle text-center" scope="col">ΔW<br>[m<sup>3</sup>]
    <th class="border-top-0 align-middle text-center" scope="col">h<sub>w</sub><br>[m]
    `
    table.querySelector('tbody').innerHTML = ``

    for (let i = 1; i <= t_length; i++) {
      T[i] = Math.round((T[i - 1] + DT) * 10) / 10


      if (T[i] < tc) {
        Qp[i] = qc * (T[i] / tc) * Math.pow(2.71, (1 - T[i] / tc))
      } else {
        Qp[i] = qc * Math.pow(2.71, -(T[i] / tc - 1))
      }

      // table.querySelector('tbody').innerHTML += `<tr class="text-center"><td class="border-right">${T[i].toFixed(2)}</td><td>${Qp[i].toFixed(2)}</td></tr>`
    }

    (function (As, phi, tc, DP, DT, a) {
      Qf[0] = 0
      DW[0] = 0
      hw[0] = 0

      Ap = Np * Math.PI * Math.pow(D, 2) / 4

      for (let i = 1; i < T.length; i++) {
        Af = Math.PI / 4 * (Math.pow((D + hw[i - 1]), 2) - Math.pow(D, 2))
        Qf[i] = 3600 * K * Af

        if (i == 1) {
          Dw1 = ((Qp[i - 1] + Qp[i]) - (Qf[i - 1] + Qf[i])) * (T[i] - T[i - 1])
        } else {
          Dw1 = 0.5 * ((Qp[i - 1] + Qp[i]) - (Qf[i - 1] + Qf[i])) * (T[i] - T[i - 1])
        }

        DW[i] = DW[i - 1] + Dw1
        hw[i] = DW[i] / Ap

        if (hw[i] > hwmax) {
          hwmax = hw[i]
        }
        table.querySelector('tbody').innerHTML += `
      <tr class="text-center">
      <td class="border-right">${T[i].toFixed(2)}</td>
      <td class="border-right">${Qp[i].toFixed(2)}</td>
      <td class="border-right">${Qf[i].toFixed(2)}</td>
      <td class="border-right">${DW[i].toFixed(2)}</td>
      <td>${hw[i].toFixed(2)}</td>
      </tr>`
      }
    })(As, phi, tc, DP, DT, a)

    output_hwmax.value = hwmax.toFixed(2)
  })()
}

export function printFile() {
  return (() => {
    createDocument()
  })()
}