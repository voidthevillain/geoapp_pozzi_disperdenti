import '@fortawesome/fontawesome-free/js/fontawesome.min.js'
import '@fortawesome/fontawesome-free/js/solid'
import '../../css/bootstrap.min.css'
import '../../css/normalize.css'
import '../../css/main.css'

// Overlay (spinner)
// const overlay = document.querySelector('#overlay')
// window.addEventListener('load', function () {
//   setTimeout(() => {
//     overlay.parentNode.removeChild(overlay)
//   }, 500)
// })

// Tooltip
$(function () {
  $('[data-toggle="tooltip"]').tooltip()
})