/**
 * Converts degrees to radians
 * @param {number} deg value in degrees
 * @returns {number} value in radians
 */
exports.rad = function(deg) {
  return deg * Math.PI / 180
}

/**
 * Converts radians to degrees
 * @param {number} rad value in radians
 * @returns {number} value in degrees
 */
exports.deg = function(rad) {
  return rad * 180 /  Math.PI
}

/**
 * Calculates the sine of an angle
 * @param {number} angle angle (in radians)
 * @returns {number} sine of the angle
 */
exports.sen = function(angle) {
  return Math.sin(angle)
}

/** 
 * Calculates the cosine of an angle
 * @param {number} angle angle (in radians)
 * @returns {number} cosine of the angle
 */
exports.cos = function(angle) {
  return Math.cos(angle)
}


/**
 * Calculates the tangent of an angle
 * @param {number} angle angle (in radians)
 * @returns {number} tangent of the angle
 */
exports.tan = function(angle) {
  return Math.tan(angle)
}