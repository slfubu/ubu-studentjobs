/*
 * config.js
 * GitHub Pages -> Google Apps Script Web App
 *
 * สำคัญ:
 * - API_URL ต้องเป็น Deployment ปัจจุบันและลงท้ายด้วย /exec
 * - ถ้าสร้าง New deployment ใหม่ URL อาจเปลี่ยน ต้องนำ URL ใหม่มาแทนที่ด้านล่าง
 */
(function (global) {
  'use strict';

  const config = Object.freeze({
    API_URL: 'https://script.google.com/macros/s/AKfycbz4ykKdwE_2K-GEkwqJWnf1yftEwh0zwDe9Tc8gDXOsjjJXrvBTCtUOj080y2rCbMw2/exec',
    REQUEST_TIMEOUT_MS: 120000,
    DEBUG_API: false
  });

  // รองรับ bridge ทั้งรุ่นใหม่และรุ่นเก่า
  global.UBU_APP_CONFIG = config;
  global.APP_CONFIG = config;
})(window);
