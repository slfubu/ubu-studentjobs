(function (global) {
  'use strict';

  const config = Object.freeze({
    API_URL: 'https://script.google.com/macros/s/AKfycbw5VxGqIOhu1gUwrlWKnMRaIxBVBttSQ4wSyf__C-mfRzVjP9iJ0r8Fub7JDRq2o9kG/exec',
    REQUEST_TIMEOUT_MS: 45000,
    READ_TIMEOUT_MS: 25000,
    // FAST V4: คำขออ่านข้อมูลไม่ตัดที่ 8 วินาทีอีกต่อไป เพราะ Apps Script cold start
    // อาจเกิน 8 วินาทีได้ แม้เครือข่ายปกติ โดยฝั่ง Server ใช้ cache ลดเวลาในคำขอถัดไป
    // Login ใช้ 25 วินาทีเป็น safety net เฉพาะกรณี Google Apps Script cold start
    AUTH_TIMEOUT_MS: 25000,
    DEBUG_API: false
  });

  global.UBU_APP_CONFIG = config;
  global.APP_CONFIG = config;
})(window);
