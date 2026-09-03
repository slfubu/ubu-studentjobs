(function (global) {
  'use strict';

  const config = Object.freeze({
    API_URL: 'https://script.google.com/macros/s/AKfycbw5VxGqIOhu1gUwrlWKnMRaIxBVBttSQ4wSyf__C-mfRzVjP9iJ0r8Fub7JDRq2o9kG/exec',
    REQUEST_TIMEOUT_MS: 45000,
    READ_TIMEOUT_MS: 8000,
    // Login ใช้ 25 วินาทีเป็น safety net เฉพาะกรณี Google Apps Script cold start
    // โดยปกติ warm-up + cache ควรตอบกลับภายในไม่กี่วินาที
    AUTH_TIMEOUT_MS: 25000,
    DEBUG_API: false
  });

  global.UBU_APP_CONFIG = config;
  global.APP_CONFIG = config;
})(window);
