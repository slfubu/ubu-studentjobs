(function (global) {
  'use strict';

  const config = Object.freeze({
    API_URL: 'https://script.google.com/macros/s/AKfycbz4ykKdwE_2K-GEkwqJWnf1yftEwh0zwDe9Tc8gDXOsjjJXrvBTCtUOj080y2rCbMw2/exec',
    REQUEST_TIMEOUT_MS: 60000,
    DEBUG_API: false
  });

  global.UBU_APP_CONFIG = config;
  global.APP_CONFIG = config;
})(window);

