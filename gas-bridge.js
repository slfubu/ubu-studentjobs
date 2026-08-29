/*
 * gas-bridge.js
 * Security-hardened bridge สำหรับ GitHub Pages -> Apps Script
 *
 * ไม่มี secret ในไฟล์นี้
 * clientId/requestId ใช้ช่วย rate-limit + anti-replay เท่านั้น
 */
(function (global) {
  'use strict';

  const CLIENT_ID_STORAGE_KEY = 'ubuSecurityClientId';

  function getConfig() {
    return global.UBU_APP_CONFIG || {};
  }

  function getApiUrl() {
    const url = String(getConfig().API_URL || '').trim();

    if (!url || url.includes('PASTE_YOUR_APPS_SCRIPT_WEB_APP_EXEC_URL_HERE')) {
      throw new Error(
        'ยังไม่ได้ตั้งค่า Apps Script API URL กรุณาแก้ไฟล์ config.js แล้วใส่ Web App URL ที่ลงท้ายด้วย /exec'
      );
    }

    if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec(?:\?.*)?$/i.test(url)) {
      throw new Error('Apps Script API URL ไม่ถูกต้อง');
    }

    return url;
  }

  function randomId() {
    if (global.crypto && typeof global.crypto.randomUUID === 'function') {
      return global.crypto.randomUUID().replace(/-/g, '');
    }

    if (global.crypto && typeof global.crypto.getRandomValues === 'function') {
      const bytes = new Uint8Array(24);
      global.crypto.getRandomValues(bytes);
      return Array.from(bytes, function (b) {
        return b.toString(16).padStart(2, '0');
      }).join('');
    }

    return (
      Date.now().toString(36) +
      Math.random().toString(36).slice(2) +
      Math.random().toString(36).slice(2)
    ).replace(/[^A-Za-z0-9_-]/g, '');
  }

  function getClientId() {
    try {
      let value = global.localStorage.getItem(CLIENT_ID_STORAGE_KEY);

      if (!/^[A-Za-z0-9_-]{16,120}$/.test(String(value || ''))) {
        value = randomId();
        global.localStorage.setItem(CLIENT_ID_STORAGE_KEY, value);
      }

      return value;
    } catch (_) {
      return randomId();
    }
  }

  async function invokeAppsScript(action, args) {
    const controller = new AbortController();
    const timeoutMs = Number(getConfig().REQUEST_TIMEOUT_MS) || 120000;
    const timeoutId = setTimeout(function () {
      controller.abort();
    }, timeoutMs);

    const request = {
      action: String(action || ''),
      args: Array.isArray(args) ? args : [],
      meta: {
        timestamp: Date.now(),
        requestId: randomId() + randomId(),
        clientId: getClientId(),
        origin: String(
          global.location && global.location.origin || ''
        )
      }
    };

    try {
      const response = await fetch(getApiUrl(), {
        method: 'POST',
        mode: 'cors',
        redirect: 'follow',
        credentials: 'omit',
        cache: 'no-store',
        referrerPolicy: 'no-referrer',
        headers: {
          'Content-Type': 'text/plain;charset=UTF-8'
        },
        body: JSON.stringify(request),
        signal: controller.signal
      });

      const raw = await response.text();

      let payload;
      try {
        payload = JSON.parse(raw);
      } catch (_) {
        throw new Error(
          'Apps Script ตอบกลับมาในรูปแบบไม่ถูกต้อง กรุณาตรวจสอบเวอร์ชันที่ Deploy'
        );
      }

      if (!payload || payload.ok !== true) {
        const message =
          payload && payload.error && payload.error.message
            ? String(payload.error.message)
            : 'เกิดข้อผิดพลาดในการประมวลผลคำขอ';

        throw new Error(message.slice(0, 300));
      }

      return payload.data;
    } catch (error) {
      if (error && error.name === 'AbortError') {
        throw new Error(
          'การเชื่อมต่อ Apps Script ใช้เวลานานเกินกำหนด กรุณาลองใหม่อีกครั้ง'
        );
      }

      if (error instanceof TypeError) {
        throw new Error(
          'ไม่สามารถเชื่อมต่อ Apps Script API ได้ กรุณาตรวจสอบ API URL และการ Deploy'
        );
      }

      throw error;
    } finally {
      clearTimeout(timeoutId);
    }
  }

  function createRunner(successHandler, failureHandler) {
    return new Proxy({}, {
      get: function (_target, property) {
        if (property === 'withSuccessHandler') {
          return function (handler) {
            return createRunner(handler, failureHandler);
          };
        }

        if (property === 'withFailureHandler') {
          return function (handler) {
            return createRunner(successHandler, handler);
          };
        }

        if (property === 'then') return undefined;

        return function () {
          const args = Array.prototype.slice.call(arguments);

          invokeAppsScript(property, args)
            .then(function (result) {
              if (typeof successHandler === 'function') {
                successHandler(result);
              }
            })
            .catch(function (error) {
              if (typeof failureHandler === 'function') {
                failureHandler(error);
              } else {
                console.error('API request failed');
              }
            });
        };
      }
    });
  }

  global.google = global.google || {};
  global.google.script = global.google.script || {};
  global.google.script.run = createRunner(null, null);

  global.UBUApi = Object.freeze({
    call: function (method) {
      const args = Array.prototype.slice.call(arguments, 1);
      return invokeAppsScript(method, args);
    }
  });
})(window);
