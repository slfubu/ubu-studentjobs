/*
 * gas-bridge.js
 * Compatibility bridge: google.script.run -> Apps Script Web App HTTP POST
 * Compatible with Code.gs security meta:
 *   timestamp, requestId, clientId
 */
(function (global) {
  'use strict';

  const CLIENT_ID_KEY = 'UBU_API_CLIENT_ID_V1';

  function getConfig() {
    return global.UBU_APP_CONFIG || global.APP_CONFIG || {};
  }

  function getApiUrl() {
    const url = String(getConfig().API_URL || '').trim();

    if (!url || url.includes('PASTE_YOUR_APPS_SCRIPT_WEB_APP_EXEC_URL_HERE')) {
      throw new Error(
        'ยังไม่ได้ตั้งค่า Apps Script API URL กรุณาแก้ config.js และใส่ Web App URL ที่ลงท้ายด้วย /exec'
      );
    }

    if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec(?:\?.*)?$/i.test(url)) {
      throw new Error('API_URL ไม่ใช่ Google Apps Script Web App URL แบบ /exec');
    }

    return url;
  }

  function randomId() {
    try {
      if (global.crypto && typeof global.crypto.randomUUID === 'function') {
        return global.crypto.randomUUID().replace(/-/g, '');
      }

      if (global.crypto && typeof global.crypto.getRandomValues === 'function') {
        const bytes = new Uint8Array(20);
        global.crypto.getRandomValues(bytes);
        return Array.from(bytes, b => b.toString(16).padStart(2, '0')).join('');
      }
    } catch (_) {}

    return (
      Date.now().toString(36) +
      Math.random().toString(36).slice(2) +
      Math.random().toString(36).slice(2)
    ).replace(/[^A-Za-z0-9_-]/g, '').padEnd(20, '0').slice(0, 64);
  }

  function getClientId() {
    try {
      let id = localStorage.getItem(CLIENT_ID_KEY);
      if (/^[A-Za-z0-9_-]{16,120}$/.test(String(id || ''))) return id;

      id = randomId();
      localStorage.setItem(CLIENT_ID_KEY, id);
      return id;
    } catch (_) {
      return randomId();
    }
  }

  function getAuthToken() {
    try {
      const direct =
        sessionStorage.getItem('sessionToken') ||
        localStorage.getItem('sessionToken') ||
        '';
      if (direct) return String(direct);

      // รองรับ staff session ของ dashboard โดยไม่บังคับให้ใช้
      const staffRaw = sessionStorage.getItem('ubuStaffSession');
      if (staffRaw) {
        const staff = JSON.parse(staffRaw);
        if (staff && staff.token) return String(staff.token);
      }
    } catch (_) {}

    return '';
  }

  function buildMeta() {
    return {
      timestamp: Date.now(),
      requestId: randomId(),
      clientId: getClientId()
    };
  }

  async function invokeAppsScript(action, args) {
    const controller = new AbortController();
    const config = getConfig();
    const timeoutMs = Number(config.REQUEST_TIMEOUT_MS) || 120000;
    const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
    const url = getApiUrl();

    try {
      const response = await fetch(url, {
        method: 'POST',
        redirect: 'follow',
        credentials: 'omit',
        cache: 'no-store',
        headers: {
          // text/plain keeps the request "simple" and avoids OPTIONS preflight.
          'Content-Type': 'text/plain;charset=UTF-8'
        },
        body: JSON.stringify({
          action: String(action || ''),
          args: Array.isArray(args) ? args : [],
          // ใช้กับ Api.gs รุ่นเดิมที่มีระบบ user session
          authToken: getAuthToken(),
          // ใช้กับ Code.gs security gateway รุ่นใหม่
          meta: buildMeta()
        }),
        signal: controller.signal
      });

      const raw = await response.text();

      let payload;
      try {
        payload = JSON.parse(raw);
      } catch (_) {
        const preview = String(raw || '').replace(/\s+/g, ' ').slice(0, 180);

        if (/accounts\.google\.com|Sign in|ลงชื่อเข้าใช้/i.test(preview)) {
          throw new Error(
            'Apps Script Web App ยังไม่เปิดสิทธิ์สาธารณะ กรุณา Deploy โดย Execute as = Me และ Who has access = Anyone'
          );
        }

        throw new Error(
          'Apps Script ตอบกลับมาไม่ใช่ JSON กรุณาตรวจสอบว่า config.js ใช้ URL /exec ของ Deployment ปัจจุบัน และ Deploy Code.gs รุ่นล่าสุดแล้ว'
        );
      }

      if (!payload || payload.ok !== true) {
        const message =
          payload && payload.error && payload.error.message
            ? String(payload.error.message)
            : 'เกิดข้อผิดพลาดในการประมวลผลคำขอ';
        throw new Error(message);
      }

      if (config.DEBUG_API) {
        console.debug('[UBU API]', action, payload.data);
      }

      if (Object.prototype.hasOwnProperty.call(payload, 'data')) {
        return payload.data;
      }
      if (Object.prototype.hasOwnProperty.call(payload, 'result')) {
        return payload.result;
      }
      return null;
    } catch (error) {
      if (error && error.name === 'AbortError') {
        throw new Error('การเชื่อมต่อ Apps Script ใช้เวลานานเกินกำหนด กรุณาลองใหม่อีกครั้ง');
      }

      // Browser fetch throws TypeError for CORS/network/redirect-to-login failures.
      if (error instanceof TypeError) {
        throw new Error(
          'เบราว์เซอร์เชื่อมต่อ Apps Script Web App ไม่ได้ กรุณาตรวจสอบว่า URL เป็น /exec ของ Deployment ปัจจุบัน และตั้ง Who has access = Anyone'
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
        if (property === 'then') return undefined;

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

        return function () {
          const args = Array.prototype.slice.call(arguments);

          invokeAppsScript(String(property), args)
            .then(function (result) {
              if (typeof successHandler === 'function') {
                successHandler(result);
              }
            })
            .catch(function (error) {
              if (typeof failureHandler === 'function') {
                failureHandler(error);
              } else {
                console.error('[Apps Script API]', property, error);
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
    },
    getUrl: getApiUrl
  });
})(window);
