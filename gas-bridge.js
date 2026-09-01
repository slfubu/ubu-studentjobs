(function (global) {
  'use strict';

  const CLIENT_ID_KEY = 'UBU_API_CLIENT_ID_V1';
  const READ_ACTIONS = new Set([
    'getJobs','getPublicContents','getApplicationStatus','getLegacyApplicantByIdCard',
    'getStaffSession','getStaffDashboard','getAdminApplicantPage','getAdminCivilRegistryIds',
    'getStaffAccountManagement','getAdminPublicContents','getAdminQualificationApplicants',
    'getAdminDepartmentForwarding','getAdminDepartmentResults','getDepartmentApplicants',
    'getStaffApplicantDetail','getStaffApplicantPhotoBatch'
  ]);
  const AUTH_ACTIONS = new Set(['staffLogin','getStaffSession']);
  const RETRY_ACTIONS = new Set([...READ_ACTIONS, 'staffLogin']);
  const inflightReads = new Map();
  const readResponseCache = new Map();
  const READ_CACHE_TTL_MS = Object.freeze({
    getJobs: 60000,
    getPublicContents: 30000,
    getStaffSession: 5000,
    getStaffDashboard: 30000,
    getAdminApplicantPage: 12000,
    getAdminCivilRegistryIds: 5000,
    getStaffAccountManagement: 30000,
    getAdminPublicContents: 30000,
    getAdminQualificationApplicants: 12000,
    getAdminDepartmentForwarding: 12000,
    getAdminDepartmentResults: 12000,
    getDepartmentApplicants: 12000,
    getStaffApplicantDetail: 5000,
    getStaffApplicantPhotoBatch: 5000
  });

  function readCacheKey(method, args) {
    try { return method + '|' + JSON.stringify(Array.isArray(args) ? args : []); }
    catch (_) { return method; }
  }

  function getCachedRead(method, args) {
    const ttl = Number(READ_CACHE_TTL_MS[method] || 0);
    if (!ttl) return null;
    const key = readCacheKey(method, args);
    const hit = readResponseCache.get(key);
    if (!hit) return null;
    if (Date.now() - hit.at > ttl) {
      readResponseCache.delete(key);
      return null;
    }
    return hit.value;
  }

  function putCachedRead(method, args, value) {
    const ttl = Number(READ_CACHE_TTL_MS[method] || 0);
    if (!ttl) return;
    const key = readCacheKey(method, args);
    readResponseCache.set(key, { at: Date.now(), value: value });
    if (readResponseCache.size > 80) {
      const firstKey = readResponseCache.keys().next().value;
      if (firstKey) readResponseCache.delete(firstKey);
    }
  }

  function clearReadCache() {
    readResponseCache.clear();
  }

  function getConfig() {
    return global.UBU_APP_CONFIG || global.APP_CONFIG || {};
  }

  function getApiUrl() {
    const url = String(getConfig().API_URL || '').trim();
    if (!url || url.includes('PASTE_CURRENT_APPS_SCRIPT_WEB_APP_EXEC_URL_HERE') || url.includes('PASTE_YOUR_APPS_SCRIPT_WEB_APP_EXEC_URL_HERE')) {
      throw new Error('ยังไม่ได้ตั้งค่า Apps Script Web App URL ปัจจุบัน กรุณานำ URL จาก Deploy > Manage deployments ที่ลงท้ายด้วย /exec มาใส่ใน config.js');
    }
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec(?:\?.*)?$/i.test(url)) {
      throw new Error('API_URL ไม่ถูกต้อง ต้องเป็น Apps Script Web App URL ปัจจุบันที่ลงท้ายด้วย /exec');
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
    return (Date.now().toString(36)+Math.random().toString(36).slice(2)+Math.random().toString(36).slice(2))
      .replace(/[^A-Za-z0-9_-]/g,'').padEnd(20,'0').slice(0,64);
  }

  function getClientId() {
    try {
      let id = localStorage.getItem(CLIENT_ID_KEY);
      if (/^[A-Za-z0-9_-]{16,120}$/.test(String(id || ''))) return id;
      id = randomId();
      localStorage.setItem(CLIENT_ID_KEY, id);
      return id;
    } catch (_) { return randomId(); }
  }

  function getAuthToken() {
    try {
      const direct = sessionStorage.getItem('sessionToken') || localStorage.getItem('sessionToken') || '';
      if (direct) return String(direct);
      const raw = sessionStorage.getItem('ubuStaffSession');
      if (raw) {
        const session = JSON.parse(raw);
        if (session && session.token) return String(session.token);
      }
    } catch (_) {}
    return '';
  }

  function buildMeta() {
    return { timestamp: Date.now(), requestId: randomId(), clientId: getClientId() };
  }

  function timeoutFor(action) {
    const config = getConfig();
    const name = String(action || '');
    if (AUTH_ACTIONS.has(name)) return Math.max(5000, Number(config.AUTH_TIMEOUT_MS) || 15000);
    if (['uploadApplicationFiles','savePublicContent','processQualificationImport'].includes(name)) {
      return Math.max(120000, Number(config.REQUEST_TIMEOUT_MS) || 45000);
    }
    if (name === 'submitApplication') return Math.max(90000, Number(config.REQUEST_TIMEOUT_MS) || 45000);
    if (READ_ACTIONS.has(name)) return Math.max(10000, Number(config.READ_TIMEOUT_MS) || 30000);
    return Math.max(10000, Number(config.REQUEST_TIMEOUT_MS) || 45000);
  }

  function httpError(status) {
    if (status === 404 || status === 410) {
      return new Error('ไม่พบ Apps Script Deployment (HTTP 404) กรุณาสร้าง/อัปเดต Web app deployment แล้วนำ URL /exec ปัจจุบันมาใส่ใน config.js');
    }
    if (status === 401 || status === 403) {
      return new Error(`Apps Script ไม่อนุญาตให้เข้าถึง (HTTP ${status}) กรุณาตรวจ Execute as = Me และ Who has access = Anyone`);
    }
    if (status === 429) return new Error('ระบบถูกเรียกใช้งานถี่เกินไป กรุณาลองใหม่อีกครั้ง');
    return new Error(`Apps Script ตอบกลับ HTTP ${status} กรุณาตรวจสอบ Deployment`);
  }

  async function invokeAppsScriptOnce(action, args) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeoutFor(action));
    const url = getApiUrl();
    try {
      const response = await fetch(url, {
        method: 'POST', redirect: 'follow', credentials: 'omit', cache: 'no-store',
        headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
        body: JSON.stringify({ action:String(action||''), args:Array.isArray(args)?args:[], authToken:getAuthToken(), meta:buildMeta() }),
        signal: controller.signal
      });
      const raw = await response.text();
      if (!response.ok) throw httpError(response.status);

      let payload;
      try { payload = JSON.parse(raw); }
      catch (_) {
        const preview = String(raw || '').replace(/\s+/g,' ').slice(0,220);
        if (/accounts\.google\.com|Sign in|ลงชื่อเข้าใช้/i.test(preview)) {
          throw new Error('Apps Script Web App ยังไม่เปิดสิทธิ์สาธารณะ กรุณา Deploy โดย Execute as = Me และ Who has access = Anyone');
        }
        throw new Error('Apps Script ตอบกลับมาไม่ใช่ JSON กรุณาตรวจว่า URL เป็น /exec ของ Deployment ปัจจุบัน และ Code.gs รุ่นล่าสุดถูก Deploy แล้ว');
      }

      if (!payload || payload.ok !== true) {
        const message = payload && payload.error && payload.error.message ? String(payload.error.message) : 'เกิดข้อผิดพลาดในการประมวลผลคำขอ';
        throw new Error(message);
      }
      if (getConfig().DEBUG_API) console.debug('[UBU API]', action, payload.data);
      if (Object.prototype.hasOwnProperty.call(payload,'data')) return payload.data;
      if (Object.prototype.hasOwnProperty.call(payload,'result')) return payload.result;
      return null;
    } catch (error) {
      if (error && error.name === 'AbortError') throw new Error('เชื่อมต่อ Apps Script เกินเวลาที่กำหนด กรุณาตรวจอินเทอร์เน็ตและ Deployment แล้วลองใหม่');
      if (error instanceof TypeError) throw new Error('เบราว์เซอร์เชื่อมต่อ Apps Script ไม่ได้ กรุณาตรวจอินเทอร์เน็ต, URL /exec ปัจจุบัน และสิทธิ์ Web app');
      throw error;
    } finally { clearTimeout(timeoutId); }
  }

  function isRetryable(error) {
    const message = String(error && error.message || '');
    if (/HTTP 404|HTTP 410|ยังไม่ได้ตั้งค่า|ไม่ถูกต้อง|HTTP 401|HTTP 403|ไม่มีสิทธิ์|เซสชัน|ชื่อผู้ใช้หรือรหัสผ่าน/i.test(message)) return false;
    return /เชื่อมต่อ|เกินเวลา|HTTP 5\d\d|HTTP 429|ไม่ใช่ JSON/i.test(message);
  }

  async function invokeAppsScriptWithRetry(action,args) {
    let lastError;
    for (let attempt=0; attempt<2; attempt++) {
      try { return await invokeAppsScriptOnce(action,args); }
      catch (error) {
        lastError=error;
        if (!isRetryable(error) || attempt===1) throw error;
        await new Promise(resolve=>setTimeout(resolve,350+Math.floor(Math.random()*250)));
      }
    }
    throw lastError;
  }

  function invokeAppsScript(action,args) {
    const method=String(action||'');
    const normalizedArgs = Array.isArray(args) ? args : [];

    if (READ_ACTIONS.has(method)) {
      const cached = getCachedRead(method, normalizedArgs);
      if (cached !== null) return Promise.resolve(cached);

      const key = readCacheKey(method, normalizedArgs);
      if (inflightReads.has(key)) return inflightReads.get(key);

      const promise = invokeAppsScriptWithRetry(method, normalizedArgs)
        .then(result => {
          putCachedRead(method, normalizedArgs, result);
          return result;
        })
        .finally(() => inflightReads.delete(key));

      inflightReads.set(key, promise);
      return promise;
    }

    const request = RETRY_ACTIONS.has(method)
      ? invokeAppsScriptWithRetry(method, normalizedArgs)
      : invokeAppsScriptOnce(method, normalizedArgs);

    return request.then(result => {
      clearReadCache();
      return result;
    });
  }

  function createRunner(successHandler,failureHandler) {
    return new Proxy({}, {
      get:function(_target,property) {
        if (property === 'then') return undefined;
        if (property === 'withSuccessHandler') return handler => createRunner(handler,failureHandler);
        if (property === 'withFailureHandler') return handler => createRunner(successHandler,handler);
        return function() {
          const args=Array.prototype.slice.call(arguments);
          invokeAppsScript(String(property),args).then(result=>{
            if (typeof successHandler === 'function') successHandler(result);
          }).catch(error=>{
            if (typeof failureHandler === 'function') failureHandler(error);
            else console.error('[Apps Script API]',property,error);
          });
        };
      }
    });
  }

  global.google=global.google||{};
  global.google.script=global.google.script||{};
  global.google.script.run=createRunner(null,null);
  global.UBUApi=Object.freeze({
    call:function(method){const args=Array.prototype.slice.call(arguments,1);return invokeAppsScript(method,args);},
    getUrl:getApiUrl,
    clearCache:clearReadCache
  });
})(window);
