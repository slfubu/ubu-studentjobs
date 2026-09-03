(function (global) {
  'use strict';

  const CLIENT_ID_KEY = 'UBU_API_CLIENT_ID_V1';
  const READ_ACTIONS = new Set([
    'warmStaffLogin','getJobs','getPublicContents','getApplicationStatus','getApplicationProfileImage','getLegacyApplicantByIdCard',
    'getStaffSession','getStaffDashboard','getAdminApplicantPage','getAdminCivilRegistryIds',
    'getStaffAccountManagement','getAdminPublicContents','getAdminQualificationApplicants',
    'getAdminDepartmentForwarding','getAdminDepartmentResults','getDepartmentApplicants',
    'getStaffApplicantDetail','getStaffApplicantPhotoBatch'
  ]);
  const AUTH_ACTIONS = new Set(['staffLogin','getStaffSession']);
  // staffLogin ห้าม retry อัตโนมัติ: การยิง Login ซ้ำทำให้ Apps Script ทำงานซ้อนและสร้าง session ซ้ำ
  const RETRY_ACTIONS = new Set([...READ_ACTIONS]);
  const inflightReads = new Map();
  const readResponseCache = new Map();
  const READ_CACHE_TTL_MS = Object.freeze({
    warmStaffLogin: 120000,
    getJobs: 300000,
    getPublicContents: 120000,
    getApplicationStatus: 10000,
    getApplicationProfileImage: 120000,
    getStaffSession: 5000,
    getStaffDashboard: 45000,
    getAdminApplicantPage: 30000,
    getAdminCivilRegistryIds: 5000,
    getStaffAccountManagement: 30000,
    getAdminPublicContents: 30000,
    getAdminQualificationApplicants: 30000,
    getAdminDepartmentForwarding: 30000,
    getAdminDepartmentResults: 30000,
    getDepartmentApplicants: 30000,
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
      throw new Error('ระบบยังไม่พร้อมใช้งาน กรุณาติดต่อเจ้าหน้าที่ดูแลระบบ');
    }
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec(?:\?.*)?$/i.test(url)) {
      throw new Error('ไม่สามารถเชื่อมต่อระบบฐานข้อมูลได้ในขณะนี้ กรุณาติดต่อเจ้าหน้าที่ดูแลระบบ');
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
    if (AUTH_ACTIONS.has(name)) return Math.max(12000, Number(config.AUTH_TIMEOUT_MS) || 25000);
    if (['uploadApplicationFiles','savePublicContent','processQualificationImport'].includes(name)) {
      return Math.max(120000, Number(config.REQUEST_TIMEOUT_MS) || 45000);
    }
    if (name === 'submitApplication') return Math.max(90000, Number(config.REQUEST_TIMEOUT_MS) || 45000);
    if (READ_ACTIONS.has(name)) return Math.max(5000, Number(config.READ_TIMEOUT_MS) || 8000);
    return Math.max(10000, Number(config.REQUEST_TIMEOUT_MS) || 45000);
  }

  function httpError(status) {
    if (status === 404 || status === 410) {
      return new Error('ไม่พบข้อมูลที่ต้องการ กรุณารีเฟรชหน้าเว็บและลองใหม่อีกครั้ง');
    }
    if (status === 401 || status === 403) {
      return new Error('คุณไม่มีสิทธิ์เข้าถึงข้อมูลส่วนนี้ กรุณาเข้าสู่ระบบหรือติดต่อเจ้าหน้าที่');
    }
    if (status === 429) return new Error('ระบบกำลังทำงานหนัก กรุณารอสักครู่แล้วลองใหม่อีกครั้ง');
    return new Error('เกิดข้อผิดพลาดในการเชื่อมต่อ กรุณาลองใหม่อีกครั้ง');
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
          throw new Error('ระบบยังไม่เปิดให้บุคคลทั่วไปเข้าใช้งาน กรุณาติดต่อเจ้าหน้าที่ดูแลระบบ');
        }
        throw new Error('เกิดข้อผิดพลาดในการรับส่งข้อมูล กรุณาลองใหม่อีกครั้ง หรือติดต่อเจ้าหน้าที่');
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
      if (error && error.name === 'AbortError') {
        if (String(action || '') === 'staffLogin') {
          throw new Error('ระบบเจ้าหน้าที่ตอบสนองช้ากว่าปกติ กรุณารอสักครู่แล้วลองเข้าสู่ระบบใหม่อีกครั้ง โดยไม่จำเป็นต้องเปลี่ยนอุปกรณ์หรืออินเทอร์เน็ต');
        }
        throw new Error('ระบบตอบสนองช้ากว่าปกติ กรุณารอสักครู่แล้วลองใหม่อีกครั้ง');
      }
      if (error instanceof TypeError) throw new Error('ไม่สามารถเชื่อมต่อกับระบบได้ กรุณาตรวจสอบอินเทอร์เน็ตของท่านและลองใหม่อีกครั้ง');
      throw error;
    } finally { clearTimeout(timeoutId); }
  }

  function isRetryable(error) {
    const message = String(error && error.message || '');
    // หากเป็นข้อผิดพลาดเรื่องสิทธิ์การเข้าถึง หรือ ระบบปิดให้บริการ จะไม่มีการโหลดซ้ำ (Retry)
    if (/ไม่พบข้อมูล|ไม่มีสิทธิ์|ระบบยังไม่พร้อม|ระบบยังไม่เปิดให้|เซสชัน|ชื่อผู้ใช้หรือรหัสผ่าน|ใช้เวลานาน/i.test(message)) return false;
    // หากเป็นปัญหาเครือข่าย หรือเซิร์ฟเวอร์ทำงานหนัก จะทำการลองใหม่ (Retry)
    return /เชื่อมต่อ|ใช้เวลานาน|ทำงานหนัก|ข้อผิดพลาด/i.test(message);
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

      const readRequest = method === 'warmStaffLogin'
        ? invokeAppsScriptOnce(method, normalizedArgs)
        : invokeAppsScriptWithRetry(method, normalizedArgs);

      const promise = readRequest
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
