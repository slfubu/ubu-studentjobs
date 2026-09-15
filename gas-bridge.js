(function (global) {
  'use strict';

  // FAST V5: heavy staff reads do not auto-retry; this prevents duplicate Apps Script executions.

  const CLIENT_ID_KEY = 'UBU_API_CLIENT_ID_V1';
  const READ_ACTIONS = new Set([
    'warmStaffLogin','getJobs','getPublicContents','getApplicationStatus','getApplicationProfileImage','getLegacyApplicantByIdCard',
    'getStaffSession','getStaffDashboard','getAdminApplicantPage','getAdminCivilRegistryIds',
    'getStaffAccountManagement','getAdminPublicContents','getAdminQualificationApplicants',
    'getAdminDepartmentForwarding','getAdminDepartmentForwardingApplicants','getAdminDepartmentResults','getDepartmentApplicants',
    'getStaffApplicantDetail','getStaffApplicantPhotoBatch'
  ]);
  const AUTH_ACTIONS = new Set(['staffLogin','getStaffSession']);
  // staffLogin ห้าม retry อัตโนมัติ: การยิง Login ซ้ำทำให้ Apps Script ทำงานซ้อนและสร้าง session ซ้ำ
  const RETRY_ACTIONS = new Set(['getJobs','getPublicContents','getLegacyApplicantByIdCard','getStaffSession','warmStaffLogin']);
  const STATUS_SAFE_RETRY_ACTIONS = new Set([]);
  const inflightReads = new Map();
  const readResponseCache = new Map();

  // FAST V8 Reliability:
  // fresh = ใช้ได้ทันที, stale = เก็บสำรองไว้เผื่อ Apps Script cold start/เครือข่ายสะดุด
  const READ_CACHE_TTL_MS = Object.freeze({
    warmStaffLogin: 120000,
    getJobs: 600000,
    getPublicContents: 300000,
    getApplicationStatus: 30000,
    getApplicationProfileImage: 300000,
    getStaffSession: 8000,
    getStaffDashboard: 180000,
    getAdminApplicantPage: 180000,
    getAdminCivilRegistryIds: 30000,
    getStaffAccountManagement: 60000,
    getAdminPublicContents: 60000,
    getAdminQualificationApplicants: 180000,
    getAdminDepartmentForwarding: 180000,
    getAdminDepartmentForwardingApplicants: 180000,
    getAdminDepartmentResults: 180000,
    getDepartmentApplicants: 180000,
    getStaffApplicantDetail: 15000,
    getStaffApplicantPhotoBatch: 30000
  });

  const STALE_CACHE_TTL_MS = Object.freeze({
    getStaffDashboard: 20 * 60 * 1000,
    getAdminApplicantPage: 20 * 60 * 1000,
    getAdminCivilRegistryIds: 10 * 60 * 1000,
    getStaffAccountManagement: 10 * 60 * 1000,
    getAdminPublicContents: 10 * 60 * 1000,
    getAdminQualificationApplicants: 20 * 60 * 1000,
    getAdminDepartmentForwarding: 20 * 60 * 1000,
    getAdminDepartmentForwardingApplicants: 20 * 60 * 1000,
    getAdminDepartmentResults: 20 * 60 * 1000,
    getDepartmentApplicants: 20 * 60 * 1000,
    getStaffApplicantDetail: 5 * 60 * 1000,
    getStaffApplicantPhotoBatch: 5 * 60 * 1000
  });

  const HEAVY_READ_ACTIONS = new Set([
    'getStaffDashboard','getAdminApplicantPage','getAdminCivilRegistryIds',
    'getAdminQualificationApplicants','getAdminDepartmentForwarding',
    'getAdminDepartmentForwardingApplicants','getAdminDepartmentResults',
    'getDepartmentApplicants','getStaffApplicantPhotoBatch'
  ]);

  function readCacheKey(method, args) {
    try { return method + '|' + JSON.stringify(Array.isArray(args) ? args : []); }
    catch (_) { return method; }
  }

  function getCachedReadEntry(method, args) {
    const key = readCacheKey(method, args);
    const hit = readResponseCache.get(key);
    if (!hit) return { key, state: 'miss', value: null };

    const age = Date.now() - Number(hit.at || 0);
    const freshTtl = Number(READ_CACHE_TTL_MS[method] || 0);
    const staleTtl = Math.max(freshTtl, Number(STALE_CACHE_TTL_MS[method] || freshTtl));

    if (freshTtl > 0 && age <= freshTtl) return { key, state: 'fresh', value: hit.value };
    if (staleTtl > 0 && age <= staleTtl) return { key, state: 'stale', value: hit.value };

    readResponseCache.delete(key);
    return { key, state: 'miss', value: null };
  }

  function putCachedRead(method, args, value) {
    const ttl = Number(READ_CACHE_TTL_MS[method] || 0);
    if (!ttl) return;
    const key = readCacheKey(method, args);
    readResponseCache.set(key, { at: Date.now(), value: value });
    if (readResponseCache.size > 100) {
      const firstKey = readResponseCache.keys().next().value;
      if (firstKey) readResponseCache.delete(firstKey);
    }
  }

  function clearReadCache() {
    readResponseCache.clear();
  }

  function clearReadCacheFor(methods) {
    const targets = new Set(Array.isArray(methods) ? methods.map(String) : []);
    if (!targets.size) return;
    for (const key of Array.from(readResponseCache.keys())) {
      const method = String(key).split('|', 1)[0];
      if (targets.has(method)) readResponseCache.delete(key);
    }
  }

  const WRITE_INVALIDATES = Object.freeze({
    setDepartmentBasket: ['getStaffDashboard','getDepartmentApplicants','getAdminDepartmentResults'],
    finalizeDepartmentSelection: ['getStaffDashboard','getDepartmentApplicants','getAdminDepartmentResults'],
    forwardQualifiedApplicants: ['getStaffDashboard','getAdminApplicantPage','getAdminDepartmentForwarding','getAdminDepartmentForwardingApplicants','getAdminDepartmentResults','getDepartmentApplicants'],
    saveQualificationReviews: ['getStaffDashboard','getAdminApplicantPage','getAdminCivilRegistryIds','getAdminQualificationApplicants','getAdminDepartmentForwarding','getAdminDepartmentForwardingApplicants','getAdminDepartmentResults','getStaffApplicantDetail'],
    processQualificationImport: ['getStaffDashboard','getAdminApplicantPage','getAdminQualificationApplicants','getAdminDepartmentForwarding','getAdminDepartmentForwardingApplicants','getAdminDepartmentResults','getStaffApplicantDetail'],
    createStaffAccount: ['getStaffAccountManagement','getStaffSession'],
    savePublicContent: ['getAdminPublicContents','getPublicContents'],
    deletePublicContent: ['getAdminPublicContents','getPublicContents']
  });

  function invalidateAfterWrite(method) {
    const targets = WRITE_INVALIDATES[String(method || '')];
    if (targets) clearReadCacheFor(targets);
    else clearReadCache();
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

  // FAST V8 Reliability: ลด false-timeout ของ Apps Script cold start
  // เมนูหนักรอได้นานขึ้น แต่จะไม่ retry ซ้ำอัตโนมัติ เพื่อไม่สร้าง execution ซ้อน
  const READ_TIMEOUT_BY_ACTION = Object.freeze({
    warmStaffLogin: 30000,
    getJobs: 45000,
    getPublicContents: 45000,
    getApplicationStatus: 120000,
    getApplicationProfileImage: 90000,
    getLegacyApplicantByIdCard: 45000,
    getStaffSession: 30000,
    getStaffDashboard: 120000,
    getAdminApplicantPage: 240000,
    getAdminCivilRegistryIds: 180000,
    getStaffAccountManagement: 90000,
    getAdminPublicContents: 90000,
    getAdminQualificationApplicants: 240000,
    getAdminDepartmentForwarding: 240000,
    getAdminDepartmentForwardingApplicants: 240000,
    getAdminDepartmentResults: 240000,
    getDepartmentApplicants: 240000,
    getStaffApplicantDetail: 90000,
    getStaffApplicantPhotoBatch: 180000
  });

  function timeoutFor(action) {
    const config = getConfig();
    const name = String(action || '');
    if (AUTH_ACTIONS.has(name)) return Math.max(30000, Number(config.AUTH_TIMEOUT_MS) || 60000);
    if (['uploadApplicationFiles','savePublicContent','processQualificationImport'].includes(name)) {
      return Math.max(180000, Number(config.REQUEST_TIMEOUT_MS) || 90000);
    }
    if (name === 'submitApplication') return Math.max(120000, Number(config.REQUEST_TIMEOUT_MS) || 90000);
    if (READ_ACTIONS.has(name)) {
      const configured = Math.max(45000, Number(config.READ_TIMEOUT_MS) || 90000);
      return Math.max(configured, Number(READ_TIMEOUT_BY_ACTION[name] || 0));
    }
    return Math.max(45000, Number(config.REQUEST_TIMEOUT_MS) || 90000);
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
    const startedAt = (global.performance && typeof global.performance.now === 'function')
      ? global.performance.now()
      : Date.now();
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
      const finishedAt = (global.performance && typeof global.performance.now === 'function')
        ? global.performance.now()
        : Date.now();
      const elapsedMs = Math.round(finishedAt - startedAt);
      const serverMs = Number(payload && payload.meta && payload.meta.serverMs || 0);
      if (getConfig().DEBUG_API) {
        console.debug('[UBU API]', action, { elapsedMs, serverMs, data: payload.data });
      } else if (elapsedMs >= 4000) {
        // แสดงชื่อ action และเวลาเท่านั้น ไม่ log ข้อมูลส่วนบุคคล
        console.warn('[UBU API SLOW]', String(action || ''), `${elapsedMs}ms`, serverMs ? `(server ${serverMs}ms)` : '');
      }
      if (Object.prototype.hasOwnProperty.call(payload,'data')) return payload.data;
      if (Object.prototype.hasOwnProperty.call(payload,'result')) return payload.result;
      return null;
    } catch (error) {
      if (error && error.name === 'AbortError') {
        if (String(action || '') === 'staffLogin') {
          throw new Error('ระบบเจ้าหน้าที่ตอบสนองช้ากว่าปกติ กรุณารอสักครู่แล้วลองเข้าสู่ระบบใหม่อีกครั้ง โดยไม่จำเป็นต้องเปลี่ยนอุปกรณ์หรืออินเทอร์เน็ต');
        }
        throw new Error('การเชื่อมต่อกับฐานข้อมูลไม่ตอบกลับภายในเวลาสูงสุด กรุณากดโหลดข้อมูลอีกครั้ง');
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

  async function invokeStatusWithSafeRetry(action, args) {
    // Human token เป็น one-time แต่ฝั่ง server ทำ idempotency สำหรับ query เดิมแล้ว
    // retry เฉพาะ network/transport error ที่เกิดเร็ว ไม่ retry กรณี timeout 120 วินาที
    try {
      return await invokeAppsScriptOnce(action, args);
    } catch (error) {
      const message = String(error && error.message || '');
      const isTimeout = /ไม่ตอบกลับภายในเวลาสูงสุด|ใช้เวลานาน/i.test(message);
      const isTransport = /ไม่สามารถเชื่อมต่อ|ข้อผิดพลาดในการเชื่อมต่อ|ข้อผิดพลาดในการรับส่งข้อมูล/i.test(message);
      if (isTimeout || !isTransport) throw error;
      await new Promise(resolve => setTimeout(resolve, 500 + Math.floor(Math.random() * 350)));
      return invokeAppsScriptOnce(action, args);
    }
  }

  function startReadRefresh(method, normalizedArgs, key) {
    if (inflightReads.has(key)) return inflightReads.get(key);

    // สำคัญ: V7 เรียก invokeAppsScriptWithRetry() กับ read ทุกตัว ทำให้เมนูหนัก timeout แล้วถูกยิงซ้ำ
    // V8 retry เฉพาะ action เบาที่อยู่ใน RETRY_ACTIONS เท่านั้น
    const request = STATUS_SAFE_RETRY_ACTIONS.has(method)
      ? invokeStatusWithSafeRetry(method, normalizedArgs)
      : RETRY_ACTIONS.has(method)
        ? invokeAppsScriptWithRetry(method, normalizedArgs)
        : invokeAppsScriptOnce(method, normalizedArgs);

    const promise = request
      .then(result => {
        putCachedRead(method, normalizedArgs, result);
        return result;
      })
      .finally(() => inflightReads.delete(key));

    inflightReads.set(key, promise);
    return promise;
  }

  function invokeAppsScript(action,args) {
    const method=String(action||'');
    const normalizedArgs = Array.isArray(args) ? args : [];

    if (READ_ACTIONS.has(method)) {
      const cached = getCachedReadEntry(method, normalizedArgs);
      if (cached.state === 'fresh') return Promise.resolve(cached.value);

      if (cached.state === 'stale' && STALE_CACHE_TTL_MS[method]) {
        // stale-while-revalidate: แสดงข้อมูลเดิมทันที แล้วรีเฟรชเงียบ ๆ ด้านหลัง
        startReadRefresh(method, normalizedArgs, cached.key).catch(error => {
          console.warn('[UBU API refresh failed]', method, error && error.message ? error.message : error);
        });
        return Promise.resolve(cached.value);
      }

      const key = cached.key;
      if (inflightReads.has(key)) return inflightReads.get(key);

      return startReadRefresh(method, normalizedArgs, key).catch(error => {
        // หากระหว่างรอมี cache เกิดขึ้นจาก request อื่น ให้ใช้ cache แทน error
        const fallback = getCachedReadEntry(method, normalizedArgs);
        if (fallback.state === 'fresh' || fallback.state === 'stale') return fallback.value;
        throw error;
      });
    }

    const request = RETRY_ACTIONS.has(method)
      ? invokeAppsScriptWithRetry(method, normalizedArgs)
      : invokeAppsScriptOnce(method, normalizedArgs);

    return request.then(result => {
      invalidateAfterWrite(method);
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
    clearCache:clearReadCache,
    clearMethodCache:function(methods){clearReadCacheFor(Array.isArray(methods)?methods:[methods]);}
  });
})(window);
