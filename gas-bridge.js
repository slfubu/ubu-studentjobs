(function (global) {
  'use strict';

  // FAST V11.3 ROBUST BRIDGE + NEWS V5: hidden iframe transport + preload bundle + delta cache.

  const CLIENT_ID_KEY = 'UBU_API_CLIENT_ID_V1';
  const READ_ACTIONS = new Set([
    'warmStaffLogin','getJobs','getPublicContents','getApplicationStatus','getApplicationProfileImage','getLegacyApplicantByIdCard',
    'getStaffSession','getStaffDashboard','getStaffPreloadBundle','getAdminApplicantPage','getAdminCivilRegistryIds',
    'getStaffAccountManagement','getAdminPublicContents','getAdminQualificationApplicants','getAdminQualificationPage',
    'getAdminDepartmentForwarding','getAdminDepartmentForwardingApplicants','getAdminDepartmentResults','getAdminDepartmentResultApplicants','getAdminDepartmentResultsExport','getDepartmentApplicants',
    'getStaffApplicantDetail','getStaffApplicantPhotoBatch'
  ]);
  const AUTH_ACTIONS = new Set(['staffLogin','getStaffSession']);
  // staffLogin ห้าม retry อัตโนมัติ: การยิง Login ซ้ำทำให้ Apps Script ทำงานซ้อนและสร้าง session ซ้ำ
  const RETRY_ACTIONS = new Set(['getJobs','getPublicContents','getLegacyApplicantByIdCard','getStaffSession','warmStaffLogin']);
  const STATUS_SAFE_RETRY_ACTIONS = new Set([]);
  const inflightReads = new Map();
  const readResponseCache = new Map();

  // FAST V11: เก็บ response หน้ารายการ/summary ใน sessionStorage
  // จึงยังเปิดได้ทันทีหลัง Refresh ใน tab เดิม โดยไม่เก็บรายละเอียด/รูป/ไฟล์ export ขนาดใหญ่
  const SESSION_READ_CACHE_KEY = 'UBU_STAFF_READ_CACHE_FAST_V11';
  const SESSION_READ_CACHE_MAX_BYTES = 1200000;
  const SESSION_PERSIST_ACTIONS = new Set([
    'getStaffDashboard','getAdminApplicantPage','getStaffAccountManagement','getAdminPublicContents',
    'getAdminQualificationPage','getAdminDepartmentForwarding','getAdminDepartmentForwardingApplicants',
    'getAdminDepartmentResults','getAdminDepartmentResultApplicants','getDepartmentApplicants'
  ]);
  let sessionCacheFlushTimer = 0;

  function hydrateSessionReadCache() {
    try {
      const raw = sessionStorage.getItem(SESSION_READ_CACHE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      const entries = parsed && Array.isArray(parsed.entries) ? parsed.entries : [];
      const now = Date.now();
      entries.forEach(entry => {
        if (!entry || typeof entry.key !== 'string' || !entry.at) return;
        const method = String(entry.key).split('|', 1)[0];
        if (!SESSION_PERSIST_ACTIONS.has(method)) return;
        const staleTtl = Math.max(
          Number(READ_CACHE_TTL_MS && READ_CACHE_TTL_MS[method] || 0),
          Number(STALE_CACHE_TTL_MS && STALE_CACHE_TTL_MS[method] || 0)
        );
        if (!staleTtl || now - Number(entry.at) > staleTtl) return;
        readResponseCache.set(entry.key, { at: Number(entry.at), value: entry.value });
      });
    } catch (_) {
      try { sessionStorage.removeItem(SESSION_READ_CACHE_KEY); } catch (_) {}
    }
  }

  function flushSessionReadCache() {
    clearTimeout(sessionCacheFlushTimer);
    sessionCacheFlushTimer = 0;
    try {
      const entries = [];
      for (const [key, hit] of readResponseCache.entries()) {
        const method = String(key).split('|', 1)[0];
        if (!SESSION_PERSIST_ACTIONS.has(method) || !hit) continue;
        entries.push({ key, at: Number(hit.at || Date.now()), value: hit.value });
      }
      entries.sort((a,b) => b.at - a.at);
      let kept = entries.slice(0, 24);
      let raw = JSON.stringify({ v: 11, entries: kept });
      while (raw.length > SESSION_READ_CACHE_MAX_BYTES && kept.length > 1) {
        kept.pop();
        raw = JSON.stringify({ v: 11, entries: kept });
      }
      if (kept.length) sessionStorage.setItem(SESSION_READ_CACHE_KEY, raw);
      else sessionStorage.removeItem(SESSION_READ_CACHE_KEY);
    } catch (_) {
      // sessionStorage เต็ม/ปิดใช้งาน: ใช้ memory cache ต่อได้ตามปกติ
    }
  }

  function scheduleSessionReadCacheFlush() {
    clearTimeout(sessionCacheFlushTimer);
    sessionCacheFlushTimer = setTimeout(flushSessionReadCache, 80);
  }

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
    getStaffPreloadBundle: 0,
    getAdminApplicantPage: 180000,
    getAdminCivilRegistryIds: 30000,
    getStaffAccountManagement: 60000,
    getAdminPublicContents: 60000,
    getAdminQualificationApplicants: 60000,
    getAdminQualificationPage: 60000,
    getAdminDepartmentForwarding: 180000,
    getAdminDepartmentForwardingApplicants: 180000,
    getAdminDepartmentResults: 180000,
    getAdminDepartmentResultApplicants: 60000,
    getAdminDepartmentResultsExport: 30000,
    getDepartmentApplicants: 180000,
    getStaffApplicantDetail: 15000,
    getStaffApplicantPhotoBatch: 30000
  });

  const STALE_CACHE_TTL_MS = Object.freeze({
    getStaffDashboard: 20 * 60 * 1000,
    getStaffPreloadBundle: 0,
    getAdminApplicantPage: 20 * 60 * 1000,
    getAdminCivilRegistryIds: 10 * 60 * 1000,
    getStaffAccountManagement: 10 * 60 * 1000,
    getAdminPublicContents: 10 * 60 * 1000,
    getAdminQualificationApplicants: 10 * 60 * 1000,
    getAdminQualificationPage: 10 * 60 * 1000,
    getAdminDepartmentForwarding: 20 * 60 * 1000,
    getAdminDepartmentForwardingApplicants: 20 * 60 * 1000,
    getAdminDepartmentResults: 20 * 60 * 1000,
    getAdminDepartmentResultApplicants: 10 * 60 * 1000,
    getAdminDepartmentResultsExport: 5 * 60 * 1000,
    getDepartmentApplicants: 20 * 60 * 1000,
    getStaffApplicantDetail: 5 * 60 * 1000,
    getStaffApplicantPhotoBatch: 5 * 60 * 1000
  });

  hydrateSessionReadCache();

  const HEAVY_READ_ACTIONS = new Set([
    'getStaffDashboard','getStaffPreloadBundle','getAdminApplicantPage','getAdminCivilRegistryIds',
    'getAdminQualificationApplicants','getAdminQualificationPage','getAdminDepartmentForwarding',
    'getAdminDepartmentForwardingApplicants','getAdminDepartmentResults','getAdminDepartmentResultApplicants','getAdminDepartmentResultsExport',
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
    if (SESSION_PERSIST_ACTIONS.has(method)) scheduleSessionReadCacheFlush();
    if (readResponseCache.size > 100) {
      const firstKey = readResponseCache.keys().next().value;
      if (firstKey) readResponseCache.delete(firstKey);
    }
  }

  function clearReadCache() {
    readResponseCache.clear();
    try { sessionStorage.removeItem(SESSION_READ_CACHE_KEY); } catch (_) {}
  }

  function clearReadCacheFor(methods) {
    const targets = new Set(Array.isArray(methods) ? methods.map(String) : []);
    if (!targets.size) return;
    for (const key of Array.from(readResponseCache.keys())) {
      const method = String(key).split('|', 1)[0];
      if (targets.has(method)) readResponseCache.delete(key);
    }
    scheduleSessionReadCacheFlush();
  }

  const WRITE_INVALIDATES = Object.freeze({
    setDepartmentBasket: ['getStaffDashboard','getDepartmentApplicants','getAdminDepartmentResults'],
    finalizeDepartmentSelection: ['getStaffDashboard','getDepartmentApplicants','getAdminDepartmentResults','getAdminDepartmentResultApplicants','getAdminDepartmentResultsExport'],
    forwardQualifiedApplicants: ['getStaffDashboard','getAdminApplicantPage','getAdminDepartmentForwarding','getAdminDepartmentForwardingApplicants','getAdminDepartmentResults','getAdminDepartmentResultApplicants','getAdminDepartmentResultsExport','getDepartmentApplicants'],
    saveQualificationReviews: ['getStaffDashboard','getAdminApplicantPage','getAdminCivilRegistryIds','getAdminQualificationApplicants','getAdminQualificationPage','getAdminDepartmentForwarding','getAdminDepartmentForwardingApplicants','getAdminDepartmentResults','getAdminDepartmentResultApplicants','getAdminDepartmentResultsExport','getStaffApplicantDetail'],
    processQualificationImport: ['getStaffDashboard','getAdminApplicantPage','getAdminQualificationApplicants','getAdminQualificationPage','getAdminDepartmentForwarding','getAdminDepartmentForwardingApplicants','getAdminDepartmentResults','getAdminDepartmentResultApplicants','getAdminDepartmentResultsExport','getStaffApplicantDetail'],
    createStaffAccount: ['getStaffAccountManagement','getStaffSession'],
    savePublicContent: ['getAdminPublicContents','getPublicContents'],
    deletePublicContent: ['getAdminPublicContents','getPublicContents'],
    // ตัวนับการอ่านไม่ควรล้าง cache ข่าว/เมนูอื่นทุกครั้งที่มีผู้ใช้เปิดข่าว
    recordPublicContentView: []
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

  // =====================================================
  // FAST V11.3 ROBUST BRIDGE: native form POST + one-shot iframe
  // No fetch CORS and no persistent bridge READY handshake.
  // =====================================================
  const STAFF_FORM_BRIDGE_ACTIONS = new Set([
    'warmStaffLogin','staffLogin','staffSsoLogin','staffLogout','getStaffSession',
    'getStaffDashboard','getStaffPreloadBundle','getAdminApplicantPage','getAdminCivilRegistryIds',
    'getStaffAccountManagement','createStaffAccount','getAdminPublicContents','savePublicContent','deletePublicContent',
    'getAdminQualificationApplicants','getAdminQualificationPage','saveQualificationReviews','processQualificationImport',
    'getAdminDepartmentForwarding','getAdminDepartmentForwardingApplicants','forwardQualifiedApplicants',
    'getAdminDepartmentResults','getAdminDepartmentResultApplicants','getAdminDepartmentResultsExport',
    'getDepartmentApplicants','setDepartmentBasket','finalizeDepartmentSelection',
    'getStaffApplicantDetail','getStaffApplicantPhotoBatch'
  ]);

  function formBridgeUrl() {
    const base = getApiUrl().split('?')[0];
    return base + '?transport=form_bridge&v=20260917-fast-v11-3';
  }

  function invokeAppsScriptViaFormBridge(action, args) {
    return new Promise((resolve, reject) => {
      const id = randomId();
      const frame = document.createElement('iframe');
      const frameName = 'ubuGasFormBridge_' + id;
      frame.name = frameName;
      frame.id = frameName;
      frame.setAttribute('aria-hidden', 'true');
      frame.tabIndex = -1;
      frame.style.cssText = 'position:fixed;width:1px;height:1px;left:-10000px;top:-10000px;border:0;opacity:0;pointer-events:none;';

      const request = {
        action: String(action || ''),
        args: Array.isArray(args) ? args : [],
        authToken: getAuthToken(),
        meta: buildMeta()
      };

      let settled = false;
      let timer = null;

      function cleanup() {
        if (timer) clearTimeout(timer);
        global.removeEventListener('message', onMessage);
        try { frame.remove(); } catch (_) {}
      }

      function finishError(message) {
        if (settled) return;
        settled = true;
        cleanup();
        reject(new Error(message));
      }

      function onMessage(event) {
        if (settled) return;
        const origin = String(event.origin || '');
        const trustedOrigin = origin === 'null' || origin === 'https://script.google.com' || /^https:\/\/[^/]*googleusercontent\.com$/i.test(origin);
        if (!trustedOrigin) return;
        const msg = event.data || {};
        if (msg.type !== 'UBU_GAS_FORM_BRIDGE_RESPONSE' || String(msg.id || '') !== id) return;
        settled = true;
        cleanup();
        const payload = msg.payload;
        if (!payload || payload.ok !== true) {
          const message = payload && payload.error && payload.error.message
            ? String(payload.error.message)
            : 'เกิดข้อผิดพลาดในการประมวลผลคำขอ';
          reject(new Error(message));
          return;
        }
        if (Object.prototype.hasOwnProperty.call(payload, 'data')) resolve(payload.data);
        else if (Object.prototype.hasOwnProperty.call(payload, 'result')) resolve(payload.result);
        else resolve(null);
      }

      global.addEventListener('message', onMessage);
      (document.body || document.documentElement).appendChild(frame);

      const form = document.createElement('form');
      form.method = 'POST';
      form.action = formBridgeUrl();
      form.target = frameName;
      form.enctype = 'application/x-www-form-urlencoded';
      form.acceptCharset = 'UTF-8';
      form.style.display = 'none';

      const transportInput = document.createElement('input');
      transportInput.type = 'hidden';
      transportInput.name = 'transport';
      transportInput.value = 'form_bridge';
      form.appendChild(transportInput);

      const idInput = document.createElement('input');
      idInput.type = 'hidden';
      idInput.name = 'id';
      idInput.value = id;
      form.appendChild(idInput);

      const payloadInput = document.createElement('input');
      payloadInput.type = 'hidden';
      payloadInput.name = 'payload';
      payloadInput.value = JSON.stringify(request);
      form.appendChild(payloadInput);

      (document.body || document.documentElement).appendChild(form);
      timer = setTimeout(() => {
        try { form.remove(); } catch (_) {}
        finishError('Apps Script ไม่ตอบกลับผ่าน Form Bridge หาก Deploy V11.3 แล้ว ให้ตรวจ Web App access ว่าเปิดแบบ Execute as me และ Who has access = Anyone');
      }, timeoutFor(action));

      try {
        form.submit();
        // Once submitted, the browser has copied the form body into the navigation request.
        setTimeout(() => { try { form.remove(); } catch (_) {} }, 0);
      } catch (_) {
        try { form.remove(); } catch (_) {}
        finishError('ไม่สามารถส่งคำขอไปยัง Apps Script ได้');
      }
    });
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
    getStaffPreloadBundle: 240000,
    getAdminApplicantPage: 240000,
    getAdminCivilRegistryIds: 180000,
    getStaffAccountManagement: 90000,
    getAdminPublicContents: 90000,
    getAdminQualificationApplicants: 240000,
    getAdminQualificationPage: 120000,
    getAdminDepartmentForwarding: 240000,
    getAdminDepartmentForwardingApplicants: 240000,
    getAdminDepartmentResults: 120000,
    getAdminDepartmentResultApplicants: 120000,
    getAdminDepartmentResultsExport: 240000,
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
    const startedAt = (global.performance && typeof global.performance.now === 'function')
      ? global.performance.now()
      : Date.now();

    // Staff Portal on GitHub Pages uses native form POST into a one-shot iframe.
    // This bypasses fetch CORS without a persistent READY bridge.
    const isGithubProduction = String(global.location && global.location.origin || '') === 'https://slfubu.github.io';

    try {
      let result;
      if (isGithubProduction && STAFF_FORM_BRIDGE_ACTIONS.has(String(action || ''))) {
        result = await invokeAppsScriptViaFormBridge(action, args);
      } else {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), timeoutFor(action));
        try {
          const response = await fetch(getApiUrl(), {
            method: 'POST', redirect: 'follow', credentials: 'omit', cache: 'no-store',
            headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
            body: JSON.stringify({ action:String(action||''), args:Array.isArray(args)?args:[], authToken:getAuthToken(), meta:buildMeta() }),
            signal: controller.signal
          });
          const raw = await response.text();
          if (!response.ok) throw httpError(response.status);
          let payload;
          try { payload = JSON.parse(raw); }
          catch (_) { throw new Error('เกิดข้อผิดพลาดในการรับส่งข้อมูล กรุณาลองใหม่อีกครั้ง'); }
          if (!payload || payload.ok !== true) {
            throw new Error(payload && payload.error && payload.error.message ? String(payload.error.message) : 'เกิดข้อผิดพลาดในการประมวลผลคำขอ');
          }
          result = Object.prototype.hasOwnProperty.call(payload,'data') ? payload.data : payload.result;
        } finally {
          clearTimeout(timeoutId);
        }
      }

      const finishedAt = (global.performance && typeof global.performance.now === 'function')
        ? global.performance.now()
        : Date.now();
      const elapsedMs = Math.round(finishedAt - startedAt);
      if (getConfig().DEBUG_API) console.debug('[UBU API BRIDGE]', action, { elapsedMs, data: result });
      else if (elapsedMs >= 4000) console.warn('[UBU API SLOW]', String(action || ''), `${elapsedMs}ms`, '(form bridge)');
      return result;
    } catch (error) {
      if (error && error.name === 'AbortError') {
        throw new Error('การเชื่อมต่อกับฐานข้อมูลไม่ตอบกลับภายในเวลาสูงสุด กรุณากดโหลดข้อมูลอีกครั้ง');
      }
      if (error instanceof TypeError) {
        throw new Error('ไม่สามารถเชื่อมต่อกับระบบได้ กรุณาตรวจสอบอินเทอร์เน็ตของท่านและลองใหม่อีกครั้ง');
      }
      throw error;
    }
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
    clearMethodCache:function(methods){clearReadCacheFor(Array.isArray(methods)?methods:[methods]);},
    // FAST V11: preload bundle สามารถเติม cache ของ endpoint จริงได้โดยไม่ยิง request ซ้ำ
    seedCache:function(method,args,value){
      const name=String(method||'');
      if(!READ_ACTIONS.has(name)) return false;
      putCachedRead(name,Array.isArray(args)?args:[],value);
      return true;
    }
  });
})(window);
