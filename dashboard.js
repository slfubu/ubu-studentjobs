      const BUILD_VERSION = '2026.09.28-V12.3-PRINT-TITLES-COMPACT-HEADER';
      console.info('[UBU Student Jobs]', BUILD_VERSION);

      const state = {
      token: '',
      role: '',
      department: '',
      selectionUnit: '',
      displayName: '',
      username: '',
      currentView: '',
      qualificationRows: [],
      qualificationPage: 1,
      qualificationPageSize: 50,
      qualificationPageCount: 1,
      qualificationSearch: '',
      qualificationStatus: 'all',
      qualificationRequestSeq: 0,
      qualificationDrafts: new Map(),
      adminApplicantRows: [],
      civilRegistryRows: [],
      departmentRows: [],
      publicContentRows: [],
      staffAccountRows: [],
      staffAccountUnits: [],
       applicantReviewPage: 1, applicantReviewPageSize: 20, applicantReviewRequestSeq: 0,
       civilRegistryPage: 1, civilRegistryPageSize: 20, civilRegistryRequestSeq: 0,
       departmentPage: 1, departmentPageSize: 20, photoLoadGeneration: 0
    };

    const listModal = new bootstrap.Modal(document.getElementById('listModal'));
    const detailModal = new bootstrap.Modal(document.getElementById('detailModal'));

    // FAST V7: เก็บรูปเฉพาะที่เคยแสดงใน session นี้ ลดการอ่าน DriveApp ซ้ำเมื่อกลับหน้าเดิม
    const applicantPhotoCache = new Map();
    const APPLICANT_PHOTO_CACHE_MAX = 30;

    function rememberApplicantPhoto(id, data) {
      if (!id || !data) return;
      applicantPhotoCache.delete(id);
      applicantPhotoCache.set(id, data);
      while (applicantPhotoCache.size > APPLICANT_PHOTO_CACHE_MAX) {
        const oldest = applicantPhotoCache.keys().next().value;
        applicantPhotoCache.delete(oldest);
      }
    }

    let xlsxLoadPromise = null;
    function ensureXlsxLoaded() {
      if (window.XLSX) return Promise.resolve(window.XLSX);
      if (xlsxLoadPromise) return xlsxLoadPromise;

      xlsxLoadPromise = new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
        script.async = true;
        script.onload = () => window.XLSX
          ? resolve(window.XLSX)
          : reject(new Error('โหลดไลบรารี Excel ไม่สำเร็จ'));
        script.onerror = () => reject(new Error('ไม่สามารถโหลดไลบรารี Excel ได้ กรุณาตรวจสอบอินเทอร์เน็ตแล้วลองใหม่'));
        document.head.appendChild(script);
      }).catch(error => {
        xlsxLoadPromise = null;
        throw error;
      });

      return xlsxLoadPromise;
    }


    // ExcelJS ใช้สำหรับไฟล์รายงานที่ต้องกำหนด Print Titles / Header / Footer / Page x/y
    let excelJsLoadPromise = null;
    function ensureExcelJsLoaded() {
      if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
      if (excelJsLoadPromise) return excelJsLoadPromise;

      excelJsLoadPromise = new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
        script.async = true;
        script.onload = () => window.ExcelJS
          ? resolve(window.ExcelJS)
          : reject(new Error('โหลด ExcelJS ไม่สำเร็จ'));
        script.onerror = () => reject(
          new Error('ไม่สามารถโหลด ExcelJS ได้ กรุณาตรวจสอบอินเทอร์เน็ตแล้วลองใหม่')
        );
        document.head.appendChild(script);
      }).catch(error => {
        excelJsLoadPromise = null;
        throw error;
      });

      return excelJsLoadPromise;
    }

    const STAFF_PRELOAD_REFRESH_WRITES = new Set([
      'saveQualificationReviews','processQualificationImport','forwardQualifiedApplicants',
      'setDepartmentBasket','finalizeDepartmentSelection'
    ]);

    function serverCall(method, ...args) {
      return new Promise((resolve, reject) => {
        try {
          const runner = google.script.run
            .withSuccessHandler(result => {
              resolve(result);
              // หลังข้อมูลหลักเปลี่ยน ให้เตรียม cache ของเมนูอื่นใหม่แบบ debounce
              if (STAFF_PRELOAD_REFRESH_WRITES.has(String(method || ''))) {
                scheduleMenuPreloadRefresh();
              }
            })
            .withFailureHandler(error => {
              reject(new Error(
                error && error.message ? error.message : 'เกิดข้อผิดพลาดในการเชื่อมต่อระบบ'
              ));
            });

          const fn = runner && runner[method];
          if (typeof fn !== 'function') {
            reject(new Error(`ฟังก์ชัน ${method} ยังไม่มีใน Web App เวอร์ชันที่ Deploy อยู่ กรุณาอัปเดตเป็นเวอร์ชันใหม่`));
            return;
          }
          fn.apply(runner, args);
        } catch (error) {
          reject(new Error(
            error && error.message
              ? error.message
              : `ไม่สามารถเรียกฟังก์ชัน ${method} ได้ กรุณาตรวจสอบเวอร์ชันที่ Deploy`
          ));
        }
      });
    }

    // ======================================================
    // FAST V11: PRELOAD ALL STAFF MENUS
    // ======================================================
    let staffMenuPreloadPromise = null;
    let staffMenuPreloadRefreshTimer = 0;
    let staffMenuPreloadQueued = false;

    function setMenuPreloadStatus(mode, text) {
      const el = document.getElementById('sidebarPreloadStatus');
      if (!el) return;
      if (mode === 'ready') {
        el.innerHTML = `<i class="fa-solid fa-circle-check me-1 text-success"></i>${escapeHtml(text || 'ข้อมูลทุกเมนูพร้อมใช้งาน')}`;
      } else if (mode === 'warning') {
        el.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1 text-warning"></i>${escapeHtml(text || 'ข้อมูลบางเมนูจะโหลดเมื่อเปิดใช้งาน')}`;
      } else {
        el.innerHTML = `<i class="fa-solid fa-spinner fa-spin me-1"></i>${escapeHtml(text || 'กำลังเตรียมข้อมูลทุกเมนู')}`;
      }
    }

    function seedStaffPreloadBundle(bundle) {
      if (!bundle || !bundle.items || !window.UBUApi || typeof window.UBUApi.seedCache !== 'function') return 0;
      const items = bundle.items;
      let seeded = 0;
      const seed = (key, method, args) => {
        if (!Object.prototype.hasOwnProperty.call(items, key)) return;
        if (window.UBUApi.seedCache(method, args, items[key])) seeded++;
      };

      if (state.role === 'admin') {
        seed('adminDashboard', 'getStaffDashboard', [state.token]);
        seed('staffAccounts', 'getStaffAccountManagement', [state.token]);
        seed('contentManagement', 'getAdminPublicContents', [state.token]);
        seed('applicantReview', 'getAdminApplicantPage', [state.token, {
          mode: 'review', page: 1, pageSize: 20, includeFilters: true,
          search: '', department: '', faculty: '', status: ''
        }]);
        seed('civilRegistry', 'getAdminApplicantPage', [state.token, {
          mode: 'civil', page: 1, pageSize: 20, includeFilters: true,
          search: '', department: '', faculty: '', registryStatus: ''
        }]);
        seed('qualification', 'getAdminQualificationPage', [state.token, {
          page: 1, pageSize: 50, search: '', status: 'all'
        }]);
        seed('forwarding', 'getAdminDepartmentForwarding', [state.token]);
        seed('adminResults', 'getAdminDepartmentResults', [state.token]);
      } else if (state.role === 'department') {
        seed('departmentDashboard', 'getStaffDashboard', [state.token]);
        seed('departmentApplicants', 'getDepartmentApplicants', [state.token]);
      }
      return seeded;
    }

    async function preloadAllStaffMenus(excludeView = '') {
      if (!state.token || !state.role || !window.UBUApi || typeof window.UBUApi.call !== 'function') return null;
      if (staffMenuPreloadPromise) {
        staffMenuPreloadQueued = true;
        return staffMenuPreloadPromise;
      }

      setMenuPreloadStatus('loading', 'กำลังเตรียมข้อมูลทุกเมนู');
      const tokenAtStart = state.token;
      staffMenuPreloadPromise = window.UBUApi.call('getStaffPreloadBundle', state.token, {
        excludeView: String(excludeView || '')
      }).then(bundle => {
        if (!state.token || state.token !== tokenAtStart) return bundle;
        const count = seedStaffPreloadBundle(bundle);
        setMenuPreloadStatus('ready', count ? 'ข้อมูลทุกเมนูพร้อมใช้งาน' : 'ข้อมูลเมนูพร้อมใช้งาน');
        try {
          sessionStorage.setItem('ubuStaffPreloadReadyAt', String(Date.now()));
        } catch (_) {}
        return bundle;
      }).catch(error => {
        console.warn('[Staff preload]', error && error.message ? error.message : error);
        setMenuPreloadStatus('warning', 'ข้อมูลบางเมนูจะโหลดเมื่อเปิดใช้งาน');
        return null;
      }).finally(() => {
        staffMenuPreloadPromise = null;
        if (staffMenuPreloadQueued && state.token) {
          staffMenuPreloadQueued = false;
          clearTimeout(staffMenuPreloadRefreshTimer);
          staffMenuPreloadRefreshTimer = setTimeout(() => preloadAllStaffMenus(state.currentView), 400);
        }
      });
      return staffMenuPreloadPromise;
    }

    function scheduleMenuPreloadRefresh(delay = 1800) {
      clearTimeout(staffMenuPreloadRefreshTimer);
      staffMenuPreloadRefreshTimer = setTimeout(() => {
        if (!state.token || !state.role) return;
        preloadAllStaffMenus(state.currentView);
      }, Math.max(300, Number(delay) || 1800));
    }

    let staffLoginWarmupPromise = null;

    function startStaffLoginWarmup() {
      if (staffLoginWarmupPromise) return staffLoginWarmupPromise;
      if (!window.UBUApi || typeof window.UBUApi.call !== 'function') {
        return Promise.resolve(null);
      }

      staffLoginWarmupPromise = window.UBUApi.call('warmStaffLogin')
        .catch(error => {
          // Warm-up เป็น best effort: ไม่แสดง error ให้ผู้ใช้ เพราะ Login จริงยังทำงานได้ตามปกติ
          console.debug('[Staff warm-up]', error && error.message ? error.message : error);
          return null;
        });

      return staffLoginWarmupPromise;
    }

    function waitForStaffWarmupBriefly(maxMs = 1500) {
      if (!staffLoginWarmupPromise) return Promise.resolve();
      return Promise.race([
        staffLoginWarmupPromise,
        new Promise(resolve => setTimeout(resolve, maxMs))
      ]).then(() => undefined, () => undefined);
    }

    function escapeHtml(value) {
      return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
    }

    function statusBadge(status) {
      const text = String(status || '-');
      if (text.includes('ไม่ผ่าน') || text.includes('ปฏิเสธ')) {
        return `<span class="status-badge badge-fail">${escapeHtml(text)}</span>`;
      }
      if (text.includes('ผ่าน')) {
        return `<span class="status-badge badge-pass">${escapeHtml(text)}</span>`;
      }
      if (text.includes('พิจารณา')) {
        return `<span class="status-badge badge-info">${escapeHtml(text)}</span>`;
      }
      if (text.includes('รอ')) {
        return `<span class="status-badge badge-wait">${escapeHtml(text)}</span>`;
      }
      return `<span class="status-badge badge-muted">${escapeHtml(text)}</span>`;
    }

    function showLoading(title = 'กำลังโหลดข้อมูล') {
      Swal.fire({
        title,
        allowOutsideClick: false,
        didOpen: () => {
          Swal.showLoading();
        }
      });
    }

    function closeLoading() {
      Swal.close();
    }

    function hideBootView() {
      const boot = document.getElementById('bootView');
      if (boot) {
        boot.classList.add('d-none');
      }
    }

    function showLoginView() {
      hideBootView();
      document.getElementById('appView').classList.add('d-none');
      document.getElementById('loginView').classList.remove('d-none');
    }

    function openSidebar() {
      document.getElementById('sidebar').classList.add('open');
      document.getElementById('sidebarOverlay').classList.add('show');
    }

    function closeSidebar() {
      document.getElementById('sidebar').classList.remove('open');
      document.getElementById('sidebarOverlay').classList.remove('show');
    }

    function backToApplicant() {
      window.location.href = './index.html';
    }

    document.getElementById('loginForm').addEventListener('submit', async event => {
      event.preventDefault();
      const button = document.getElementById('loginBtn');
      const oldHtml = button.innerHTML;
      button.disabled = true;
      button.innerHTML = `<i class="fa-solid fa-spinner fa-spin me-2"></i> กำลังเข้าสู่ระบบ`;

      try {
        // ถ้า warm-up ที่เริ่มตอนเปิดหน้าใกล้เสร็จ ให้รอสั้น ๆ เพื่อหลีกเลี่ยงการยิง Apps Script ซ้อนกัน
        await waitForStaffWarmupBriefly(1500);

        const result = await serverCall(
          'staffLogin',
          document.getElementById('username').value.trim(),
          document.getElementById('password').value
        );

        if (!result || !result.success) {
          throw new Error(result && result.message ? result.message : 'เข้าสู่ระบบไม่สำเร็จ');
        }

        saveStaffSessionFromLoginResult(
          result,
          document.getElementById('username').value.trim()
        );

        sessionStorage.removeItem('ubuStaffView');
        enterApp();
      } catch (error) {
        Swal.fire({ icon: 'error', title: 'เข้าสู่ระบบไม่สำเร็จ', text: error.message });
      } finally {
        button.disabled = false;
        button.innerHTML = oldHtml;
      }
    });

    function applySession(session) {
      state.token = session.token || '';
      state.role = session.role || '';
      state.department = session.department || '';
      state.selectionUnit = session.selectionUnit || session.department || '';
      state.displayName = session.displayName || '';
      state.username = session.username || '';
    }

    function saveStaffSessionFromLoginResult(result, usernameFallback='') {
      applySession({
        token: result && result.token,
        role: result && result.role,
        department: result && result.department || '',
        selectionUnit: result && (result.selectionUnit || result.department) || '',
        displayName: result && result.displayName || '',
        username: result && result.username || usernameFallback
      });

      sessionStorage.setItem('ubuStaffSession', JSON.stringify({
        token: state.token,
        role: state.role,
        department: state.department,
        selectionUnit: state.selectionUnit,
        displayName: state.displayName,
        username: state.username
      }));
    }

    function enterApp() {
      hideBootView();
      document.getElementById('loginView').classList.add('d-none');
      document.getElementById('appView').classList.remove('d-none');
      resetClientIdleTimer();
      document.getElementById('sidebarUserName').textContent = state.displayName || state.username;

      document.getElementById('sidebarRole').textContent =
        state.role === 'admin'
          ? 'สิทธิ์ผู้ดูแลระบบ'
          : `สิทธิ์หน่วยงาน · ${state.selectionUnit || state.department}`;

      renderSidebar();
      const savedView = sessionStorage.getItem('ubuStaffView');
      const adminViews = ['adminDashboard', 'staffAccounts', 'contentManagement', 'applicantReview', 'civilRegistry', 'qualification', 'forwarding', 'adminResults', 'announcementPdf'];
      const departmentViews = ['departmentDashboard', 'departmentSelection', 'departmentBasket'];

      const targetView = state.role === 'admin'
        ? (adminViews.includes(savedView) ? savedView : 'adminDashboard')
        : (departmentViews.includes(savedView) ? savedView : 'departmentDashboard');

      // เมนูปัจจุบันให้โหลดก่อน จากนั้น preload ทุกเมนูที่เหลือแบบ background
      // หน่วงเล็กน้อยเพื่อให้ request ของหน้าที่ผู้ใช้กำลังเห็นได้ priority ก่อน
      Promise.resolve().then(() => navigate(targetView));
      clearTimeout(staffMenuPreloadRefreshTimer);
      staffMenuPreloadRefreshTimer = setTimeout(() => preloadAllStaffMenus(targetView), 250);
    }

    function renderSidebar() {
      const adminItems = [
        ['adminDashboard', 'fa-chart-line', 'ภาพรวมข้อมูลระบบกลาง'],
        ['staffAccounts', 'fa-user-gear', 'กำหนดสิทธิ์ผู้ใช้งาน S-MIS'],
        ['contentManagement', 'fa-bullhorn', 'จัดการข่าวประกาศ/เอกสาร'],
        ['applicantReview', 'fa-address-card', 'ตรวจสอบข้อมูลรายชื่อผู้สมัคร'],
        ['civilRegistry', 'fa-id-card', 'ตรวจสอบข้อมูลบัตรประชาชน'],
        ['qualification', 'fa-user-check', 'ตรวจสอบคุณสมบัติเกรดเฉลี่ย'],
        ['forwarding', 'fa-paper-plane', 'ส่งรายชื่อให้หน่วยงานคัดเลือก'],
        ['adminResults', 'fa-inbox', 'รับข้อมูลส่งกลับจากหน่วยงาน'],
        ['announcementPdf', 'fa-file-pdf', 'จัดทำประกาศรายชื่อ']
      ];
      const departmentItems = [
        ['departmentDashboard', 'fa-chart-pie', 'ภาพรวมข้อมูลหน่วยงาน'],
        ['departmentSelection', 'fa-user-plus', 'คัดเลือกรายชื่อลงตะกร้า'],
        ['departmentBasket', 'fa-basket-shopping', 'ยืนยันส่งรายชื่อที่คัดเลือก']
      ];

      const items = state.role === 'admin' ? adminItems : departmentItems;

      document.getElementById('sidebarMenu').innerHTML = `
        <div class="nav-label">เมนูการทำงาน</div>
        ${items.map(([key, icon, label]) => `
          <button class="side-link" data-view="${key}" onclick="navigate('${key}')">
            <span class="icon"><i class="fa-solid ${icon}"></i></span>
            <span>${escapeHtml(label)}</span>
            ${key === 'departmentBasket' ? `<span id="sidebarBasketBadge" class="menu-count d-none">0</span>` : ''}
          </button>
        `).join('')}
      `;
    }

    async function navigate(view) {
      state.currentView = view;
      state.photoLoadGeneration = Number(state.photoLoadGeneration || 0) + 1;
      sessionStorage.setItem('ubuStaffView', view);
      closeSidebar();

      document.querySelectorAll('.side-link').forEach(button => {
        button.classList.toggle('active', button.dataset.view === view);
      });

      try {
        if (view === 'adminDashboard') return await renderAdminDashboard();
        if (view === 'staffAccounts') return await renderStaffAccounts();
        if (view === 'contentManagement') return await renderContentManagement();
        if (view === 'applicantReview') return await renderApplicantReview();
        if (view === 'civilRegistry') return await renderCivilRegistry();
        if (view === 'qualification') return await renderQualification();
        if (view === 'forwarding') return await renderForwarding();
        if (view === 'adminResults') return await renderAdminResults();
        if (view === 'announcementPdf') return await renderAnnouncementExcel();
        if (view === 'departmentDashboard') return await renderDepartmentDashboard();
        if (view === 'departmentSelection') return await renderDepartmentSelection();
        if (view === 'departmentBasket') return await renderDepartmentBasket();
      } catch (error) {
        if (error.message.includes('เซสชัน')) {
          await forceLogout(error.message);
          return;
        }
        const slow = /ตอบสนองช้า|ประมวลผลนาน|ใช้เวลานาน/i.test(String(error && error.message || ''));
        Swal.fire({
          icon: slow ? 'warning' : 'error',
          title: slow ? 'ระบบกำลังประมวลผลข้อมูลจำนวนมาก' : 'ไม่สามารถโหลดข้อมูลได้',
          text: error.message,
          confirmButtonText: 'ตกลง'
        });
      }
    }

    function setHeader(title, subtitle = '') {
      document.getElementById('topTitle').textContent = title;
      document.getElementById('topSubtitle').textContent = subtitle;
    }

    function statCard(icon, number, label) {
      return `
        <div class="col-6 col-xl-3">
          <div class="stat-card">
            <div class="d-flex justify-content-between align-items-start gap-3">
              <div>
                <div class="stat-number">${Number(number || 0).toLocaleString('th-TH')}</div>
                <div class="stat-label">${escapeHtml(label)}</div>
              </div>
              <div class="stat-icon"><i class="fa-solid ${icon}"></i></div>
            </div>
          </div>
        </div>
      `;
    }


    function dashboardPercent(value) {
      const n = Number(value || 0);
      return `${Math.max(0, Math.min(100, n)).toFixed(n % 1 ? 1 : 0)}%`;
    }

    function dashboardMetric(icon, value, label, meta = '', metaClass = '', theme = 'yellow') {
      return `
        <div class="col-12 col-sm-6 col-xl-3">
          <div class="intel-kpi-card wm-card wm-${escapeHtml(theme)}">
            <div class="wm-bg-icon"><i class="fa-solid ${icon}"></i></div>
            <div class="intel-kpi-top">
              <div class="intel-kpi-label">${escapeHtml(label)}</div>
              <div class="intel-kpi-icon"><i class="fa-solid ${icon}"></i></div>
            </div>
            <div class="intel-kpi-value">${Number(value || 0).toLocaleString('th-TH')}</div>
            <div class="intel-kpi-meta ${metaClass}">${meta}</div>
          </div>
        </div>`;
    }

    function buildAdminInsights(data) {
      const cards = data.cards || {};
      const analytics = data.analytics || {};
      const quality = data.quality || {};
      const top = analytics.topDepartment || null;
      const growth = analytics.growth7d;
      const total = Number(cards.totalApplicants || 0);
      const pending = Number(quality.pendingQualification || 0);
      const pendingRate = total ? Math.round((pending / total) * 1000) / 10 : 0;
      const top3Share = Number(analytics.top3Share || 0);

      const trendText = growth === null
        ? `7 วันล่าสุดมี <strong>${Number(cards.last7Days || 0).toLocaleString('th-TH')} คน</strong> แต่ช่วง 7 วันก่อนหน้าไม่มีฐานเปรียบเทียบ`
        : growth > 0
          ? `ยอดสมัคร 7 วันล่าสุด <strong>เพิ่มขึ้น ${Math.abs(growth).toLocaleString('th-TH')}%</strong> เทียบกับ 7 วันก่อนหน้า`
          : growth < 0
            ? `ยอดสมัคร 7 วันล่าสุด <strong>ลดลง ${Math.abs(growth).toLocaleString('th-TH')}%</strong> เทียบกับ 7 วันก่อนหน้า`
            : `ยอดสมัคร 7 วันล่าสุด <strong>ทรงตัว</strong> เมื่อเทียบกับ 7 วันก่อนหน้า`;

      return [
        {
          icon: 'fa-ranking-star',
          title: 'หน่วยงานนำ',
          text: top
            ? `<strong>${escapeHtml(top.label || top.department || '-')}</strong> มีผู้สมัครสูงสุด ${Number(top.count || 0).toLocaleString('th-TH')} คน คิดเป็น ${Number(top.share || 0).toLocaleString('th-TH')}% ของทั้งระบบ`
            : 'ยังไม่มีข้อมูลเพียงพอสำหรับจัดอันดับหน่วยงาน'
        },
        { icon: 'fa-arrow-trend-up', title: 'Momentum 7 วัน', text: trendText },
        {
          icon: 'fa-magnifying-glass-chart',
          title: 'ภาระงานตรวจสอบ',
          text: pending
            ? `ยังรอตรวจคุณสมบัติ <strong>${pending.toLocaleString('th-TH')} คน (${pendingRate.toLocaleString('th-TH')}%)</strong> ควรเร่งเคลียร์ก่อนส่งรายชื่อให้หน่วยงาน`
            : '<strong>ไม่มีรายการค้างตรวจคุณสมบัติ</strong> ในข้อมูลปัจจุบัน'
        },
        {
          icon: 'fa-chart-pie',
          title: 'การกระจุกตัว',
          text: total
            ? `3 หน่วยงานแรกครองสัดส่วนรวม <strong>${top3Share.toLocaleString('th-TH')}%</strong> ${top3Share >= 60 ? 'ถือว่าการสมัครกระจุกตัวค่อนข้างสูง' : 'การกระจายตัวยังไม่กระจุกเฉพาะไม่กี่หน่วยงานมากเกินไป'}`
            : 'ยังไม่มีข้อมูลสำหรับวิเคราะห์การกระจุกตัว'
        }
      ];
    }

    function buildAdminTrendChart(points) {
      const data = Array.isArray(points) ? points : [];
      if (!data.length) return '<div class="text-muted text-center py-5">ยังไม่มีข้อมูลแนวโน้ม</div>';
      const width = 760;
      const height = 238;
      const left = 38;
      const right = 20;
      const top = 20;
      const bottom = 40;
      const chartW = width - left - right;
      const chartH = height - top - bottom;
      const maxValue = Math.max(1, ...data.map(x => Number(x.count || 0)));
      const stepX = data.length > 1 ? chartW / (data.length - 1) : chartW;
      const yFor = value => top + chartH - (Number(value || 0) / maxValue * chartH);
      const coords = data.map((item, index) => ({
        x: left + (index * stepX),
        y: yFor(item.count),
        count: Number(item.count || 0),
        label: item.label || ''
      }));
      const line = coords.map(p => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
      const baseY = top + chartH;
      const area = `${left},${baseY} ${line} ${coords[coords.length - 1].x.toFixed(1)},${baseY}`;
      const hottest = Math.max(...coords.map(p => p.count));
      const grid = [0, .25, .5, .75, 1].map(r => {
        const y = top + chartH - (r * chartH);
        const val = Math.round(maxValue * r);
        return `<line class="trend-grid-line" x1="${left}" y1="${y}" x2="${width-right}" y2="${y}"></line><text class="trend-axis-label" x="4" y="${y+4}">${val}</text>`;
      }).join('');
      const labels = coords.map((p, index) => index % 2 === 0 || index === coords.length - 1
        ? `<text class="trend-axis-label" text-anchor="middle" x="${p.x}" y="${height-13}">${escapeHtml(p.label)}</text>`
        : '').join('');
      const dots = coords.map(p => `<circle class="${p.count === hottest && hottest > 0 ? 'trend-point-hot' : 'trend-point'}" cx="${p.x}" cy="${p.y}" r="4.5"><title>${escapeHtml(p.label)}: ${p.count} คน</title></circle>`).join('');
      return `<svg class="trend-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="แนวโน้มจำนวนผู้สมัคร 14 วัน">${grid}<polygon class="trend-area" points="${area}"></polygon><polyline class="trend-line" points="${line}"></polyline>${dots}${labels}</svg>`;
    }

    function buildAdminFunnel(funnel) {
      const rows = Array.isArray(funnel) ? funnel : [];
      const max = Math.max(1, Number(rows[0]?.count || 0));
      return rows.length ? rows.map((item, index) => {
        const count = Number(item.count || 0);
        const rate = max ? count / max * 100 : 0;
        return `
          <div class="funnel-step">
            <div class="funnel-label-row">
              <div class="funnel-label">${index + 1}. ${escapeHtml(item.label || '-')}</div>
              <div class="funnel-value">${count.toLocaleString('th-TH')} <span class="text-muted fw-normal">(${rate.toFixed(1)}%)</span></div>
            </div>
            <div class="funnel-track"><div class="funnel-fill" style="width:${Math.max(count ? 3 : 0, rate)}%"></div></div>
          </div>`;
      }).join('') : '<div class="text-muted text-center py-4">ยังไม่มีข้อมูล</div>';
    }

    function buildDepartmentIntelligence(stats) {
      const rows = Array.isArray(stats) ? stats : [];
      const max = Math.max(1, ...rows.map(x => Number(x.count || 0)));
      if (!rows.length) return '<div class="text-muted text-center py-5">ยังไม่มีข้อมูลการสมัคร</div>';
      return rows.map((item, index) => {
        const count = Number(item.count || 0);
        const relative = count / max * 100;
        const pending = Number(item.pendingReview || 0);
        return `
          <div class="dept-intel-row">
            <div class="dept-rank ${index < 3 ? 'top' : ''}">${index + 1}</div>
            <div>
              <div class="dept-name">${escapeHtml(item.label || item.department || '-')}</div>
              <div class="d-flex flex-wrap gap-1 mt-1">
                <span class="analysis-badge">${count.toLocaleString('th-TH')} คน · ${Number(item.share || 0).toFixed(1)}%</span>
                ${pending ? `<span class="analysis-badge warn">ค้างตรวจ ${pending.toLocaleString('th-TH')}</span>` : ''}
              </div>
            </div>
            <div class="dept-progress-cell">
              <div class="dept-metric-label">ยอดสมัครเทียบอันดับ 1</div>
              <div class="dept-progress-line"><div class="dept-progress-fill" style="width:${relative}%"></div></div>
            </div>
            <div class="dept-forward-metric">
              <div class="dept-metric-label">ความคืบหน้า</div>
              <div class="dept-metric-value">ตรวจ ${Number(item.reviewRate || 0).toFixed(0)}% · ส่ง ${Number(item.forwardRate || 0).toFixed(0)}%</div>
            </div>
          </div>`;
      }).join('');
    }

    function buildFacultyDistribution(rows) {
      const data = Array.isArray(rows) ? rows : [];
      const max = Math.max(1, ...data.map(x => Number(x.count || 0)));
      return data.length ? data.map(item => `
        <div class="distribution-row">
          <div class="distribution-head">
            <div class="distribution-name" title="${escapeHtml(item.faculty || '')}">${escapeHtml(item.faculty || '-')}</div>
            <div class="distribution-value">${Number(item.count || 0).toLocaleString('th-TH')} คน · ${Number(item.share || 0).toLocaleString('th-TH')}%</div>
          </div>
          <div class="distribution-track"><div class="distribution-fill" style="width:${Number(item.count || 0) / max * 100}%"></div></div>
        </div>`).join('') : '<div class="text-muted text-center py-4">ยังไม่มีข้อมูลคณะ</div>';
    }

    const publicContentTypeLabels = {
      news: 'ข่าวประกาศ',
      rules: 'ระเบียบการรับสมัคร',
      manual: 'คู่มือการสมัคร'
    };

    function localDateInputValue(date = new Date()) {
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const d = String(date.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }

    function safeDriveLink(value) {
      try {
        const url = new URL(String(value || ''));
        const host = url.hostname.toLowerCase();
        if (url.protocol !== 'https:') return '';
        if (host === 'drive.google.com' || host === 'docs.google.com' || host.endsWith('.googleusercontent.com')) return url.href;
      } catch (_) {}
      return '';
    }

    async function renderContentManagement() {
      setHeader('จัดการข่าวประกาศ/เอกสาร', 'จัดการข้อมูลที่แสดงบนหน้าเว็บไซต์');
      const content = document.getElementById('content');
      content.innerHTML = `<div class="text-center py-5"><div class="spinner-border text-primary"></div></div>`;

      state.publicContentRows = await serverCall('getAdminPublicContents', state.token) || [];

      content.innerHTML = `
        <div class="d-flex flex-wrap justify-content-between gap-3 align-items-end mb-4">
          <div>
            <div class="page-title">ข่าวประกาศและเอกสารหน้าเว็บ</div>
            <div class="page-subtitle">ข้อมูลจะบันทึกลงชีต Announcements และแสดงบนหน้าเว็บไซต์อัตโนมัติ</div>
          </div>
          <button class="btn btn-outline-secondary btn-sm" onclick="renderContentManagement()">
            <i class="fa-solid fa-rotate me-1"></i>อัปเดตข้อมูล
          </button>
        </div>

        <div class="panel mb-4">
          <div class="panel-header">
            <h3 class="panel-title"><i class="fa-solid fa-pen-to-square me-2"></i>เพิ่ม / แก้ไขข้อมูล</h3>
          </div>
          <div class="p-3 p-md-4">
            <input type="hidden" id="publicContentId">
            <div class="row g-3">
              <div class="col-md-4">
                <label class="form-label">ประเภท <span class="text-danger">*</span></label>
                <select id="publicContentType" class="form-select">
                  <option value="news">ข่าวประกาศ</option>
                  <option value="rules">ระเบียบการรับสมัคร</option>
                  <option value="manual">คู่มือการสมัคร</option>
                </select>
              </div>
              <div class="col-md-4">
                <label class="form-label">วันที่ประกาศ <span class="text-danger">*</span></label>
                <input id="publicContentDate" type="date" class="form-control" value="${localDateInputValue()}">
              </div>
              <div class="col-md-4">
                <label class="form-label">สถานะการแสดง</label>
                <select id="publicContentActive" class="form-select">
                  <option value="true">แสดงบนหน้าเว็บ</option>
                  <option value="false">ซ่อนชั่วคราว</option>
                </select>
              </div>
              <div class="col-12">
                <div class="border rounded-3 p-3 bg-light">
                  <div class="fw-semibold mb-2"><i class="fa-solid fa-bullhorn me-1"></i>การประชาสัมพันธ์ข่าวบนหน้าแรก <span class="text-muted fw-normal" style="font-size:12px">(ใช้กับประเภทข่าวประกาศ)</span></div>
                  <div class="row g-3">
                    <div class="col-md-4">
                      <div class="form-check form-switch">
                        <input class="form-check-input" type="checkbox" id="publicContentShowHome" checked>
                        <label class="form-check-label" for="publicContentShowHome">แสดงในข่าวล่าสุดหน้าแรก</label>
                      </div>
                    </div>
                    <div class="col-md-4">
                      <div class="form-check form-switch">
                        <input class="form-check-input" type="checkbox" id="publicContentShowPopup">
                        <label class="form-check-label" for="publicContentShowPopup">แสดง Popup เมื่อเข้าเว็บไซต์</label>
                      </div>
                    </div>
                    <div class="col-md-4">
                      <div class="form-check form-switch">
                        <input class="form-check-input" type="checkbox" id="publicContentImportant">
                        <label class="form-check-label" for="publicContentImportant">ข่าวสำคัญ / แถบประกาศด่วน</label>
                      </div>
                    </div>
                    <div class="col-md-6">
                      <label class="form-label mb-1">เริ่มแสดง</label>
                      <input id="publicContentDisplayStart" type="date" class="form-control">
                    </div>
                    <div class="col-md-6">
                      <label class="form-label mb-1">สิ้นสุดการแสดง</label>
                      <input id="publicContentDisplayEnd" type="date" class="form-control">
                    </div>
                  </div>
                  <div class="form-text mt-2">หากไม่กำหนดวันเริ่ม/สิ้นสุด ข่าวจะแสดงตามสถานะจนกว่าจะซ่อนเอง</div>
                </div>
              </div>
              <div class="col-12">
                <label class="form-label">เรื่อง <span class="text-danger">*</span></label>
                <input id="publicContentSubject" class="form-control" maxlength="300" placeholder="ระบุชื่อข่าวประกาศหรือชื่อเอกสาร">
              </div>
              <div class="col-12">
                <label class="form-label">รายละเอียด</label>
                <textarea id="publicContentDetails" class="form-control" rows="5" maxlength="10000" placeholder="ระบุรายละเอียด (ถ้ามี)"></textarea>
              </div>
              <div class="col-md-8">
                <label class="form-label">ไฟล์แนบ</label>
                <input id="publicContentFile" type="file" class="form-control" accept=".pdf,.jpg,.jpeg,.png,application/pdf,image/jpeg,image/png">
                <div class="form-text">รองรับ PDF, JPG/JPEG, PNG ขนาดไม่เกิน 8 MB · ระเบียบการรับสมัครและคู่มือการสมัครต้องแนบไฟล์</div>
                <div id="publicContentExistingFile" class="small mt-2"></div>
              </div>
              <div class="col-md-4">
                <label class="form-label">ลำดับ</label>
                <input id="publicContentSortOrder" type="number" min="0" max="99999" class="form-control" value="0">
              </div>
              <div class="col-12 d-flex flex-wrap gap-2">
                <button class="btn btn-primary" onclick="savePublicContentForm()">
                  <i class="fa-solid fa-floppy-disk me-1"></i>บันทึกข้อมูล
                </button>
                <button class="btn btn-outline-secondary" onclick="resetPublicContentForm()">
                  <i class="fa-solid fa-rotate-left me-1"></i>ล้างแบบฟอร์ม
                </button>
              </div>
            </div>
          </div>
        </div>

        <div class="panel">
          <div class="panel-header">
            <h3 class="panel-title">รายการที่บันทึกแล้ว</h3>
            <span class="text-muted">${state.publicContentRows.length.toLocaleString('th-TH')} รายการ</span>
          </div>
          <div class="table-wrap">
            <table class="table">
              <thead>
                <tr>
                  <th>วันที่ประกาศ</th>
                  <th>ประเภท</th>
                  <th>เรื่อง</th>
                  <th>ไฟล์แนบ</th>
                  <th class="text-center">ผู้อ่าน</th>
                  <th>สถานะ</th>
                  <th class="text-end">จัดการ</th>
                </tr>
              </thead>
              <tbody>
                ${renderPublicContentAdminRows(state.publicContentRows)}
              </tbody>
            </table>
          </div>
        </div>`;
    }

    function renderPublicContentAdminRows(rows) {
      if (!rows.length) return `<tr><td colspan="7" class="text-center text-muted py-5">ยังไม่มีข้อมูล</td></tr>`;

      return rows.map(item => {
        const fileUrl = safeDriveLink(item.fileUrl);
        return `
          <tr>
            <td class="text-nowrap">${escapeHtml(item.publishDate || '-')}</td>
            <td><span class="status-badge badge-info">${escapeHtml(item.typeLabel || publicContentTypeLabels[item.type] || '-')}</span></td>
            <td>
              <div class="fw-medium">${escapeHtml(item.subject || '-')}</div>
              ${item.details ? `<div class="text-muted mt-1" style="font-size:12px;max-width:520px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(item.details)}</div>` : ''}
            </td>
            <td>${fileUrl ? `<a href="${escapeHtml(fileUrl)}" target="_blank" rel="noopener noreferrer" class="btn btn-outline-primary btn-sm"><i class="fa-solid fa-paperclip me-1"></i>เปิดไฟล์</a>` : '<span class="text-muted">-</span>'}</td>
            <td class="text-center text-nowrap">${item.type === 'news' ? `<span class="status-badge badge-info" title="นับโดยประมาณ 1 ครั้งต่อข่าวต่ออุปกรณ์/เบราว์เซอร์"><i class="fa-regular fa-eye me-1"></i>${Math.max(0, Number(item.viewCount) || 0).toLocaleString('th-TH')} คน</span>` : '<span class="text-muted">-</span>'}</td>
            <td>
              ${item.active ? '<span class="status-badge badge-pass">แสดง</span>' : '<span class="status-badge badge-muted">ซ่อน</span>'}
              ${item.type === 'news' && item.showHome ? '<span class="status-badge badge-info ms-1">หน้าแรก</span>' : ''}
              ${item.type === 'news' && item.showPopup ? '<span class="status-badge badge-warning ms-1">Popup</span>' : ''}
              ${item.type === 'news' && item.important ? '<span class="status-badge badge-danger ms-1">สำคัญ</span>' : ''}
            </td>
            <td class="text-end text-nowrap">
              <button class="btn btn-outline-primary btn-sm me-1" onclick="editPublicContent('${escapeHtml(item.id)}')"><i class="fa-solid fa-pen"></i></button>
              <button class="btn btn-outline-danger btn-sm" onclick="deletePublicContentItem('${escapeHtml(item.id)}')"><i class="fa-solid fa-trash"></i></button>
            </td>
          </tr>`;
      }).join('');
    }

    function resetPublicContentForm() {
      const id = document.getElementById('publicContentId');
      if (!id) return;
      id.value = '';
      document.getElementById('publicContentType').value = 'news';
      document.getElementById('publicContentDate').value = localDateInputValue();
      document.getElementById('publicContentActive').value = 'true';
      document.getElementById('publicContentSubject').value = '';
      document.getElementById('publicContentDetails').value = '';
      document.getElementById('publicContentFile').value = '';
      document.getElementById('publicContentSortOrder').value = '0';
      document.getElementById('publicContentShowHome').checked = true;
      document.getElementById('publicContentShowPopup').checked = false;
      document.getElementById('publicContentImportant').checked = false;
      document.getElementById('publicContentDisplayStart').value = '';
      document.getElementById('publicContentDisplayEnd').value = '';
      document.getElementById('publicContentExistingFile').innerHTML = '';
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function editPublicContent(id) {
      const item = state.publicContentRows.find(row => row.id === id);
      if (!item) return;

      document.getElementById('publicContentId').value = item.id || '';
      document.getElementById('publicContentType').value = item.type || 'news';
      document.getElementById('publicContentDate').value = item.publishDateInput || localDateInputValue();
      document.getElementById('publicContentActive').value = item.active ? 'true' : 'false';
      document.getElementById('publicContentSubject').value = item.subject || '';
      document.getElementById('publicContentDetails').value = item.details || '';
      document.getElementById('publicContentFile').value = '';
      document.getElementById('publicContentSortOrder').value = String(item.sortOrder || 0);
      document.getElementById('publicContentShowHome').checked = item.showHome !== false;
      document.getElementById('publicContentShowPopup').checked = item.showPopup === true;
      document.getElementById('publicContentImportant').checked = item.important === true;
      document.getElementById('publicContentDisplayStart').value = item.displayStartInput || '';
      document.getElementById('publicContentDisplayEnd').value = item.displayEndInput || '';

      const fileUrl = safeDriveLink(item.fileUrl);
      document.getElementById('publicContentExistingFile').innerHTML = fileUrl
        ? `<div class="alert alert-light border py-2 px-3 mb-0"><i class="fa-solid fa-paperclip me-1"></i>ไฟล์ปัจจุบัน: <a href="${escapeHtml(fileUrl)}" target="_blank" rel="noopener noreferrer">${escapeHtml(item.fileName || 'เปิดไฟล์')}</a> <button class="btn btn-link btn-sm text-danger p-0 ms-2" type="button" onclick="markPublicContentFileForRemoval()">นำไฟล์ออก</button><input id="publicContentRemoveFile" type="hidden" value="false"></div>`
        : '<input id="publicContentRemoveFile" type="hidden" value="false">';

      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function markPublicContentFileForRemoval() {
      const input = document.getElementById('publicContentRemoveFile');
      if (input) input.value = 'true';
      document.getElementById('publicContentExistingFile').innerHTML = '<div class="text-danger small"><i class="fa-solid fa-circle-minus me-1"></i>ไฟล์เดิมจะถูกนำออกเมื่อกดบันทึก</div><input id="publicContentRemoveFile" type="hidden" value="true">';
    }

    function readPublicContentFile(file) {
      return new Promise((resolve, reject) => {
        if (!file) return resolve(null);
        const maxBytes = 8 * 1024 * 1024;
        if (file.size <= 0 || file.size > maxBytes) return reject(new Error('ไฟล์แนบต้องมีขนาดไม่เกิน 8 MB'));
        if (!/\.(pdf|jpe?g|png)$/i.test(file.name || '')) return reject(new Error('รองรับเฉพาะ PDF, JPG/JPEG หรือ PNG'));

        const reader = new FileReader();
        reader.onload = () => {
          const result = String(reader.result || '');
          const comma = result.indexOf(',');
          resolve({
            name: file.name,
            mimeType: file.type || '',
            size: file.size,
            data: comma >= 0 ? result.slice(comma + 1) : result
          });
        };
        reader.onerror = () => reject(new Error('ไม่สามารถอ่านไฟล์แนบได้'));
        reader.readAsDataURL(file);
      });
    }

    async function savePublicContentForm() {
      const type = document.getElementById('publicContentType').value;
      const publishDate = document.getElementById('publicContentDate').value;
      const subject = document.getElementById('publicContentSubject').value.trim();
      const details = document.getElementById('publicContentDetails').value.trim();
      const fileInput = document.getElementById('publicContentFile');
      const existingId = document.getElementById('publicContentId').value.trim();
      const removeInput = document.getElementById('publicContentRemoveFile');

      if (!publishDate || !subject) {
        Swal.fire({ icon: 'warning', title: 'ข้อมูลไม่ครบ', text: 'กรุณาระบุวันที่ประกาศและเรื่อง' });
        return;
      }

      const existing = existingId ? state.publicContentRows.find(row => row.id === existingId) : null;
      const hasRetainedFile = existing && existing.fileUrl && !(removeInput && removeInput.value === 'true');
      if (type === 'rules' && !fileInput.files.length && !hasRetainedFile) {
        Swal.fire({ icon: 'warning', title: 'กรุณาแนบไฟล์ประกาศ', text: 'ระเบียบการรับสมัครต้องมีไฟล์ประกาศแนบ' });
        return;
      }
      if (type === 'manual' && !fileInput.files.length && !hasRetainedFile) {
        Swal.fire({ icon: 'warning', title: 'กรุณาแนบไฟล์คู่มือ', text: 'เมนูคู่มือการสมัครจะแสดงเป็นไฟล์แนบ จึงต้องมีไฟล์คู่มือ' });
        return;
      }

      const displayStart = document.getElementById('publicContentDisplayStart').value;
      const displayEnd = document.getElementById('publicContentDisplayEnd').value;
      if (displayStart && displayEnd && displayStart > displayEnd) {
        Swal.fire({ icon: 'warning', title: 'ช่วงวันที่ไม่ถูกต้อง', text: 'วันที่เริ่มแสดงต้องไม่มากกว่าวันที่สิ้นสุดการแสดง' });
        return;
      }

      try {
        showLoading('กำลังบันทึกข้อมูล');
        const file = fileInput.files.length ? await readPublicContentFile(fileInput.files[0]) : null;
        const result = await serverCall('savePublicContent', state.token, {
          id: existingId,
          type,
          publishDate,
          subject,
          details,
          active: document.getElementById('publicContentActive').value === 'true',
          sortOrder: Number(document.getElementById('publicContentSortOrder').value) || 0,
          showHome: document.getElementById('publicContentShowHome').checked,
          showPopup: document.getElementById('publicContentShowPopup').checked,
          important: document.getElementById('publicContentImportant').checked,
          displayStart: document.getElementById('publicContentDisplayStart').value,
          displayEnd: document.getElementById('publicContentDisplayEnd').value,
          removeFile: Boolean(removeInput && removeInput.value === 'true'),
          file
        });
        closeLoading();
        await Swal.fire({ icon: 'success', title: 'บันทึกเรียบร้อย', text: result.message || 'บันทึกข้อมูลเรียบร้อยแล้ว' });
        await renderContentManagement();
      } catch (error) {
        closeLoading();
        Swal.fire({ icon: 'error', title: 'บันทึกไม่สำเร็จ', text: error.message });
      }
    }

    async function deletePublicContentItem(id) {
      const item = state.publicContentRows.find(row => row.id === id);
      const confirm = await Swal.fire({
        icon: 'warning',
        title: 'ลบรายการนี้?',
        html: `ต้องการลบ <strong>${escapeHtml(item?.subject || '')}</strong> หรือไม่?<br><small class="text-muted">ไฟล์ที่ระบบอัปโหลดไว้จะถูกย้ายไปถังขยะด้วย</small>`,
        showCancelButton: true,
        confirmButtonText: 'ลบรายการ',
        cancelButtonText: 'ยกเลิก',
        confirmButtonColor: '#b4424d'
      });
      if (!confirm.isConfirmed) return;

      try {
        showLoading('กำลังลบข้อมูล');
        await serverCall('deletePublicContent', state.token, id);
        closeLoading();
        await renderContentManagement();
      } catch (error) {
        closeLoading();
        Swal.fire({ icon: 'error', title: 'ลบไม่สำเร็จ', text: error.message });
      }
    }

    async function renderStaffAccounts() {
      setHeader('กำหนดสิทธิ์ผู้ใช้งาน S-MIS', 'ใช้บัญชีและรหัสผ่านเดียวกับ S-MIS โดยกำหนดสิทธิ์ของระบบนี้แยกตามรหัสผู้ใช้งาน');
      const content = document.getElementById('content');
      content.innerHTML = `<div class="text-center py-5"><div class="spinner-border text-primary"></div></div>`;

      const data = await serverCall('getStaffAccountManagement', state.token);
      state.staffAccountRows = Array.isArray(data && data.accounts) ? data.accounts : [];
      state.staffAccountUnits = Array.isArray(data && data.units) ? data.units : [];

      const departmentOptions = [...new Set(
        state.staffAccountUnits.map(item => String(item.department || '').trim()).filter(Boolean)
      )].sort((a, b) => a.localeCompare(b, 'th'));

      const selectionOptions = [...new Set(
        state.staffAccountUnits.map(item => String(item.selectionUnit || '').trim()).filter(Boolean)
      )].sort((a, b) => a.localeCompare(b, 'th'));

      content.innerHTML = `
        <div class="d-flex flex-wrap justify-content-between gap-3 align-items-end mb-4">
          <div>
            <div class="page-title">กำหนดสิทธิ์ผู้ใช้งาน S-MIS</div>
            <div class="page-subtitle">ผู้ใช้เข้าสู่ระบบด้วยรหัสผู้ใช้งานและรหัสผ่านเดียวกับ S-MIS · ระบบนี้เก็บเฉพาะสิทธิ์และหน่วยงาน</div>
          </div>
          <div class="text-muted small"><i class="fa-solid fa-users-gear me-1"></i>ได้รับสิทธิ์แล้ว <strong>${state.staffAccountRows.length.toLocaleString('th-TH')}</strong> บัญชี</div>
        </div>

        <div class="alert alert-info d-flex gap-3 align-items-start mb-3" role="alert">
          <i class="fa-solid fa-circle-info mt-1"></i>
          <div><strong>Single Account:</strong> ต้องมีรหัสผู้ใช้งานอยู่ใน S-MIS ก่อน จากนั้นผู้ดูแลระบบจึงเพิ่มรหัสนั้นเข้าหน้านี้เพื่อกำหนดสิทธิ์ Admin หรือหน่วยงาน โดยไม่ต้องสร้างรหัสผ่านใหม่</div>
        </div>

        <div class="row g-3 align-items-stretch">
          <div class="col-12 col-xl-5">
            <div class="panel h-100">
              <div class="panel-header">
                <div>
                  <h2 class="panel-title"><i class="fa-solid fa-user-shield me-2"></i>เพิ่มสิทธิ์ผู้ใช้งาน</h2>
                  <div class="text-muted small mt-1">ระบบจะตรวจสอบว่ารหัสผู้ใช้งานมีอยู่จริงใน S-MIS ก่อนบันทึกสิทธิ์</div>
                </div>
              </div>
              <div class="p-3 p-lg-4">
                <form id="staffAccountForm" autocomplete="off">
                  <div class="row g-3">
                    <div class="col-12">
                      <label class="form-label">รหัสผู้ใช้งาน S-MIS <span class="text-danger">*</span></label>
                      <input id="staffAccountUsername" class="form-control" maxlength="80" autocomplete="off" required
                             placeholder="กรอกรหัสผู้ใช้งานเดียวกับ S-MIS">
                      <div class="form-text">ต้องเป็นรหัสที่มีอยู่ในชีต Users ของ S-MIS</div>
                    </div>

                    <div class="col-12">
                      <label class="form-label">ชื่อที่แสดง <span class="text-muted">(ไม่บังคับ)</span></label>
                      <input id="staffAccountDisplayName" class="form-control" maxlength="150"
                             placeholder="เว้นว่างเพื่อใช้ชื่อ-นามสกุลจาก S-MIS">
                    </div>

                    <div class="col-12">
                      <label class="form-label">ประเภทสิทธิ์ <span class="text-danger">*</span></label>
                      <select id="staffAccountRole" class="form-select" required onchange="toggleStaffAccountUnitFields()">
                        <option value="department" selected>เจ้าหน้าที่หน่วยงาน</option>
                        <option value="admin">ผู้ดูแลระบบ (Admin)</option>
                      </select>
                    </div>

                    <div id="staffAccountUnitFields" class="col-12">
                      <div class="row g-3">
                        <div class="col-12">
                          <label class="form-label">หน่วยงาน / กลุ่มงาน <span class="text-danger">*</span></label>
                          <input id="staffAccountDepartment" class="form-control" list="staffDepartmentList" maxlength="180"
                                 placeholder="เลือกหรือพิมพ์ชื่อหน่วยงาน">
                          <datalist id="staffDepartmentList">
                            ${departmentOptions.map(item => `<option value="${escapeHtml(item)}"></option>`).join('')}
                          </datalist>
                        </div>
                        <div class="col-12">
                          <label class="form-label">หน่วยคัดเลือก <span class="text-danger">*</span></label>
                          <input id="staffAccountSelectionUnit" class="form-control" list="staffSelectionUnitList" maxlength="180"
                                 placeholder="เช่น งานสวัสดิการนักศึกษา">
                          <datalist id="staffSelectionUnitList">
                            ${selectionOptions.map(item => `<option value="${escapeHtml(item)}"></option>`).join('')}
                          </datalist>
                          <div class="form-text">ควรกำหนดให้ตรงกับหน่วยคัดเลือกในชีต Jobs</div>
                        </div>
                      </div>
                    </div>

                    <div class="col-12 pt-1">
                      <button id="staffAccountSubmitBtn" class="btn btn-primary w-100 py-2" type="submit">
                        <i class="fa-solid fa-user-shield me-2"></i>บันทึกสิทธิ์ผู้ใช้งาน
                      </button>
                    </div>
                  </div>
                </form>
              </div>
            </div>
          </div>

          <div class="col-12 col-xl-7">
            <div class="panel h-100">
              <div class="panel-header">
                <div>
                  <h2 class="panel-title">ผู้ใช้งานที่ได้รับสิทธิ์</h2>
                  <div class="text-muted small mt-1">Authentication: S-MIS · Authorization: ระบบรับสมัครทำงานระหว่างเรียน</div>
                </div>
                <button class="btn btn-outline-secondary btn-sm" onclick="renderStaffAccounts()">
                  <i class="fa-solid fa-rotate me-1"></i>อัปเดตข้อมูล
                </button>
              </div>
              <div class="table-wrap">
                <table class="table align-middle staff-account-table">
                  <thead>
                    <tr>
                      <th style="width:55px">ลำดับ</th>
                      <th>รหัสผู้ใช้งาน S-MIS</th>
                      <th>ชื่อที่แสดง</th>
                      <th>สิทธิ์</th>
                      <th>หน่วยงาน / หน่วยคัดเลือก</th>
                    </tr>
                  </thead>
                  <tbody>${renderStaffAccountRows(state.staffAccountRows)}</tbody>
                </table>
              </div>
            </div>
          </div>
        </div>`;

      toggleStaffAccountUnitFields();
      document.getElementById('staffAccountForm').addEventListener('submit', submitStaffAccountForm);
    }

    function renderStaffAccountRows(rows) {
      if (!Array.isArray(rows) || !rows.length) {
        return `<tr><td colspan="5" class="text-center text-muted py-5">ยังไม่มีผู้ใช้งานที่ได้รับสิทธิ์ในระบบ</td></tr>`;
      }
      return rows.map((item, index) => {
        const role = String(item.role || '').toLowerCase();
        const roleBadge = role === 'admin'
          ? `<span class="status-badge badge-info"><i class="fa-solid fa-shield-halved me-1"></i>ผู้ดูแลระบบ</span>`
          : `<span class="status-badge badge-pass"><i class="fa-solid fa-building-user me-1"></i>เจ้าหน้าที่หน่วยงาน</span>`;
        const department = item.department || '-';
        const unit = item.selectionUnit || item.department || '-';
        return `<tr>
          <td class="text-muted">${index + 1}</td>
          <td><span class="candidate-code fw-semibold">${escapeHtml(item.username || '-')}</span><div class="text-muted table-subtext">S-MIS</div></td>
          <td>${escapeHtml(item.displayName || '-')}</td>
          <td>${roleBadge}</td>
          <td>${role === 'admin' ? `<span class="text-muted">สิทธิ์ส่วนกลาง</span>` : `<div class="staff-account-unit"><strong>${escapeHtml(department)}</strong><span>${escapeHtml(unit)}</span></div>`}</td>
        </tr>`;
      }).join('');
    }

    function toggleStaffAccountUnitFields() {
      const role = document.getElementById('staffAccountRole');
      const wrap = document.getElementById('staffAccountUnitFields');
      const department = document.getElementById('staffAccountDepartment');
      const selectionUnit = document.getElementById('staffAccountSelectionUnit');
      if (!role || !wrap || !department || !selectionUnit) return;
      const isDepartment = role.value === 'department';
      wrap.classList.toggle('d-none', !isDepartment);
      department.required = isDepartment;
      selectionUnit.required = isDepartment;
      if (!isDepartment) {
        department.value = '';
        selectionUnit.value = '';
      }
    }

    async function submitStaffAccountForm(event) {
      event.preventDefault();

      const username = document.getElementById('staffAccountUsername').value.trim();
      const displayName = document.getElementById('staffAccountDisplayName').value.trim();
      const role = document.getElementById('staffAccountRole').value;
      const department = document.getElementById('staffAccountDepartment').value.trim();
      const selectionUnit = document.getElementById('staffAccountSelectionUnit').value.trim();
      const button = document.getElementById('staffAccountSubmitBtn');

      if (!/^[A-Za-z0-9._-]{3,80}$/.test(username)) {
        Swal.fire({ icon: 'warning', title: 'รหัสผู้ใช้งานไม่ถูกต้อง', text: 'กรุณากรอกรหัสผู้ใช้งาน S-MIS ให้ถูกต้อง' });
        return;
      }
      if (role === 'department' && (!department || !selectionUnit)) {
        Swal.fire({ icon: 'warning', title: 'กรอกข้อมูลไม่ครบ', text: 'บัญชีเจ้าหน้าที่หน่วยงานต้องระบุหน่วยงานและหน่วยคัดเลือก' });
        return;
      }

      const confirm = await Swal.fire({
        icon: 'question',
        title: 'ยืนยันการกำหนดสิทธิ์',
        html: `ต้องการให้รหัส S-MIS <strong>${escapeHtml(username)}</strong> เข้าใช้งานระบบนี้ด้วยสิทธิ์ <strong>${role === 'admin' ? 'ผู้ดูแลระบบ' : 'เจ้าหน้าที่หน่วยงาน'}</strong> ใช่หรือไม่`,
        showCancelButton: true,
        confirmButtonText: 'ยืนยันกำหนดสิทธิ์',
        cancelButtonText: 'ยกเลิก'
      });
      if (!confirm.isConfirmed) return;

      const oldHtml = button.innerHTML;
      button.disabled = true;
      button.innerHTML = `<i class="fa-solid fa-spinner fa-spin me-2"></i>กำลังตรวจสอบ S-MIS`;

      try {
        const result = await serverCall('createStaffAccount', state.token, {
          username,
          role,
          department: role === 'department' ? department : '',
          selectionUnit: role === 'department' ? selectionUnit : '',
          displayName
        });

        if (!result || !result.success) {
          throw new Error(result && result.message ? result.message : 'ไม่สามารถกำหนดสิทธิ์ได้');
        }

        await Swal.fire({
          icon: 'success',
          title: 'กำหนดสิทธิ์สำเร็จ',
          text: result.message || 'ผู้ใช้งานสามารถเข้าสู่ระบบด้วยบัญชี S-MIS ได้แล้ว',
          confirmButtonText: 'ตกลง'
        });
        await renderStaffAccounts();
      } catch (error) {
        Swal.fire({ icon: 'error', title: 'กำหนดสิทธิ์ไม่สำเร็จ', text: error.message });
      } finally {
        if (document.body.contains(button)) {
          button.disabled = false;
          button.innerHTML = oldHtml;
        }
      }
    }

    async function renderAdminDashboard() {
      setHeader('Dashboard ผู้ดูแลระบบ', 'ภาพรวมระบบรับสมัครและการดำเนินงาน');
      const content = document.getElementById('content');
      content.innerHTML = `<div class="dashboard-skeleton"><div class="spinner-border text-primary"></div><div>กำลังโหลด Dashboard</div></div>`;

      const data = await serverCall('getStaffDashboard', state.token);
      const cards = data.cards || {};
      const analytics = data.analytics || {};
      const quality = data.quality || {};
      const stats = Array.isArray(data.departmentStats) ? data.departmentStats : [];
      const total = Number(cards.totalApplicants || 0);
      const growth = analytics.growth7d;
      const pending = Number(quality.pendingQualification || 0);
      const top = analytics.topDepartment || stats[0] || null;
      const top3Share = Number(analytics.top3Share || 0);
      const growthMeta = growth === null
        ? `${Number(cards.last7Days || 0).toLocaleString('th-TH')} คนใน 7 วันล่าสุด`
        : growth > 0
          ? `<i class="fa-solid fa-arrow-trend-up me-1"></i>เพิ่ม ${Math.abs(growth).toLocaleString('th-TH')}% จากช่วงก่อน`
          : growth < 0
            ? `<i class="fa-solid fa-arrow-trend-down me-1"></i>ลด ${Math.abs(growth).toLocaleString('th-TH')}% จากช่วงก่อน`
            : 'ทรงตัวเมื่อเทียบ 7 วันก่อน';
      const growthClass = growth > 0 ? 'intel-positive' : growth < 0 ? 'intel-negative' : 'intel-neutral';
      const pendingRate = total ? pending / total * 100 : 0;

      const insightRows = [
        top ? `<strong>${escapeHtml(top.label || top.department || '-')}</strong> มีผู้สมัครสูงสุด ${Number(top.count || 0).toLocaleString('th-TH')} คน (${Number(top.share || 0).toFixed(1)}%)` : 'ยังไม่มีข้อมูลเพียงพอสำหรับจัดอันดับหน่วยงาน',
        pending ? `ยังรอตรวจคุณสมบัติ <strong>${pending.toLocaleString('th-TH')} คน</strong> หรือ ${pendingRate.toFixed(1)}% ของผู้สมัครทั้งหมด` : '<strong>ไม่มีรายการค้างตรวจคุณสมบัติ</strong>',
        total ? `ผู้สมัครใน 3 หน่วยงานแรกคิดเป็น <strong>${top3Share.toFixed(1)}%</strong> ของทั้งระบบ` : 'ยังไม่มีข้อมูลสำหรับวิเคราะห์การกระจายตัว'
      ];

      content.innerHTML = `
        <div class="d-flex flex-wrap justify-content-between gap-3 align-items-end mb-4">
          <div>
            <div class="page-title">Dashboard ผู้ดูแลระบบ</div>
            <div class="page-subtitle">ติดตามจำนวนผู้สมัคร การตรวจสอบคุณสมบัติ และการส่งรายชื่อให้หน่วยงาน</div>
          </div>
          <button class="btn btn-outline-secondary btn-sm" onclick="renderAdminDashboard()"><i class="fa-solid fa-rotate me-1"></i>อัปเดตข้อมูล</button>
        </div>

        <div class="row g-3 mb-3">
          ${dashboardMetric('fa-users', cards.totalApplicants, 'ผู้สมัครทั้งหมด', `วันนี้ +${Number(cards.todayApplications || 0).toLocaleString('th-TH')} คน`, '', 'yellow')}
          ${dashboardMetric('fa-calendar-week', cards.last7Days, 'ผู้สมัคร 7 วันล่าสุด', growthMeta, growthClass, 'blue')}
          ${dashboardMetric('fa-user-check', quality.reviewed, 'ตรวจคุณสมบัติแล้ว', `ค้าง ${pending.toLocaleString('th-TH')} คน`, pending ? 'intel-negative' : 'intel-positive', 'green')}
          ${dashboardMetric('fa-paper-plane', quality.forwarded, 'ส่งให้หน่วยงานแล้ว', `${Number(analytics.forwardedRate || 0).toFixed(1)}% ของผู้สมัครทั้งหมด`, '', 'purple')}
        </div>

        <div class="row g-3 mb-4">
          ${dashboardMetric('fa-person', cards.maleApplicants, 'ผู้สมัครเพศชาย', `${Number(cards.maleShare || 0).toFixed(1)}% ของผู้สมัครทั้งหมด`, '', 'blue')}
          ${dashboardMetric('fa-person-dress', cards.femaleApplicants, 'ผู้สมัครเพศหญิง', `${Number(cards.femaleShare || 0).toFixed(1)}% ของผู้สมัครทั้งหมด`, '', 'red')}
          ${dashboardMetric('fa-user-plus', cards.newApplicants, 'ผู้สมัครรายใหม่', `${Number(cards.newShare || 0).toFixed(1)}% ของผู้สมัครทั้งหมด`, '', 'green')}
          ${dashboardMetric('fa-clock-rotate-left', cards.returningApplicants, 'ผู้สมัครรายเก่า', `${Number(cards.returningShare || 0).toFixed(1)}% ของผู้สมัครทั้งหมด`, '', 'orange')}
        </div>

        <div class="row g-3 mb-4">
          <div class="col-xl-8">
            <div class="intel-panel chart-panel h-100">
              <div class="intel-panel-head">
                <div><h2 class="intel-panel-title"><i class="fa-solid fa-chart-line me-2"></i>แนวโน้มการสมัคร 14 วัน</h2><div class="intel-panel-subtitle">จำนวนผู้สมัครรายวันจากข้อมูลจริงในระบบ</div></div>
                <span class="analysis-badge">สูงสุด ${Number(analytics.busiestDay?.count || 0).toLocaleString('th-TH')} คน · ${escapeHtml(analytics.busiestDay?.label || '-')}</span>
              </div>
              <div class="trend-wrap">${buildAdminTrendChart(data.trend14Days)}</div>
            </div>
          </div>
          <div class="col-xl-4">
            <div class="intel-panel h-100 watch-panel">
              <div class="intel-panel-head"><div><h2 class="intel-panel-title"><i class="fa-solid fa-lightbulb me-2"></i>ข้อมูลที่ควรติดตาม</h2><div class="intel-panel-subtitle">สรุปประเด็นสำคัญอัตโนมัติ</div></div></div>
              <div class="p-3 pt-1">${insightRows.map((text,index)=>`<div class="watch-row"><span class="watch-icon"><i class="fa-solid ${index===0?'fa-ranking-star':index===1?'fa-hourglass-half':'fa-chart-pie'}"></i></span><div>${text}</div></div>`).join('')}</div>
            </div>
          </div>
        </div>

        <div class="intel-panel department-stat-panel">
          <div class="intel-panel-head">
            <div><h2 class="intel-panel-title"><i class="fa-solid fa-building me-2"></i>สถิติยอดการสมัครแยกตามหน่วยงาน</h2><div class="intel-panel-subtitle">แสดงตามหน่วยคัดเลือกจริง เช่น กลุ่มงานพัฒนานักศึกษา › งานพัฒนานักศึกษา / งานกีฬาและนันทนาการ และกลุ่มงานศิษย์เก่าสัมพันธ์ › งานศิษย์เก่าสัมพันธ์ / งานอำนวยการ</div></div>
            <span class="analysis-badge">${stats.length.toLocaleString('th-TH')} หน่วยคัดเลือก</span>
          </div>
          <div class="dept-intel-list">${buildDepartmentIntelligence(stats)}</div>
        </div>`;
    }

    /* ======================================================
       ADMIN APPLICANT REVIEW
    ====================================================== */
    let applicantReviewSearchTimer = 0;

    async function renderApplicantReview() {
      setHeader('ตรวจสอบข้อมูลผู้สมัคร', 'โหลดเฉพาะข้อมูลหน้าที่กำลังดูเพื่อลดเวลาและปริมาณข้อมูล');
      state.applicantReviewPage = 1;
      const content = document.getElementById('content');
      content.innerHTML = `
        <div class="d-flex flex-wrap justify-content-between gap-3 align-items-end mb-4">
          <div><div class="page-title">ตรวจสอบข้อมูลผู้สมัคร</div><div class="page-subtitle">กรองตามหน่วยงาน คณะ และสถานะ โดยโหลดเฉพาะข้อมูลหน้าปัจจุบัน</div></div>
        </div>
        <div id="applicantReviewStats" class="row g-3 mb-4"></div>
        <div class="panel premium-panel">
          <div class="panel-header filter-toolbar">
            <div><h2 class="panel-title">รายชื่อผู้สมัครทั้งหมด</h2><div id="applicantReviewResultText" class="text-muted small mt-1">กำลังโหลดข้อมูล...</div></div>
            <div class="filter-grid">
              <select id="applicantReviewDepartmentFilter" class="form-select form-select-sm" onchange="filterApplicantReviewTable()"><option value="">ทุกหน่วยงาน</option></select>
              <select id="applicantReviewFacultyFilter" class="form-select form-select-sm" onchange="filterApplicantReviewTable()"><option value="">ทุกคณะ</option></select>
              <select id="applicantReviewStatusFilter" class="form-select form-select-sm" onchange="filterApplicantReviewTable()"><option value="">ทุกสถานะ</option></select>
              <select id="applicantReviewPageSize" class="form-select form-select-sm" onchange="changeApplicantReviewPageSize(this.value)"><option value="20">20 แถว</option><option value="50">50 แถว</option><option value="100">100 แถว</option></select>
              <div class="search-box"><i class="fa-solid fa-magnifying-glass"></i><input id="applicantReviewSearch" class="form-control form-control-sm" placeholder="ค้นหารหัส / ชื่อ / คณะ / งาน" oninput="queueApplicantReviewSearch()"></div>
            </div>
          </div>
          <div class="table-wrap"><table class="table table-hover align-middle"><thead><tr><th>ลำดับ</th><th>วันที่สมัคร</th><th>รหัสนักศึกษา</th><th>ชื่อ-สกุล</th><th>คณะ / สาขา</th><th>ชั้นปี</th><th>GPAX</th><th>งานที่สมัคร</th><th>สถานะ</th><th class="text-center">จัดการ</th></tr></thead><tbody id="applicantReviewBody"><tr><td colspan="10" class="text-center py-5"><div class="spinner-border spinner-border-sm text-primary me-2"></div>กำลังโหลดรายชื่อ</td></tr></tbody></table></div>
          <div id="applicantReviewPager" class="pager-shell"></div>
        </div>`;
      await loadApplicantReviewPage(true);
    }

    function watermarkStatCard(icon, number, label, theme='yellow') {
      return `<div class="col-12 col-sm-6 col-xl-3"><div class="stat-card stat-card-watermark wm-card wm-${escapeHtml(theme)}"><div class="wm-bg-icon"><i class="fa-solid ${icon}"></i></div><div class="position-relative"><div class="stat-icon mb-3"><i class="fa-solid ${icon}"></i></div><div class="stat-number">${Number(number||0).toLocaleString('th-TH')}</div><div class="stat-label">${escapeHtml(label)}</div></div></div></div>`;
    }

    function getApplicantReviewOptions(includeFilters=false) {
      return {
        mode: 'review', page: state.applicantReviewPage, pageSize: state.applicantReviewPageSize, includeFilters: !!includeFilters,
        search: (document.getElementById('applicantReviewSearch')?.value || '').trim(),
        department: document.getElementById('applicantReviewDepartmentFilter')?.value || '',
        faculty: document.getElementById('applicantReviewFacultyFilter')?.value || '',
        status: document.getElementById('applicantReviewStatusFilter')?.value || ''
      };
    }

    async function loadApplicantReviewPage(firstLoad=false) {
      const seq = ++state.applicantReviewRequestSeq;
      const body = document.getElementById('applicantReviewBody');
      if (body) body.innerHTML = `<tr><td colspan="10" class="text-center py-5"><div class="spinner-border spinner-border-sm text-primary me-2"></div>กำลังโหลดข้อมูล</td></tr>`;
      try {
        const result = await serverCall('getAdminApplicantPage', state.token, getApplicantReviewOptions(firstLoad));
        if (seq !== state.applicantReviewRequestSeq || state.currentView !== 'applicantReview') return;
        state.applicantReviewPage = Number(result.page || 1);
        state.applicantReviewPageSize = Number(result.pageSize || 20);
        state.adminApplicantRows = Array.isArray(result.rows) ? result.rows : [];
        const summary = result.summary || {};
        document.getElementById('applicantReviewStats').innerHTML = `
          ${watermarkStatCard('fa-users', summary.total, 'ผู้สมัครทั้งหมด', 'yellow')}
          ${watermarkStatCard('fa-clipboard-check', summary.reviewed, 'ตรวจคุณสมบัติแล้ว', 'green')}
          ${watermarkStatCard('fa-hourglass-half', summary.pendingReview, 'ยังไม่ตรวจสอบ', 'orange')}
          ${watermarkStatCard('fa-paper-plane', summary.forwarded, 'ส่งให้หน่วยงานแล้ว', 'purple')}`;
        if (firstLoad) populateApplicantReviewFilters(result.filters || {});
        renderApplicantReviewRows(state.adminApplicantRows, (state.applicantReviewPage - 1) * state.applicantReviewPageSize);
        renderSimplePager('applicantReviewPager', Number(result.totalFiltered || 0), state.applicantReviewPage, state.applicantReviewPageSize, 'goApplicantReviewPage');
        const text = document.getElementById('applicantReviewResultText');
        if (text) text.textContent = `พบ ${Number(result.totalFiltered || 0).toLocaleString('th-TH')} รายการ · โหลดเฉพาะหน้าปัจจุบัน`;
      } catch (error) {
        if (seq !== state.applicantReviewRequestSeq) return;
        if (body) body.innerHTML = `<tr><td colspan="10" class="text-center text-danger py-5">${escapeHtml(error.message)}</td></tr>`;
        throw error;
      }
    }

    function populateApplicantReviewFilters(filters) {
      const fill = (id, values, label) => {
        const el=document.getElementById(id); if(!el) return; const selected=el.value;
        el.innerHTML=`<option value="">${label}</option>${(Array.isArray(values)?values:[]).map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('')}`;
        el.value=selected;
      };
      fill('applicantReviewDepartmentFilter', filters.departments, 'ทุกหน่วยงาน');
      fill('applicantReviewFacultyFilter', filters.faculties, 'ทุกคณะ');
      fill('applicantReviewStatusFilter', filters.statuses, 'ทุกสถานะ');
      const size=document.getElementById('applicantReviewPageSize'); if(size) size.value=String(state.applicantReviewPageSize);
    }

    function queueApplicantReviewSearch(){ clearTimeout(applicantReviewSearchTimer); applicantReviewSearchTimer=setTimeout(()=>filterApplicantReviewTable(),650); }
    function filterApplicantReviewTable(){ state.applicantReviewPage=1; loadApplicantReviewPage(false).catch(error=>Swal.fire({icon:'error',title:'โหลดข้อมูลไม่สำเร็จ',text:error.message})); }
    function changeApplicantReviewPageSize(v){ state.applicantReviewPageSize=[20,50,100].includes(Number(v))?Number(v):20; state.applicantReviewPage=1; loadApplicantReviewPage(false).catch(error=>Swal.fire({icon:'error',title:'โหลดข้อมูลไม่สำเร็จ',text:error.message})); }
    function goApplicantReviewPage(p){ state.applicantReviewPage=Number(p)||1; loadApplicantReviewPage(false).catch(error=>Swal.fire({icon:'error',title:'โหลดข้อมูลไม่สำเร็จ',text:error.message})); }

    function renderApplicantReviewRows(rows, offset=0) {
      const body=document.getElementById('applicantReviewBody'); if(!body) return;
      body.innerHTML=rows.length ? rows.map((item,index)=>`<tr><td>${offset+index+1}</td><td class="text-nowrap">${escapeHtml(item.submittedAt||'-')}</td><td class="candidate-code fw-medium">${escapeHtml(item.studentId||'-')}</td><td class="candidate-name">${escapeHtml(item.fullName||'-')}</td><td><div>${escapeHtml(item.faculty||'-')}</div>${item.major?`<div class="text-muted table-subtext">${escapeHtml(item.major)}</div>`:''}</td><td>${escapeHtml(item.year||'-')}</td><td class="fw-semibold">${escapeHtml(item.gpax||'-')}</td><td style="min-width:220px"><div>${escapeHtml(item.job||'-')}</div>${item.department?`<div class="text-muted table-subtext">${escapeHtml(item.selectionUnit||item.department)}</div>`:''}</td><td>${statusBadge(item.status)}</td><td class="text-center text-nowrap"><button class="btn btn-outline-primary btn-sm" onclick="showApplicantDetail('${escapeHtml(item.applicationId)}')"><i class="fa-solid fa-magnifying-glass me-1"></i>ตรวจสอบข้อมูล</button></td></tr>`).join('') : `<tr><td colspan="10" class="text-center text-muted py-5">ไม่พบข้อมูลตามตัวกรอง</td></tr>`;
    }

    function renderSimplePager(targetId,total,page,size,handler){
      const el=document.getElementById(targetId); if(!el) return; const pages=Math.max(1,Math.ceil(total/size));
      el.innerHTML=`<div class="pager-bar"><div class="pager-summary">แสดง ${total?((page-1)*size+1):0}-${Math.min(page*size,total)} จาก ${Number(total||0).toLocaleString('th-TH')} รายการ</div><div class="pager-actions"><button class="btn btn-outline-secondary btn-sm" ${page<=1?'disabled':''} onclick="${handler}(${page-1})"><i class="fa-solid fa-chevron-left"></i></button><span>หน้า <strong>${page}</strong> / ${pages}</span><button class="btn btn-outline-secondary btn-sm" ${page>=pages?'disabled':''} onclick="${handler}(${page+1})"><i class="fa-solid fa-chevron-right"></i></button></div></div>`;
    }


    /* ======================================================
       CIVIL REGISTRY / DOPA CHECK
    ====================================================== */
    let civilRegistrySearchTimer = 0;

    async function renderCivilRegistry() {
      setHeader('ตรวจสอบข้อมูลบุคคลกับฐานข้อมูลทะเบียนราษฎร กรมการปกครอง', 'แสดงข้อมูลแบบแบ่งหน้าเพื่อลดภาระการโหลด');
      state.civilRegistryPage = 1;
      const content = document.getElementById('content');
      content.innerHTML = `
        <div class="d-flex flex-wrap justify-content-between gap-3 align-items-end mb-3">
          <div><div class="page-title">ตรวจสอบข้อมูลบุคคลกับฐานข้อมูลทะเบียนราษฎร กรมการปกครอง</div><div class="page-subtitle">ตรวจความถูกต้องของเลขประจำตัวประชาชนและสถานะข้อมูลที่เกี่ยวข้อง</div></div>
          <div class="d-flex flex-wrap gap-2"><button class="btn btn-outline-primary" onclick="runCivilRegistryCheck('selected')"><i class="fa-solid fa-list-check me-1"></i>ตรวจสอบที่เลือก</button><button class="btn btn-primary" onclick="runCivilRegistryCheck('all')"><i class="fa-solid fa-users-gear me-1"></i>ตรวจสอบทั้งหมดตามตัวกรอง</button></div>
        </div>
        <div id="civilRegistryStats" class="row g-3 mb-4"></div>
        <div class="panel premium-panel">
          <div class="panel-header filter-toolbar"><div><h2 class="panel-title">รายการตรวจสอบเลขประจำตัวประชาชน</h2><div id="civilRegistryResultText" class="text-muted small mt-1">กำลังโหลดข้อมูล...</div></div>
            <div class="filter-grid civil-filter-grid">
              <select id="civilRegistryDepartmentFilter" class="form-select form-select-sm" onchange="filterCivilRegistryTable()"><option value="">ทุกหน่วยงาน</option></select>
              <select id="civilRegistryFacultyFilter" class="form-select form-select-sm" onchange="filterCivilRegistryTable()"><option value="">ทุกคณะ</option></select>
              <select id="civilRegistryFilter" class="form-select form-select-sm" onchange="filterCivilRegistryTable()"><option value="">ทุกสถานะ</option><option value="verified">ตรวจสอบแล้ว</option><option value="pending">ยังไม่ตรวจสอบ</option><option value="invalid">เลข 13 หลักไม่ถูกต้อง</option></select>
              <select id="civilRegistryPageSize" class="form-select form-select-sm" onchange="changeCivilRegistryPageSize(this.value)"><option value="20">20 แถว</option><option value="50">50 แถว</option><option value="100">100 แถว</option></select>
              <div class="search-box"><i class="fa-solid fa-magnifying-glass"></i><input id="civilRegistrySearch" class="form-control form-control-sm" placeholder="ค้นหารหัส / ชื่อ / คณะ" oninput="queueCivilRegistrySearch()"></div>
            </div>
          </div>
          <div class="table-wrap"><table class="table table-hover align-middle"><thead><tr><th class="text-center"><input id="civilRegistrySelectAll" class="form-check-input" type="checkbox" onchange="toggleAllCivilRegistryRows(this)"></th><th>ลำดับ</th><th>รหัสนักศึกษา</th><th>ชื่อ-สกุล</th><th>เลขประจำตัวประชาชน</th><th>ผลเลข 13 หลัก</th><th>สถานภาพบุคคล</th><th>สถานะภูมิลำเนา</th><th>สถานะการอยู่อาศัย</th><th>สัญชาติ</th><th>สถานะการตรวจสอบ</th><th class="text-center">จัดการ</th></tr></thead><tbody id="civilRegistryBody"><tr><td colspan="12" class="text-center py-5"><div class="spinner-border spinner-border-sm text-primary me-2"></div>กำลังโหลดข้อมูล</td></tr></tbody></table></div>
          <div id="civilRegistryPager" class="pager-shell"></div>
        </div>`;
      await loadCivilRegistryPage(true);
    }

    function getCivilRegistryOptions(includeFilters=false){ return { mode:'civil', page:state.civilRegistryPage, pageSize:state.civilRegistryPageSize, includeFilters:!!includeFilters, search:(document.getElementById('civilRegistrySearch')?.value||'').trim(), department:document.getElementById('civilRegistryDepartmentFilter')?.value||'', faculty:document.getElementById('civilRegistryFacultyFilter')?.value||'', registryStatus:document.getElementById('civilRegistryFilter')?.value||'' }; }

    async function loadCivilRegistryPage(firstLoad=false){
      const seq=++state.civilRegistryRequestSeq; const body=document.getElementById('civilRegistryBody');
      if(body) body.innerHTML=`<tr><td colspan="12" class="text-center py-5"><div class="spinner-border spinner-border-sm text-primary me-2"></div>กำลังโหลดข้อมูล</td></tr>`;
      const result=await serverCall('getAdminApplicantPage',state.token,getCivilRegistryOptions(firstLoad));
      if(seq!==state.civilRegistryRequestSeq||state.currentView!=='civilRegistry') return;
      state.civilRegistryPage=Number(result.page||1); state.civilRegistryPageSize=Number(result.pageSize||20); state.civilRegistryRows=Array.isArray(result.rows)?result.rows:[];
      const summary=result.summary||{};
      document.getElementById('civilRegistryStats').innerHTML=`${watermarkStatCard('fa-users',summary.total,'ผู้สมัครทั้งหมด','yellow')}${watermarkStatCard('fa-user-check',summary.registryVerified,'ตรวจสอบแล้ว','green')}${watermarkStatCard('fa-id-card-clip',summary.registryInvalid,'เลข 13 หลักไม่ถูกต้อง','red')}${watermarkStatCard('fa-clock',summary.registryPending,'ยังไม่ตรวจสอบ','blue')}`;
      if(firstLoad){
        const fill=(id,values,label)=>{const el=document.getElementById(id);if(!el)return;el.innerHTML=`<option value="">${label}</option>${(Array.isArray(values)?values:[]).map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('')}`;};
        fill('civilRegistryDepartmentFilter',result.filters?.departments,'ทุกหน่วยงาน'); fill('civilRegistryFacultyFilter',result.filters?.faculties,'ทุกคณะ'); const size=document.getElementById('civilRegistryPageSize');if(size)size.value=String(state.civilRegistryPageSize);
      }
      renderCivilRegistryRows(state.civilRegistryRows,(state.civilRegistryPage-1)*state.civilRegistryPageSize); renderSimplePager('civilRegistryPager',Number(result.totalFiltered||0),state.civilRegistryPage,state.civilRegistryPageSize,'goCivilRegistryPage');
      const text=document.getElementById('civilRegistryResultText');if(text)text.textContent=`พบ ${Number(result.totalFiltered||0).toLocaleString('th-TH')} รายการ · โหลดเฉพาะหน้าปัจจุบัน`;
      const box=document.getElementById('civilRegistrySelectAll');if(box)box.checked=false;
    }
    function queueCivilRegistrySearch(){clearTimeout(civilRegistrySearchTimer);civilRegistrySearchTimer=setTimeout(()=>filterCivilRegistryTable(),650);}
    function filterCivilRegistryTable(){state.civilRegistryPage=1;loadCivilRegistryPage(false).catch(error=>Swal.fire({icon:'error',title:'โหลดข้อมูลไม่สำเร็จ',text:error.message}));}
    function changeCivilRegistryPageSize(v){state.civilRegistryPageSize=[20,50,100].includes(Number(v))?Number(v):20;state.civilRegistryPage=1;loadCivilRegistryPage(false).catch(error=>Swal.fire({icon:'error',title:'โหลดข้อมูลไม่สำเร็จ',text:error.message}));}
    function goCivilRegistryPage(p){state.civilRegistryPage=Number(p)||1;loadCivilRegistryPage(false).catch(error=>Swal.fire({icon:'error',title:'โหลดข้อมูลไม่สำเร็จ',text:error.message}));}
    function renderCivilRegistryRows(rows,offset=0){const body=document.getElementById('civilRegistryBody');if(!body)return;body.innerHTML=rows.length?rows.map((item,index)=>`<tr><td class="text-center"><input class="form-check-input civil-row-check" type="checkbox" value="${escapeHtml(item.applicationId)}"></td><td>${offset+index+1}</td><td class="candidate-code fw-medium">${escapeHtml(item.studentId||'-')}</td><td><div class="candidate-name">${escapeHtml(item.fullName||'-')}</div><div class="text-muted table-subtext">${escapeHtml(item.faculty||'-')}</div></td><td class="text-nowrap">${escapeHtml(item.idCardMasked||'-')}</td><td>${civilChecksumBadge(item.checksumStatus)}</td><td>${civilRegistryValueBadge(item.personStatus)}</td><td>${civilRegistryValueBadge(item.domicileStatus)}</td><td>${civilRegistryValueBadge(item.residenceStatus)}</td><td>${civilRegistryValueBadge(item.nationality)}</td><td>${civilVerificationBadge(item.verificationStatus)}${item.verifiedAt?`<div class="text-muted mt-1 table-subtext">${escapeHtml(item.verifiedAt)}</div>`:''}</td><td class="text-center text-nowrap"><button class="btn btn-outline-primary btn-sm" onclick="runCivilRegistrySingle('${escapeHtml(item.applicationId)}')"><i class="fa-solid fa-shield-halved me-1"></i>ตรวจสอบ</button></td></tr>`).join(''):`<tr><td colspan="12" class="text-center text-muted py-5">ไม่พบข้อมูลตามตัวกรอง</td></tr>`;}
    function civilChecksumBadge(status){const value=String(status||'ยังไม่ตรวจสอบ');if(value==='ถูกต้อง')return `<span class="status-badge badge-pass"><i class="fa-solid fa-check me-1"></i>ถูกต้อง</span>`;if(value==='ไม่ถูกต้อง')return `<span class="status-badge badge-fail"><i class="fa-solid fa-xmark me-1"></i>ไม่ถูกต้อง</span>`;return `<span class="status-badge badge-wait">ยังไม่ตรวจสอบ</span>`;}
    function civilRegistryValueBadge(value){const text=String(value||'').trim();return text?`<span class="status-badge badge-info">${escapeHtml(text)}</span>`:`<span class="text-muted">-</span>`;}
    function civilVerificationBadge(status){const value=String(status||'ยังไม่ตรวจสอบ');if(value==='ตรวจสอบเลขประจำตัวแล้ว')return `<span class="status-badge badge-pass">ตรวจสอบแล้ว</span>`;if(value.includes('ตรวจสอบไม่ได้')||value.includes('ไม่สำเร็จ'))return `<span class="status-badge badge-fail">${escapeHtml(value)}</span>`;return `<span class="status-badge badge-wait">${escapeHtml(value)}</span>`;}
    function toggleAllCivilRegistryRows(source){document.querySelectorAll('#civilRegistryBody .civil-row-check').forEach(box=>box.checked=!!source.checked);}


    async function runCivilRegistrySingle(applicationId) {
      await runCivilRegistryIds([applicationId]);
    }

    async function runCivilRegistryCheck(mode) {
      let ids = [];
      if (mode === 'all') {
        ids = await serverCall('getAdminCivilRegistryIds', state.token, getCivilRegistryOptions());
      } else {
        ids = [...document.querySelectorAll('.civil-row-check:checked')].map(box => box.value).filter(Boolean);
      }
      if (!ids.length) {
        Swal.fire({ icon: 'info', title: 'ยังไม่ได้เลือกรายชื่อ', text: 'กรุณาเลือกรายชื่อที่ต้องการตรวจสอบ' });
        return;
      }

      const confirm = await Swal.fire({
        icon: 'question',
        title: mode === 'all' ? 'ตรวจสอบผู้สมัครทั้งหมด' : `ตรวจสอบ ${ids.length} รายการ`,
        html: `<div class="text-start" style="font-size:14px;">ระบบจะตรวจความถูกต้องของเลขประจำตัวประชาชน 13 หลัก</div>`,
        showCancelButton: true,
        confirmButtonText: 'เริ่มตรวจสอบ',
        cancelButtonText: 'ยกเลิก',
        confirmButtonColor: '#165f9c'
      });
      if (!confirm.isConfirmed) return;
      await runCivilRegistryIds(ids);
    }

    async function runCivilRegistryIds(ids) {
      showLoading(`กำลังตรวจสอบ ${ids.length} รายการ`);
      try {
        const totals = { processed:0, validId:0, invalidId:0, autoFilled:0 };
        const allErrors = [];
        const chunkSize = 50;
        for (let i = 0; i < ids.length; i += chunkSize) {
          const batchIds = ids.slice(i, i + chunkSize);
          const updates = batchIds.map(applicationId => ({ applicationId, idValidationOnly: true }));
          const result = await serverCall('saveQualificationReviews', state.token, updates);
          if (!result || typeof result.processed === 'undefined') {
            throw new Error('ยังไม่สามารถบันทึกผลตรวจเลข 13 หลักได้ กรุณา Deploy Code.gs ตัวล่าสุดอีกครั้ง แล้วรีเฟรชหน้าเว็บ');
          }
          Object.keys(totals).forEach(key => totals[key] += Number(result && result[key] || 0));
          if (Array.isArray(result && result.errors)) allErrors.push(...result.errors);
        }
        closeLoading();
        await Swal.fire({
          icon: totals.invalidId ? 'warning' : 'success',
          title: 'ตรวจสอบเรียบร้อย',
          html: `
            <div class="text-start mx-auto" style="max-width:460px;font-size:14px;">
              <div>ประมวลผล <strong>${totals.processed}</strong> รายการ</div>
              <div>เลข 13 หลักถูกต้อง <strong>${totals.validId}</strong> รายการ</div>
              <div>เลข 13 หลักไม่ถูกต้อง <strong>${totals.invalidId}</strong> รายการ</div>
              <div>เติมค่ามาตรฐานอัตโนมัติ <strong>${totals.autoFilled}</strong> รายการ</div>
            </div>`
        });
        await loadCivilRegistryPage(false);
      } catch (error) {
        closeLoading();
        Swal.fire({ icon: 'error', title: 'ตรวจสอบไม่สำเร็จ', text: error.message });
      }
    }

    /* ======================================================
       DEPARTMENT DASHBOARD
    ====================================================== */
    async function renderDepartmentDashboard() {
      setHeader('ภาพรวมหน่วยงาน', state.selectionUnit || state.department);
      const content = document.getElementById('content');
      content.innerHTML = `<div class="text-center py-5"><div class="spinner-border text-primary"></div></div>`;

      // Dashboard หน่วยงานโหลดเฉพาะ summary ไม่ดึงรายชื่อทั้งหมด
      const data = await serverCall('getStaffDashboard', state.token) || {};
      const cards = data.cards || {};
      const received = Number(cards.received || 0);
      const pending = Number(cards.pending || 0);
      const basket = Number(cards.basket || 0);
      const selected = Number(cards.selected || cards.sentForAnnouncement || 0);
      const notSelected = Number(cards.notSelected || 0);

      let actionIcon = 'fa-circle-check';
      let actionTitle = 'ไม่มีรายการที่ต้องดำเนินการ';
      let actionText = received ? 'หน่วยงานดำเนินการกับรายชื่อที่ได้รับครบแล้ว' : 'ยังไม่มีรายชื่อที่ผู้ดูแลระบบส่งมาให้พิจารณา';

      if (pending > 0) {
        actionIcon = 'fa-user-clock';
        actionTitle = `มี ${pending.toLocaleString('th-TH')} รายชื่อรอพิจารณา`;
        actionText = 'ตรวจสอบข้อมูลผู้สมัครและเลือกรายชื่อที่ต้องการเข้าสู่ตะกร้า';
      } else if (basket > 0) {
        actionIcon = 'fa-basket-shopping';
        actionTitle = `มี ${basket.toLocaleString('th-TH')} รายชื่อในตะกร้า`;
        actionText = 'ตรวจสอบความถูกต้องของรายชื่อก่อนยืนยันส่งผลการคัดเลือก';
      }

      content.innerHTML = `
        <div class="d-flex flex-wrap justify-content-between gap-3 align-items-end mb-4">
          <div>
            <div class="page-title">Dashboard หน่วยงาน</div>
            <div class="page-subtitle">
              ${escapeHtml(state.department)}
              ${state.selectionUnit && state.selectionUnit !== state.department ? ` › ${escapeHtml(state.selectionUnit)}` : ''}
            </div>
          </div>
          <button class="btn btn-outline-secondary btn-sm" onclick="renderDepartmentDashboard()">
            <i class="fa-solid fa-rotate me-1"></i>อัปเดตข้อมูล
          </button>
        </div>

        <div class="row g-3 mb-4">
          ${watermarkStatCard('fa-inbox', received, 'รายชื่อที่ได้รับ', 'yellow')}
          ${watermarkStatCard('fa-user-clock', pending, 'รอพิจารณา', 'blue')}
          ${watermarkStatCard('fa-basket-shopping', basket, 'อยู่ในตะกร้า', 'orange')}
          ${watermarkStatCard('fa-circle-check', selected, 'ยืนยันคัดเลือกแล้ว', 'green')}
        </div>

        <div class="panel">
          <div class="panel-header">
            <div>
              <h2 class="panel-title"><i class="fa-solid ${actionIcon} me-2"></i>งานที่ต้องดำเนินการ</h2>
              <div class="text-muted mt-1" style="font-size:13px;">${escapeHtml(actionText)}</div>
            </div>
          </div>
          <div class="p-3 p-md-4 d-flex flex-wrap align-items-center justify-content-between gap-3">
            <div>
              <div class="fw-semibold" style="font-size:17px;color:var(--navy-dark);">${escapeHtml(actionTitle)}</div>
              ${notSelected > 0 ? `<div class="text-muted mt-1" style="font-size:13px;">ไม่คัดเลือกแล้ว ${notSelected.toLocaleString('th-TH')} รายชื่อ</div>` : ''}
            </div>
            <div class="d-flex flex-wrap gap-2">
              <button class="btn btn-primary" onclick="navigate('departmentSelection')">
                <i class="fa-solid fa-user-check me-1"></i>พิจารณารายชื่อ
              </button>
              <button class="btn btn-outline-primary" onclick="navigate('departmentBasket')">
                <i class="fa-solid fa-basket-shopping me-1"></i>เปิดตะกร้า
                ${basket > 0 ? `<span class="badge text-bg-warning ms-1">${basket}</span>` : ''}
              </button>
            </div>
          </div>
        </div>
      `;

      updateBasketIndicator();
    }

    /* ======================================================
       QUALIFICATION
    ====================================================== */
    const QUALIFICATION_THRESHOLD = 2.00;
    const QUALIFICATION_FAIL_REASON = 'เกรดเฉลี่ยสะสมไม่เป็นไปตามหลักเกณฑ์ที่ประกาศรับสมัคร';
    let qualificationSearchTimer = null;

    async function renderQualification() {
      setHeader('ตรวจสอบคุณสมบัติเกรดเฉลี่ย', 'โหลดรายชื่อแบบแบ่งหน้า เพื่อลดข้อมูลและ DOM ที่ต้องประมวลผลพร้อมกัน');
      const content = document.getElementById('content');
      content.innerHTML = `
        <div class="d-flex flex-wrap justify-content-between gap-3 align-items-end mb-3">
          <div>
            <div class="page-title">ตรวจสอบคุณสมบัติเกรดเฉลี่ย</div>
            <div class="page-subtitle">เกณฑ์ผ่าน GPAX ตั้งแต่ ${QUALIFICATION_THRESHOLD.toFixed(2)} ขึ้นไป · แสดงทีละหน้าเพื่อให้ระบบลื่นขึ้น</div>
          </div>
          <div class="d-flex flex-wrap gap-2">
            <button class="btn btn-outline-success" onclick="exportQualificationCheckFile()">
              <i class="fa-solid fa-file-arrow-down me-1"></i>ดาวน์โหลดไฟล์ตรวจสอบ (.xls)
            </button>
            <button class="btn btn-outline-primary" onclick="document.getElementById('qualificationResultFile').click()">
              <i class="fa-solid fa-file-arrow-up me-1"></i>นำเข้าผลตรวจ
            </button>
            <input id="qualificationResultFile" type="file" class="d-none" accept=".xls,.xlsx,.csv,.html,.htm" onchange="importQualificationResultFile(this)">
            <button class="btn btn-primary" onclick="saveQualification()">
              <i class="fa-solid fa-floppy-disk me-1"></i>บันทึกการปรับแก้
            </button>
          </div>
        </div>

        <div class="result-summary mb-3">
          <div class="result-summary-item"><div id="qualificationSummaryTotal" class="result-summary-value">0</div><div class="result-summary-label">ทั้งหมด</div></div>
          <div class="result-summary-item"><div id="qualificationSummaryPass" class="result-summary-value result-pass">0</div><div class="result-summary-label">ผ่าน</div></div>
          <div class="result-summary-item"><div id="qualificationSummaryFail" class="result-summary-value result-fail">0</div><div class="result-summary-label">ไม่ผ่าน</div></div>
          <div class="result-summary-item"><div id="qualificationSummaryPending" class="result-summary-value result-wait">0</div><div class="result-summary-label">ยังไม่ตรวจสอบ</div></div>
        </div>

        <div class="panel mb-3">
          <div class="panel-header">
            <div>
              <h2 class="panel-title mb-1">ปรับสถานะแบบกลุ่ม</h2>
              <div class="text-muted small">เลือกผู้สมัครในหน้าปัจจุบัน แล้วกำหนดผลตรวจพร้อมกัน</div>
            </div>
            <div class="d-flex flex-wrap gap-2 align-items-center">
              <select id="qualificationBulkStatus" class="form-select form-select-sm" style="width:190px;">
                <option value="">เลือกสถานะ</option>
                <option value="pass">ผ่านคุณสมบัติ</option>
                <option value="fail">ไม่ผ่านคุณสมบัติ</option>
                <option value="pending">ยังไม่ตรวจสอบ</option>
              </select>
              <button class="btn btn-outline-primary btn-sm" onclick="applyQualificationBulkStatus()">
                <i class="fa-solid fa-layer-group me-1"></i>ปรับสถานะที่เลือก
              </button>
            </div>
          </div>
        </div>

        <div class="panel">
          <div class="panel-header">
            <div>
              <h2 class="panel-title">รายชื่อผู้สมัคร</h2>
              <div id="qualificationFilteredMeta" class="small text-muted mt-1"></div>
            </div>
            <div class="d-flex flex-wrap gap-2 align-items-center">
              <select id="qualificationStatusFilter" class="form-select form-select-sm" style="width:175px;" onchange="filterQualificationTable()">
                <option value="all">ทุกสถานะ</option>
                <option value="pass">ผ่าน</option>
                <option value="fail">ไม่ผ่าน</option>
                <option value="pending">ยังไม่ตรวจสอบ</option>
              </select>
              <input id="qualificationSearch" class="form-control form-control-sm" style="width:min(300px, 70vw);" placeholder="ค้นหารหัสนักศึกษา / ชื่อ / คณะ" oninput="filterQualificationTable()">
              <select id="qualificationPageSize" class="form-select form-select-sm" style="width:110px;" onchange="changeQualificationPageSize(this)">
                <option value="20">20 / หน้า</option>
                <option value="50">50 / หน้า</option>
                <option value="100">100 / หน้า</option>
              </select>
            </div>
          </div>
          <div id="qualificationTableLoading" class="d-none text-center py-4"><div class="spinner-border spinner-border-sm text-primary me-2"></div>กำลังโหลดรายชื่อ</div>
          <div class="table-wrap">
            <table class="table table-hover">
              <thead>
                <tr>
                  <th class="text-center"><input id="qualificationSelectAll" class="form-check-input" type="checkbox" onchange="toggleQualificationSelectAll(this)"></th>
                  <th>ลำดับ</th>
                  <th>รหัสนักศึกษา</th>
                  <th>ชื่อ-สกุล</th>
                  <th>คณะ</th>
                  <th>ชั้นปี</th>
                  <th>GPAX ที่กรอก</th>
                  <th>GPAX ที่ตรวจพบ</th>
                  <th style="min-width:165px;">ผลตรวจ</th>
                  <th style="min-width:360px;">สาเหตุที่ไม่ผ่าน</th>
                  <th class="text-center">แสดงผลให้นักศึกษา</th>
                  <th>สถานะระบบ</th>
                </tr>
              </thead>
              <tbody id="qualificationBody"><tr><td colspan="12" class="text-center text-muted py-4">กำลังโหลดข้อมูล...</td></tr></tbody>
            </table>
          </div>
          <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 p-3 border-top">
            <div id="qualificationPageInfo" class="small text-muted"></div>
            <div class="btn-group btn-group-sm">
              <button id="qualificationPrevBtn" class="btn btn-outline-secondary" onclick="qualificationPreviousPage()">ก่อนหน้า</button>
              <button id="qualificationNextBtn" class="btn btn-outline-secondary" onclick="qualificationNextPage()">ถัดไป</button>
            </div>
          </div>
        </div>
      `;

      document.getElementById('qualificationSearch').value = state.qualificationSearch || '';
      document.getElementById('qualificationStatusFilter').value = state.qualificationStatus || 'all';
      document.getElementById('qualificationPageSize').value = String(state.qualificationPageSize || 50);
      await loadQualificationPage(state.qualificationPage || 1);
    }

    function getQualificationPageOptions(page) {
      return {
        page: Math.max(1, Number(page || state.qualificationPage || 1)),
        pageSize: Number(state.qualificationPageSize || 50),
        search: state.qualificationSearch || '',
        status: state.qualificationStatus || 'all'
      };
    }

    async function loadQualificationPage(page = 1) {
      const seq = ++state.qualificationRequestSeq;
      const loading = document.getElementById('qualificationTableLoading');
      if (loading) loading.classList.remove('d-none');
      const selectAll = document.getElementById('qualificationSelectAll');
      if (selectAll) selectAll.checked = false;

      try {
        const data = await serverCall('getAdminQualificationPage', state.token, getQualificationPageOptions(page));
        if (seq !== state.qualificationRequestSeq || state.currentView !== 'qualification') return;

        state.qualificationRows = Array.isArray(data && data.rows) ? data.rows : [];
        state.qualificationPage = Number(data && data.page || 1);
        state.qualificationPageSize = Number(data && data.pageSize || state.qualificationPageSize || 50);
        state.qualificationPageCount = Number(data && data.pageCount || 1);

        const summary = data && data.summary || {};
        const setNumber = (id, value) => {
          const el = document.getElementById(id);
          if (el) el.textContent = Number(value || 0).toLocaleString('th-TH');
        };
        setNumber('qualificationSummaryTotal', summary.total);
        setNumber('qualificationSummaryPass', summary.passed);
        setNumber('qualificationSummaryFail', summary.failed);
        setNumber('qualificationSummaryPending', summary.pending);

        const meta = document.getElementById('qualificationFilteredMeta');
        if (meta) meta.textContent = `พบ ${Number(data.totalFiltered || 0).toLocaleString('th-TH')} รายการ`;
        const pageInfo = document.getElementById('qualificationPageInfo');
        if (pageInfo) pageInfo.textContent = `หน้า ${state.qualificationPage.toLocaleString('th-TH')} / ${state.qualificationPageCount.toLocaleString('th-TH')}`;
        const prev = document.getElementById('qualificationPrevBtn');
        const next = document.getElementById('qualificationNextBtn');
        if (prev) prev.disabled = state.qualificationPage <= 1;
        if (next) next.disabled = state.qualificationPage >= state.qualificationPageCount;

        renderQualificationRows(state.qualificationRows, Number(data && data.offset || 0));
      } finally {
        if (seq === state.qualificationRequestSeq && loading) loading.classList.add('d-none');
      }
    }

    function qualificationResultLabel(result) {
      if (result === 'pass') return 'ผ่านคุณสมบัติ';
      if (result === 'fail') return 'ไม่ผ่านคุณสมบัติ';
      return 'ยังไม่ตรวจสอบ';
    }

    function renderQualificationRows(rows, offset = 0) {
      const body = document.getElementById('qualificationBody');
      if (!body) return;
      body.innerHTML = rows.length ? rows.map((item, index) => {
        const locked = !!item.forwarded;
        const draft = state.qualificationDrafts.get(item.applicationId) || null;
        const result = draft
          ? draft.result
          : (['pass', 'fail'].includes(item.qualificationResult) ? item.qualificationResult : 'pending');
        const reasonValue = draft ? draft.reason : (item.qualificationReason || '');
        const visibleValue = draft ? !!draft.visible : !!item.qualificationVisible;
        const verified = item.verifiedGpax !== '' && item.verifiedGpax !== null && item.verifiedGpax !== undefined
          ? Number(item.verifiedGpax).toFixed(2)
          : '-';
        return `
          <tr data-app-id="${escapeHtml(item.applicationId)}" data-qualification-status="${escapeHtml(result)}">
            <td class="text-center">
              <input class="form-check-input qual-select" type="checkbox" ${locked ? 'disabled' : ''}>
            </td>
            <td>${offset + index + 1}</td>
            <td class="candidate-code">${escapeHtml(item.studentId)}</td>
            <td>${escapeHtml(item.fullName)}</td>
            <td>${escapeHtml(item.faculty)}</td>
            <td>${escapeHtml(item.year)}</td>
            <td class="fw-semibold">${escapeHtml(item.gpax || '-')}</td>
            <td class="fw-semibold ${result === 'fail' ? 'result-fail' : result === 'pass' ? 'result-pass' : ''}">${escapeHtml(verified)}</td>
            <td>
              <select class="form-select form-select-sm qual-result" ${locked ? 'disabled' : ''} onchange="onQualificationResultChange(this)">
                <option value="pending" ${result === 'pending' ? 'selected' : ''}>ยังไม่ตรวจสอบ</option>
                <option value="pass" ${result === 'pass' ? 'selected' : ''}>ผ่านคุณสมบัติ</option>
                <option value="fail" ${result === 'fail' ? 'selected' : ''}>ไม่ผ่านคุณสมบัติ</option>
              </select>
            </td>
            <td>
              <input class="form-control form-control-sm qual-reason" value="${escapeHtml(reasonValue)}" ${locked ? 'disabled' : ''} placeholder="ระบุสาเหตุกรณีไม่ผ่าน" oninput="rememberQualificationRow(this.closest('tr'))">
            </td>
            <td class="text-center">
              <div class="form-check form-switch d-inline-block">
                <input class="form-check-input qual-visible" type="checkbox" ${visibleValue ? 'checked' : ''} ${locked ? 'disabled' : ''} onchange="rememberQualificationRow(this.closest('tr'))">
              </div>
            </td>
            <td>${statusBadge(item.status)}</td>
          </tr>`;
      }).join('') : `<tr><td colspan="12" class="text-center text-muted py-4">ไม่พบรายชื่อ</td></tr>`;
    }

    function rememberQualificationRow(row) {
      if (!row || !row.dataset || !row.dataset.appId) return;
      const result = row.querySelector('.qual-result')?.value || 'pending';
      let reason = row.querySelector('.qual-reason')?.value.trim() || '';
      const visible = !!row.querySelector('.qual-visible')?.checked;
      if (result === 'fail' && !reason) reason = QUALIFICATION_FAIL_REASON;
      if (result !== 'fail') reason = '';
      state.qualificationDrafts.set(row.dataset.appId, {
        applicationId: row.dataset.appId,
        result,
        reason,
        visible
      });
      row.dataset.qualificationStatus = result;
    }

    function onQualificationResultChange(select) {
      const row = select.closest('tr');
      const reason = row.querySelector('.qual-reason');
      const visible = row.querySelector('.qual-visible');
      row.dataset.qualificationStatus = select.value;

      if (select.value === 'fail') {
        if (!reason.value.trim()) reason.value = QUALIFICATION_FAIL_REASON;
        visible.checked = true;
      } else {
        reason.value = '';
        visible.checked = select.value === 'pass';
      }
      rememberQualificationRow(row);
    }

    function toggleQualificationSelectAll(checkbox) {
      document.querySelectorAll('#qualificationBody tr[data-app-id]').forEach(row => {
        const item = row.querySelector('.qual-select');
        if (item && !item.disabled) item.checked = checkbox.checked;
      });
    }

    function applyQualificationBulkStatus() {
      const result = document.getElementById('qualificationBulkStatus').value;
      if (!result) {
        Swal.fire({ icon: 'info', title: 'กรุณาเลือกสถานะ', text: 'เลือก ผ่าน / ไม่ผ่าน / ยังไม่ตรวจสอบ ก่อนทำรายการ' });
        return;
      }

      const selectedRows = [...document.querySelectorAll('#qualificationBody tr[data-app-id]')]
        .filter(row => row.querySelector('.qual-select')?.checked && !row.querySelector('.qual-select')?.disabled);
      if (!selectedRows.length) {
        Swal.fire({ icon: 'info', title: 'ยังไม่ได้เลือกรายชื่อ', text: 'กรุณาเลือกรายชื่ออย่างน้อย 1 คนในหน้าปัจจุบัน' });
        return;
      }

      selectedRows.forEach(row => {
        const select = row.querySelector('.qual-result');
        select.value = result;
        onQualificationResultChange(select);
      });
      Swal.fire({ icon: 'success', title: 'ปรับสถานะในตารางแล้ว', text: `ปรับ ${selectedRows.length} รายการเป็น “${qualificationResultLabel(result)}” กรุณากดบันทึกการปรับแก้` });
    }

    function filterQualificationTable() {
      const search = (document.getElementById('qualificationSearch')?.value || '').trim();
      const status = document.getElementById('qualificationStatusFilter')?.value || 'all';
      state.qualificationSearch = search;
      state.qualificationStatus = status;
      state.qualificationPage = 1;
      clearTimeout(qualificationSearchTimer);
      qualificationSearchTimer = setTimeout(() => {
        loadQualificationPage(1).catch(error => {
          Swal.fire({ icon: 'error', title: 'ค้นหารายชื่อไม่สำเร็จ', text: error.message });
        });
      }, 650);
    }

    function goQualificationPage(page) {
      const target = Math.max(1, Math.min(Number(page || 1), Number(state.qualificationPageCount || 1)));
      if (target === state.qualificationPage) return;
      loadQualificationPage(target).catch(error => {
        Swal.fire({ icon: 'error', title: 'โหลดรายชื่อไม่สำเร็จ', text: error.message });
      });
    }

    function qualificationPreviousPage() {
      goQualificationPage(Number(state.qualificationPage || 1) - 1);
    }

    function qualificationNextPage() {
      goQualificationPage(Number(state.qualificationPage || 1) + 1);
    }

    function changeQualificationPageSize(select) {
      const size = Number(select && select.value || 50);
      state.qualificationPageSize = [20, 50, 100].includes(size) ? size : 50;
      state.qualificationPage = 1;
      loadQualificationPage(1).catch(error => {
        Swal.fire({ icon: 'error', title: 'โหลดรายชื่อไม่สำเร็จ', text: error.message });
      });
    }

    async function saveQualification() {
      const updates = Array.from(state.qualificationDrafts.values());
      if (!updates.length) {
        Swal.fire({ icon: 'info', title: 'ยังไม่มีรายการที่แก้ไข', text: 'ปรับผลตรวจหรือสถานะในตารางก่อนกดบันทึก' });
        return;
      }

      const confirm = await Swal.fire({
        icon: 'question',
        title: 'บันทึกผลการตรวจสอบ',
        text: `มีรายการที่แก้ไข ${updates.length.toLocaleString('th-TH')} รายการ`,
        showCancelButton: true,
        confirmButtonText: 'บันทึก',
        cancelButtonText: 'ยกเลิก',
        confirmButtonColor: '#165f9c'
      });
      if (!confirm.isConfirmed) return;

      showLoading('กำลังบันทึกผล');
      try {
        const result = await serverCall('saveQualificationReviews', state.token, updates);
        state.qualificationDrafts.clear();
        closeLoading();
        await Swal.fire({ icon: 'success', title: 'บันทึกเรียบร้อย', text: `อัปเดต ${result.updated || 0} รายการ` });
        await loadQualificationPage(state.qualificationPage || 1);
      } catch (error) {
        closeLoading();
        Swal.fire({ icon: 'error', title: 'บันทึกไม่สำเร็จ', text: error.message });
      }
    }

    async function exportQualificationCheckFile() {
      const header = ['ประทับเวลา', 'คำนำหน้า', 'ชื่อ (ไม่ต้องมีคำนำหน้า)', 'นามสกุล', 'เลขบัตรประจำตัวประชาชน', 'รหัสนักศึกษา', 'สังกัดคณะ'];
      showLoading('กำลังเตรียมไฟล์รายชื่อ');
      try {
        const allRows = await serverCall('getAdminQualificationApplicants', state.token) || [];
        const rows = allRows.map(item => [
          item.submittedAtExport || item.submittedAt || '',
          item.prefix || '',
          item.firstName || '',
          item.lastName || '',
          String(item.idCard || ''),
          String(item.studentId || ''),
          item.faculty || ''
        ]);

        try {
          await ensureXlsxLoaded();
        } catch (_) {
          closeLoading();
          downloadCsv('รายชื่อสำหรับตรวจสอบคุณสมบัติ.csv', [header, ...rows]);
          return;
        }

        const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
        ws['!cols'] = [
          { wch: 22 }, { wch: 15 }, { wch: 28 }, { wch: 28 }, { wch: 22 }, { wch: 18 }, { wch: 35 }
        ];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        const xlsBytes = XLSX.write(wb, { bookType: 'biff8', type: 'array' });
        const blob = new Blob([xlsBytes], { type: 'application/vnd.ms-excel' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'รายชื่อสำหรับตรวจสอบคุณสมบัติ.xls';
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        closeLoading();
      } catch (error) {
        closeLoading();
        Swal.fire({ icon: 'error', title: 'ดาวน์โหลดไฟล์ไม่สำเร็จ', text: error.message });
      }
    }

    async function importQualificationResultFile(input) {
      const file = input && input.files && input.files[0];
      if (input) input.value = '';
      if (!file) return;
      showLoading('กำลังอ่านและประมวลผลไฟล์');
      try {
        await ensureXlsxLoaded();
        const records = await readQualificationResultRows(file);
        if (!records.length) throw new Error('ไม่พบข้อมูลรหัสนักศึกษาและ GPAX ในไฟล์ ระบบรองรับทั้ง Excel ปกติและไฟล์ HTML ที่บันทึกนามสกุลเป็น .xls');

        const result = await serverCall('processQualificationImport', state.token, records, file.name);
        closeLoading();

        const notFoundText = result.notFound && result.notFound.length
          ? `<div class="mt-2 text-warning">ไม่พบรหัสในระบบ ${result.notFound.length} รายการ</div>` : '';
        const invalidText = result.invalid && result.invalid.length
          ? `<div class="mt-1 text-warning">เกรดไม่ถูกต้อง ${result.invalid.length} รายการ</div>` : '';
        const lockedText = Number(result.skippedForwarded || 0) > 0
          ? `<div class="mt-1 text-muted">ข้ามรายการที่ส่งให้หน่วยงานแล้ว ${result.skippedForwarded} รายการ</div>` : '';

        await Swal.fire({
          icon: 'success',
          title: 'ประมวลผลผลตรวจเรียบร้อย',
          html: `
            <div class="text-start mx-auto" style="max-width:420px;">
              <div>จับคู่กับผู้สมัครได้ <strong>${Number(result.matched || 0).toLocaleString('th-TH')}</strong> คน</div>
              <div class="text-success">ผ่าน (GPAX ≥ ${QUALIFICATION_THRESHOLD.toFixed(2)}) <strong>${Number(result.passed || 0).toLocaleString('th-TH')}</strong> คน</div>
              <div class="text-danger">ไม่ผ่าน (GPAX &lt; ${QUALIFICATION_THRESHOLD.toFixed(2)}) <strong>${Number(result.failed || 0).toLocaleString('th-TH')}</strong> คน</div>
              ${notFoundText}${invalidText}${lockedText}
            </div>`
        });
        state.qualificationDrafts.clear();
        await renderQualification();
      } catch (error) {
        closeLoading();
        Swal.fire({ icon: 'error', title: 'นำเข้าผลตรวจไม่สำเร็จ', text: error.message });
      }
    }

    async function readQualificationResultRows(file) {
      const bytes = await file.arrayBuffer();
      const recordsFromHtml = readQualificationRowsFromHtml(bytes);
      if (recordsFromHtml.length) return recordsFromHtml;

      await ensureXlsxLoaded();

      let workbook;
      try {
        workbook = XLSX.read(bytes, { type: 'array', raw: false, cellText: true });
      } catch (err) {
        throw new Error('รูปแบบไฟล์ไม่รองรับ กรุณาใช้ไฟล์ .xls, .xlsx หรือ .csv ที่เปิดตารางข้อมูลได้');
      }

      const records = [];
      const seen = new Set();
      (workbook.SheetNames || []).forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) return;
        const rows = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: '',
          raw: false,
          blankrows: false
        });
        extractQualificationRecordsFromRows(rows, records, seen);
      });

      if (!records.length) {
        const text = decodeQualificationFileText(bytes);
        if (/WorksheetSource\s+HRef=.*\.files\//i.test(text) || /sheet001\.htm/i.test(text)) {
          throw new Error('ไฟล์ .xls นี้เป็น Excel แบบหลายไฟล์ และตัวข้อมูลจริงอยู่ในโฟลเดอร์ .files ที่ไม่ได้แนบมาด้วย กรุณาเปิดไฟล์แล้ว Save As เป็น Excel Workbook (.xls หรือ .xlsx) แบบไฟล์เดียวก่อนนำเข้า');
        }
      }
      return records;
    }

    function readQualificationRowsFromHtml(bytes) {
      const text = decodeQualificationFileText(bytes);
      const head = text.slice(0, 1000).toLowerCase();
      if (!head.includes('<html') && !head.includes('<!doctype html')) return [];

      const doc = new DOMParser().parseFromString(text, 'text/html');
      const records = [];
      const seen = new Set();

      doc.querySelectorAll('table').forEach(table => {
        const rows = [...table.querySelectorAll('tr')].map(tr =>
          [...tr.querySelectorAll('th,td')].map(cell =>
            String(cell.textContent || '')
              .replace(/\u00a0/g, ' ')
              .replace(/[\t\r\n]+/g, ' ')
              .replace(/\s+/g, ' ')
              .trim()
          )
        );
        extractQualificationRecordsFromRows(rows, records, seen);
      });

      return records;
    }

    function decodeQualificationFileText(bytes) {
      try {
        return new TextDecoder('utf-8').decode(bytes);
      } catch (_) {
        const arr = new Uint8Array(bytes);
        let out = '';
        const limit = Math.min(arr.length, 2000000);
        for (let i = 0; i < limit; i++) out += String.fromCharCode(arr[i]);
        return out;
      }
    }

    function normalizeQualificationHeader(value) {
      return String(value ?? '')
        .replace(/\u00a0/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    }

    function findQualificationColumnIndexes(rows) {
      for (let r = 0; r < Math.min(rows.length, 30); r++) {
        const row = Array.isArray(rows[r]) ? rows[r] : [];
        let studentIdIndex = -1;
        let gpaxIndex = -1;

        row.forEach((cell, index) => {
          const h = normalizeQualificationHeader(cell);
          if (studentIdIndex < 0 && (
            h === 'รหัสนักศึกษา' ||
            h.includes('รหัสนักศึกษา') ||
            h === 'student id' ||
            h === 'studentid'
          )) studentIdIndex = index;

          if (gpaxIndex < 0 && (
            h.includes('ผลการเรียนเฉลี่ย') ||
            h.includes('เกรดเฉลี่ย') ||
            h.includes('gpax')
          )) gpaxIndex = index;
        });

        if (studentIdIndex >= 0 && gpaxIndex >= 0) {
          return { headerRow: r, studentIdIndex, gpaxIndex };
        }
      }

      // รูปแบบมาตรฐานจากระบบตรวจสอบ: E = รหัสนักศึกษา, H = GPAX
      return { headerRow: -1, studentIdIndex: 4, gpaxIndex: 7 };
    }

    function parseVerifiedGpax(value) {
      const text = String(value ?? '')
        .replace(/,/g, '.')
        .replace(/\u00a0/g, ' ')
        .trim();
      if (!text) return null;

      // ตัวอย่างจากระบบภายนอก: "2.53 (1/2569)" ต้องอ่าน 2.53 ไม่อ่าน 1/2569
      const candidates = text.match(/\d+(?:\.\d+)?/g) || [];
      for (const token of candidates) {
        const n = Number(token);
        if (Number.isFinite(n) && n >= 0 && n <= 4) return n;
      }
      return null;
    }

    function extractQualificationRecordsFromRows(rows, records, seen) {
      if (!Array.isArray(rows) || !rows.length) return;
      const indexes = findQualificationColumnIndexes(rows);
      const startRow = indexes.headerRow >= 0 ? indexes.headerRow + 1 : 0;

      for (let index = startRow; index < rows.length; index++) {
        const row = Array.isArray(rows[index]) ? rows[index] : [];
        const idRaw = String(row[indexes.studentIdIndex] ?? '');
        const idMatch = idRaw.match(/\d{8,13}/);
        const studentId = idMatch ? idMatch[0] : '';
        if (!/^\d{8,13}$/.test(studentId) || seen.has(studentId)) continue;

        const gradeRaw = row[indexes.gpaxIndex];
        const verifiedGpax = parseVerifiedGpax(gradeRaw);
        if (verifiedGpax === null) {
          // มีรหัสจริงแต่ช่องเกรดอ่านไม่ได้ ให้ส่งไป backend เพื่อรายงานเป็น invalid
          if (String(gradeRaw ?? '').trim()) {
            seen.add(studentId);
            records.push({ studentId, verifiedGpax: null, sourceRow: index + 1 });
          }
          continue;
        }

        seen.add(studentId);
        records.push({ studentId, verifiedGpax, sourceRow: index + 1 });
      }
    }

    function downloadCsv(filename, rows) {
      const quote = value => `"${String(value ?? '').replace(/"/g, '""')}"`;
      const csv = '\uFEFF' + rows.map(row => row.map(quote).join(',')).join('\r\n');
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    }

    /* ======================================================
       FORWARDING
    ====================================================== */
    async function renderForwarding() {
      setHeader('ส่งรายชื่อให้หน่วยงานคัดเลือก', 'ส่งเฉพาะรายชื่อที่ผ่านคุณสมบัติไปยังหน่วยคัดเลือกปลายทาง');
      const content=document.getElementById('content');
      content.innerHTML=`<div class="dashboard-skeleton"><div class="spinner-border text-primary"></div><div>กำลังโหลดข้อมูลการส่งรายชื่อ</div></div>`;
      const groups=await serverCall('getAdminDepartmentForwarding',state.token);
      window.forwardingGroups=Array.isArray(groups)?groups:[];
      const total=window.forwardingGroups.reduce((n,g)=>n+Number(g.total||0),0);
      const passed=window.forwardingGroups.reduce((n,g)=>n+Number(g.pass||0),0);
      const sent=window.forwardingGroups.reduce((n,g)=>n+Number(g.forwarded||0),0);
      const waiting=window.forwardingGroups.reduce((n,g)=>n+Number(g.unforwardedPass||0),0);
      content.innerHTML=`
        <div class="d-flex flex-wrap justify-content-between gap-3 align-items-end mb-4"><div><div class="page-title">ส่งรายชื่อให้หน่วยงานคัดเลือก</div><div class="page-subtitle">แยกรายชื่อตามหน่วยคัดเลือกปลายทางเดียวกับ Dashboard และหน้ารับผลการคัดเลือก</div></div></div>
        <div class="row g-3 mb-4">${watermarkStatCard('fa-users',total,'ผู้สมัครทั้งหมด','yellow')}${watermarkStatCard('fa-user-check',passed,'ผ่านคุณสมบัติ','green')}${watermarkStatCard('fa-paper-plane',sent,'ส่งแล้ว','purple')}${watermarkStatCard('fa-clock',waiting,'ผ่านแล้วรอส่ง','orange')}</div>
        <div class="panel premium-panel"><div class="panel-header"><div><h2 class="panel-title">รายชื่อแยกตามหน่วยคัดเลือก</h2><div class="text-muted small mt-1">กดจำนวนเพื่อโหลดรายชื่อของหน่วยนั้นโดยตรง · Excel จะดึงข้อมูลใหม่จากระบบก่อนสร้างไฟล์</div></div><span class="analysis-badge">${window.forwardingGroups.length.toLocaleString('th-TH')} หน่วยคัดเลือก</span></div><div class="table-wrap"><table class="table table-hover align-middle"><thead><tr><th>กลุ่มงาน › หน่วยคัดเลือก</th><th class="text-center">ผู้สมัครทั้งหมด</th><th class="text-center">ผ่าน</th><th class="text-center">ไม่ผ่าน</th><th class="text-center">รอตรวจ</th><th class="text-center">สถานะการส่ง</th><th class="text-end">ดำเนินการ</th></tr></thead><tbody>${window.forwardingGroups.length?window.forwardingGroups.map((group,index)=>{const sentAll=group.pass>0&&group.unforwardedPass===0;const partial=group.forwarded>0&&group.unforwardedPass>0;return `<tr><td><div class="fw-semibold">${escapeHtml(group.label||group.selectionUnit||group.department)}</div><div class="text-muted table-subtext">${Number(group.forwarded||0).toLocaleString('th-TH')} ส่งแล้ว · ${Number(group.unforwardedPass||0).toLocaleString('th-TH')} ผ่านแล้วรอส่ง</div></td><td class="text-center"><button class="count-link" onclick="showForwardingApplicants(${index},'all')">${Number(group.total||0).toLocaleString('th-TH')} คน <i class="fa-solid fa-chevron-right ms-1"></i></button></td><td class="text-center"><button class="count-link success" onclick="showForwardingApplicants(${index},'pass')">${Number(group.pass||0).toLocaleString('th-TH')} คน</button></td><td class="text-center"><button class="count-link danger" onclick="showForwardingApplicants(${index},'fail')">${Number(group.fail||0).toLocaleString('th-TH')} คน</button></td><td class="text-center"><button class="count-link" onclick="showForwardingApplicants(${index},'pending')">${Number(group.pendingQualification||0).toLocaleString('th-TH')} คน</button></td><td class="text-center">${sentAll?`<span class="status-badge badge-pass">ส่งแล้ว</span>`:partial?`<span class="status-badge badge-info">ส่งแล้วบางส่วน</span>`:`<span class="status-badge badge-wait">ยังไม่ส่ง</span>`}</td><td class="text-end text-nowrap"><button class="btn btn-sm btn-outline-success me-1" onclick="downloadForwardingGroup(${index})"><i class="fa-solid fa-file-excel me-1"></i>Excel</button><button class="btn btn-sm ${group.unforwardedPass>0?'btn-primary':'btn-outline-secondary'}" ${group.unforwardedPass>0?'':'disabled'} onclick="sendToDepartment(${index})"><i class="fa-solid fa-paper-plane me-1"></i>${partial?`ส่งเพิ่ม ${group.unforwardedPass} คน`:group.unforwardedPass>0?'ส่งรายชื่อ':'ส่งแล้ว'}</button></td></tr>`;}).join(''):`<tr><td colspan="7" class="text-center text-muted py-5">ยังไม่มีข้อมูล</td></tr>`}</tbody></table></div></div>`;
    }

    function qualificationResultLabel(value) {
      const result=String(value||'pending').toLowerCase();
      if(result==='pass') return 'ผ่านคุณสมบัติ';
      if(result==='fail') return 'ไม่ผ่านคุณสมบัติ';
      return 'ยังไม่ตรวจสอบ';
    }

    async function fetchForwardingApplicants(group, type='all') {
      if(!group) return [];
      const rows=await serverCall('getAdminDepartmentForwardingApplicants',state.token,group.department,group.selectionUnit,type);
      return Array.isArray(rows)?rows:[];
    }

    async function showForwardingApplicants(index,type='all') {
      const group=(window.forwardingGroups||[])[index]; if(!group) return;
      const typeLabel=type==='pass'?'ผ่านคุณสมบัติ':type==='fail'?'ไม่ผ่านคุณสมบัติ':type==='pending'?'ยังไม่ตรวจสอบคุณสมบัติ':'ผู้สมัครทั้งหมด';
      document.getElementById('listModalTitle').textContent=`${typeLabel} · ${group.label||group.selectionUnit||group.department}`;
      document.getElementById('listModalBody').innerHTML=`<div class="text-center py-5"><div class="spinner-border spinner-border-sm text-primary me-2"></div>กำลังโหลดรายชื่อ</div>`;
      listModal.show();
      try {
        const rows=await fetchForwardingApplicants(group,type);
        document.getElementById('listModalBody').innerHTML=rows.length?`<div class="table-responsive"><table class="table table-sm align-middle"><thead><tr><th>รหัสนักศึกษา</th><th>ชื่อ-สกุล</th><th>คณะ</th><th>ตำแหน่งงาน</th><th>ผลคุณสมบัติ</th><th>สถานะส่ง</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${escapeHtml(x.studentId||'-')}</td><td>${escapeHtml(x.fullName||'-')}</td><td>${escapeHtml(x.faculty||'-')}</td><td>${escapeHtml(x.job||'-')}</td><td>${statusBadge(qualificationResultLabel(x.qualificationResult))}</td><td>${x.forwarded?'<span class="status-badge badge-info">ส่งแล้ว</span>':'<span class="status-badge badge-wait">ยังไม่ส่ง</span>'}</td></tr>`).join('')}</tbody></table></div>`:'<div class="text-center text-muted py-4">ไม่มีรายชื่อ</div>';
      } catch(error) {
        document.getElementById('listModalBody').innerHTML=`<div class="text-center text-danger py-4">${escapeHtml(error.message)}</div>`;
      }
    }

    async function downloadForwardingGroup(index){
      const group=(window.forwardingGroups||[])[index]; if(!group)return;
      const safe=safeFilenamePart(group.label||group.selectionUnit||group.department||'หน่วยงาน');
      showLoading('กำลังโหลดรายชื่อสำหรับ Excel');
      try {
        const rows=await fetchForwardingApplicants(group,'all');
        if(!rows.length) throw new Error('ไม่พบรายชื่อผู้สมัครของหน่วยงานนี้');
        await ensureXlsxLoaded();
        const aoa=[
          ['รหัสนักศึกษา','ชื่อ-สกุล','คณะ','GPAX','ตำแหน่งงาน','กลุ่มงาน','หน่วยคัดเลือก','ผลคุณสมบัติ','สถานะการส่ง'],
          ...rows.map(x=>[
            String(x.studentId||''),x.fullName||'',x.faculty||'',String(x.gpax||''),x.job||'',group.department||'',group.selectionUnit||group.department||'',qualificationResultLabel(x.qualificationResult),x.forwarded?'ส่งแล้ว':'ยังไม่ส่ง'
          ])
        ];
        const ws=XLSX.utils.aoa_to_sheet(aoa);
        ws['!cols']=[{wch:16},{wch:30},{wch:30},{wch:10},{wch:36},{wch:30},{wch:30},{wch:20},{wch:16}];
        const wb=XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb,ws,'รายชื่อผู้สมัคร');
        const bytes=XLSX.write(wb,{bookType:'xlsx',type:'array'});
        const blob=new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
        const url=URL.createObjectURL(blob);
        const link=document.createElement('a');
        link.href=url; link.download=`รายชื่อ_${safe}.xlsx`;
        document.body.appendChild(link); link.click(); link.remove(); URL.revokeObjectURL(url);
        closeLoading();
      } catch(error) {
        closeLoading();
        Swal.fire({icon:'error',title:'ดาวน์โหลด Excel ไม่สำเร็จ',text:error.message});
      }
    }

    async function sendToDepartment(index) {
      const group = window.forwardingGroups[index];
      const confirm = await Swal.fire({
        icon: 'question',
        title: 'ส่งรายชื่อให้หน่วยงาน',
        html: `ส่งผู้ผ่านคุณสมบัติที่ยังไม่ส่งจำนวน <strong>${group.unforwardedPass}</strong> คน<br>ไปยัง <strong>${escapeHtml(group.label || group.selectionUnit || group.department)}</strong>`,
        showCancelButton: true,
        confirmButtonText: 'ยืนยันส่ง',
        cancelButtonText: 'ยกเลิก',
        confirmButtonColor: '#165f9c'
      });
      if (!confirm.isConfirmed) return;
      showLoading('กำลังส่งรายชื่อ');
      try {
        const result = await serverCall('forwardQualifiedApplicants', state.token, group.department, group.selectionUnit);
        closeLoading();
        await Swal.fire({ icon: 'success', title: 'ส่งรายชื่อเรียบร้อย', text: result.message });
        await renderForwarding();
      } catch (error) {
        closeLoading();
        Swal.fire({ icon: 'error', title: 'ส่งรายชื่อไม่สำเร็จ', text: error.message });
      }
    }

    /* ======================================================
       GROUP LIST
    ====================================================== */
    async function showGroupList(mode, index, type) {
      if (mode === 'forward') {
        showForwardingApplicants(index, type === 'pass' ? 'pass' : 'fail');
        return;
      }

      const group = (window.resultGroups || [])[index];
      if (!group) return;
      const typeLabel = type === 'selected' ? 'พิจารณารับ' : type === 'rejected' ? 'ไม่พิจารณารับ' : 'รอผลการพิจารณา';
      document.getElementById('listModalTitle').textContent = `${group.label || group.selectionUnit || group.department} · ${typeLabel}`;
      document.getElementById('listModalBody').innerHTML = `<div class="text-center py-5"><div class="spinner-border spinner-border-sm text-primary me-2"></div>กำลังโหลดรายชื่อ</div>`;
      listModal.show();

      try {
        const list = await serverCall(
          'getAdminDepartmentResultApplicants',
          state.token,
          group.department,
          group.selectionUnit,
          type
        ) || [];
        document.getElementById('listModalBody').innerHTML = list.length ? `
          <div class="table-responsive">
            <table class="table table-sm">
              <thead><tr><th>รหัสนักศึกษา</th><th>ชื่อ-สกุล</th><th>คณะ</th><th>ตำแหน่งงาน</th></tr></thead>
              <tbody>${list.map(item => `
                <tr>
                  <td>${escapeHtml(item.studentId)}</td>
                  <td>${escapeHtml(item.fullName)}</td>
                  <td>${escapeHtml(item.faculty)}</td>
                  <td>${escapeHtml(item.job || '-')}</td>
                </tr>`).join('')}</tbody>
            </table>
          </div>` : `<div class="text-center text-muted py-4">ไม่มีรายชื่อ</div>`;
      } catch (error) {
        document.getElementById('listModalBody').innerHTML = `<div class="text-center text-danger py-4">${escapeHtml(error.message)}</div>`;
      }
    }

    function resultTypeLabel(type) {
      if (type === 'selected') return 'พิจารณารับ';
      if (type === 'rejected') return 'ไม่พิจารณารับ';
      return 'รอผลการพิจารณา';
    }

    function buildDepartmentResultRowsFromApplicants(group, applicants) {
      const unit = group.label || group.selectionUnit || group.department || '';
      return (Array.isArray(applicants) ? applicants : []).map(item => [
        unit,
        item.studentId || '',
        item.fullName || '',
        item.faculty || '',
        item.job || '',
        resultTypeLabel(item.resultType),
        item.status || ''
      ]);
    }

    function safeFilenamePart(value) {
      return String(value || 'รายชื่อ').replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 90) || 'รายชื่อ';
    }

    async function downloadDepartmentResult(index) {
      const group = (window.resultGroups || [])[index];
      if (!group) return;
      showLoading('กำลังเตรียมรายชื่อ');
      try {
        const applicants = await serverCall(
          'getAdminDepartmentResultApplicants',
          state.token,
          group.department,
          group.selectionUnit,
          'all'
        ) || [];
        const rows = buildDepartmentResultRowsFromApplicants(group, applicants);
        if (!rows.length) throw new Error('หน่วยงานนี้ยังไม่มีข้อมูลสำหรับดาวน์โหลด');
        const header = ['กลุ่มงาน › หน่วยคัดเลือก', 'รหัสนักศึกษา', 'ชื่อ-สกุล', 'คณะ', 'ตำแหน่งงาน', 'ผลการพิจารณา', 'สถานะ'];
        const name = safeFilenamePart(group.label || group.selectionUnit || group.department);
        closeLoading();
        downloadCsv(`ผลการคัดเลือก_${name}.csv`, [header, ...rows]);
      } catch (error) {
        closeLoading();
        Swal.fire({ icon: 'error', title: 'ดาวน์โหลดรายชื่อไม่สำเร็จ', text: error.message });
      }
    }

    async function downloadAllDepartmentResults() {
      showLoading('กำลังเตรียมรายชื่อทั้งหมด');
      try {
        const source = await serverCall('getAdminDepartmentResultsExport', state.token) || [];
        if (!source.length) throw new Error('ยังไม่มีรายชื่อจากหน่วยงานสำหรับดาวน์โหลด');
        const header = ['กลุ่มงาน › หน่วยคัดเลือก', 'รหัสนักศึกษา', 'ชื่อ-สกุล', 'คณะ', 'ตำแหน่งงาน', 'ผลการพิจารณา', 'สถานะ'];
        const rows = source.map(item => [
          item.label || '',
          item.studentId || '',
          item.fullName || '',
          item.faculty || '',
          item.job || '',
          resultTypeLabel(item.resultType),
          item.status || ''
        ]);
        closeLoading();
        downloadCsv('ผลการคัดเลือกจากหน่วยงาน_ทั้งหมด.csv', [header, ...rows]);
      } catch (error) {
        closeLoading();
        Swal.fire({ icon: 'error', title: 'ดาวน์โหลดรายชื่อไม่สำเร็จ', text: error.message });
      }
    }

    function filterAdminResults() {
      const input = document.getElementById('adminResultSearch');
      if (!input) return;
      const value = input.value.trim().toLowerCase();
      document.querySelectorAll('#adminResultBody tr[data-search]').forEach(row => {
        row.style.display = row.dataset.search.includes(value) ? '' : 'none';
      });
    }

    async function renderAdminResults() {
      setHeader('รับข้อมูลผลการคัดเลือกจากหน่วยงาน', 'ผลการพิจารณารายชื่อจากแต่ละหน่วยงาน');
      const content = document.getElementById('content');
      content.innerHTML = `<div class="text-center py-5"><div class="spinner-border text-primary"></div></div>`;
      const groups = await serverCall('getAdminDepartmentResults', state.token);
      window.resultGroups = groups || [];

      const totalSent = groups.reduce((sum, item) => sum + Number(item.sent || 0), 0);
      const totalSelected = groups.reduce((sum, item) => sum + Number(item.selected || 0), 0);
      const totalRejected = groups.reduce((sum, item) => sum + Number(item.rejected || 0), 0);
      const totalPending = groups.reduce((sum, item) => sum + Number(item.pending || 0), 0);

      content.innerHTML = `
        <div class="result-toolbar">
          <div>
            <div class="page-title">รับข้อมูลผลการคัดเลือกจากหน่วยงาน</div>
            <div class="page-subtitle">ตรวจสอบผลการคัดเลือก และดาวน์โหลดรายชื่อที่หน่วยงานส่งกลับ</div>
          </div>
          <button class="btn btn-outline-primary" onclick="downloadAllDepartmentResults()" ${groups.length ? '' : 'disabled'}>
            <i class="fa-solid fa-download me-1"></i>ดาวน์โหลดรายชื่อทั้งหมด
          </button>
        </div>
        <div class="result-summary mb-3">
          <div class="result-summary-item">
            <div class="result-summary-value">${totalSent.toLocaleString('th-TH')}</div>
            <div class="result-summary-label">รายชื่อที่ส่งพิจารณา</div>
          </div>
          <div class="result-summary-item">
            <div class="result-summary-value">${totalSelected.toLocaleString('th-TH')}</div>
            <div class="result-summary-label">พิจารณารับ</div>
          </div>
          <div class="result-summary-item">
            <div class="result-summary-value">${totalRejected.toLocaleString('th-TH')}</div>
            <div class="result-summary-label">ไม่พิจารณารับ</div>
          </div>
          <div class="result-summary-item">
            <div class="result-summary-value">${totalPending.toLocaleString('th-TH')}</div>
            <div class="result-summary-label">รอผล</div>
          </div>
        </div>
        <div class="panel">
          <div class="panel-header">
            <div>
              <h2 class="panel-title">ผลการพิจารณาแยกตามหน่วยคัดเลือก</h2>
              <div class="small text-muted mt-1">${groups.length} หน่วยคัดเลือก</div>
            </div>
            <div class="input-group input-group-sm" style="max-width:300px;">
              <span class="input-group-text bg-white border-end-0"><i class="fa-solid fa-magnifying-glass text-muted"></i></span>
              <input id="adminResultSearch" class="form-control border-start-0 ps-0" placeholder="ค้นหาหน่วยงาน" oninput="filterAdminResults()">
            </div>
          </div>
          <div class="table-wrap">
            <table class="table">
              <thead>
                <tr>
                  <th>กลุ่มงาน › หน่วยคัดเลือก</th>
                  <th class="text-center">ส่งพิจารณา</th>
                  <th class="text-center">พิจารณารับ</th>
                  <th class="text-center">ไม่พิจารณารับ</th>
                  <th class="text-center">รอผล</th>
                  <th class="text-center">สถานะ</th>
                  <th class="text-end">ไฟล์รายชื่อ</th>
                </tr>
              </thead>
              <tbody id="adminResultBody">
                ${groups.length ? groups.map((group, index) => {
                  const label = group.label || group.selectionUnit || group.department;
                  const completed = Number(group.pending || 0) === 0 && Number(group.sent || 0) > 0;
                  return `
                    <tr data-search="${escapeHtml(String(label).toLowerCase())}">
                      <td>
                        <div class="fw-medium">${escapeHtml(label)}</div>
                      </td>
                      <td class="text-center">${Number(group.sent || 0).toLocaleString('th-TH')}</td>
                      <td class="text-center">
                        <button class="result-link result-pass" onclick="showGroupList('result', ${index}, 'selected')">
                          ${Number(group.selected || 0).toLocaleString('th-TH')} คน
                        </button>
                      </td>
                      <td class="text-center">
                        <button class="result-link result-fail" onclick="showGroupList('result', ${index}, 'rejected')">
                          ${Number(group.rejected || 0).toLocaleString('th-TH')} คน
                        </button>
                      </td>
                      <td class="text-center">
                        <button class="result-link result-wait" onclick="showGroupList('result', ${index}, 'pending')">
                          ${Number(group.pending || 0).toLocaleString('th-TH')} คน
                        </button>
                      </td>
                      <td class="text-center">
                        ${completed ? `<span class="status-badge badge-pass">ส่งผลครบแล้ว</span>` : `<span class="status-badge badge-wait">รอผลจากหน่วยงาน</span>`}
                      </td>
                      <td class="text-end">
                        <button class="btn btn-outline-secondary btn-sm" onclick="downloadDepartmentResult(${index})" ${Number(group.sent || 0) > 0 ? '' : 'disabled'}>
                          <i class="fa-solid fa-file-arrow-down me-1"></i>ดาวน์โหลดรายชื่อ
                        </button>
                      </td>
                    </tr>
                  `;
                }).join('') : `<tr><td colspan="7" class="text-center text-muted py-4">ยังไม่มีข้อมูลที่ส่งไปยังหน่วยงาน</td></tr>`}
              </tbody>
            </table>
          </div>
        </div>
      `;
    }

    /* ======================================================
       ANNOUNCEMENT EXCEL
    ====================================================== */
    function announcementSortNumber(value) {
      const text = String(value || '').trim();
      return /^\d{2}$/.test(text) ? Number(text) : 999;
    }

    function sortAnnouncementApplicants(rows) {
      return (Array.isArray(rows) ? rows.slice() : []).sort((a, b) => {
        const aId = String(a.studentId || '').trim();
        const bId = String(b.studentId || '').trim();
        const facultyDiff = announcementSortNumber(aId.slice(2, 4)) - announcementSortNumber(bId.slice(2, 4));
        if (facultyDiff) return facultyDiff;
        const yearDiff = announcementSortNumber(aId.slice(0, 2)) - announcementSortNumber(bId.slice(0, 2));
        if (yearDiff) return yearDiff;
        const idDiff = aId.localeCompare(bId, 'th');
        if (idDiff) return idDiff;
        return String(a.fullName || '').localeCompare(String(b.fullName || ''), 'th');
      });
    }

    function thaiDateInputDefault() {
      try {
        return new Intl.DateTimeFormat('th-TH', {
          timeZone: 'Asia/Bangkok', day: 'numeric', month: 'long', year: 'numeric'
        }).format(new Date());
      } catch (_) {
        return '';
      }
    }

    function filterAnnouncementPreview() {
      const value = (document.getElementById('announcementSearch')?.value || '').trim().toLowerCase();
      document.querySelectorAll('#announcementPreviewBody tr[data-search]').forEach(row => {
        row.style.display = !value || row.dataset.search.includes(value) ? '' : 'none';
      });
    }

    function downloadBase64Blob(base64, mimeType, filename) {
      const clean = String(base64 || '').replace(/\s/g, '');
      if (!clean) throw new Error('ระบบไม่ได้ส่งข้อมูลไฟล์กลับมา');
      const binary = atob(clean);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      const blob = new Blob([bytes], {
        type: mimeType || 'application/octet-stream'
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = filename || 'รายชื่อนักศึกษาทำงานระหว่างเรียน.xlsx';
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1500);
    }

    function formatAnnouncementReportDateTime(dateValue) {
      const value = dateValue instanceof Date ? dateValue : new Date();

      try {
        const parts = new Intl.DateTimeFormat('th-TH', {
          timeZone: 'Asia/Bangkok',
          day: '2-digit',
          month: '2-digit',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit',
          hour12: false
        }).formatToParts(value).reduce((acc, part) => {
          acc[part.type] = part.value;
          return acc;
        }, {});

        return `วันที่ ${parts.day}/${parts.month}/${parts.year} เวลา ${parts.hour}:${parts.minute} น.`;
      } catch (_) {
        return '';
      }
    }

    function announcementExcelFilename(dateValue) {
      const value = dateValue instanceof Date ? dateValue : new Date();
      try {
        const parts = new Intl.DateTimeFormat('en-GB', {
          timeZone: 'Asia/Bangkok',
          year: 'numeric',
          month: '2-digit',
          day: '2-digit',
          hour: '2-digit',
          minute: '2-digit',
          hour12: false
        }).formatToParts(value).reduce((acc, part) => {
          acc[part.type] = part.value;
          return acc;
        }, {});

        return `รายชื่อนักศึกษาทำงานระหว่างเรียน_${parts.year}${parts.month}${parts.day}_${parts.hour}${parts.minute}.xlsx`;
      } catch (_) {
        return 'รายชื่อนักศึกษาทำงานระหว่างเรียน.xlsx';
      }
    }

    function downloadArrayBufferFile(buffer, mimeType, filename) {
      const blob = new Blob([buffer], {
        type: mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = filename || 'รายชื่อนักศึกษาทำงานระหว่างเรียน.xlsx';
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1500);
    }

    async function loadAnnouncementLogoImage(workbook) {
      try {
        // ฝังโลโก้ไว้ใน dashboard.js โดยตรง เพื่อไม่ให้เกิดปัญหา path / cache / 404
        return workbook.addImage({
          base64: 'data:image/png;base64,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',
          extension: 'png'
        });
      } catch (error) {
        console.warn('[Announcement Excel] Logo:', error);
        return null;
      }
    }

    async function generateAnnouncementExcel() {
      const rows = sortAnnouncementApplicants(window.announcementApplicants || []);

      if (!rows.length) {
        Swal.fire({
          icon: 'info',
          title: 'ยังไม่มีรายชื่อ',
          text: 'ยังไม่มีผู้ผ่านการคัดเลือกสำหรับจัดทำประกาศ'
        });
        return;
      }

      showLoading('กำลังจัดทำ Excel พร้อมรูปแบบสำหรับพิมพ์');

      try {
        const ExcelJS = await ensureExcelJsLoaded();
        const generatedAt = new Date();
        const reportDateTime = formatAnnouncementReportDateTime(generatedAt);

        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'มหาวิทยาลัยอุบลราชธานี';
        workbook.lastModifiedBy = 'ระบบรับสมัคร';
        workbook.created = generatedAt;
        workbook.modified = generatedAt;
        workbook.subject = 'รายชื่อนักศึกษาทำงานระหว่างเรียน';
        workbook.title = 'รายชื่อนักศึกษาทำงานระหว่างเรียน';

        const sheet = workbook.addWorksheet('รายชื่อนักศึกษา', {
          properties: {
            defaultRowHeight: 20
          },
          pageSetup: {
            paperSize: 9,                 // A4
            orientation: 'portrait',
            fitToPage: true,
            fitToWidth: 1,
            fitToHeight: 0,
            horizontalCentered: true,
            verticalCentered: false,
            margins: {
              left: 0.32,
              right: 0.32,
              top: 0.32,
              bottom: 0.55,
              header: 0.15,
              footer: 0.22
            }
          }
        });

        const logoImageId = await loadAnnouncementLogoImage(workbook);

        const FONT_NAME = 'TH SarabunPSK';
        const FONT_SIZE = 14;

        function applyBorder(cell) {
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF555555' } },
            left: { style: 'thin', color: { argb: 'FF555555' } },
            bottom: { style: 'thin', color: { argb: 'FF555555' } },
            right: { style: 'thin', color: { argb: 'FF555555' } }
          };
        }

        // ====================================================
        // หัวรายงานแสดงใน Sheet เพียงครั้งเดียว
        // ตอน Print ใช้ printTitlesRow ให้ Excel ทำซ้ำอัตโนมัติทุกหน้า
        // ====================================================
        sheet.mergeCells('A1:B1');
        sheet.mergeCells('C1:D1');
        sheet.mergeCells('A2:B2');
        sheet.mergeCells('C2:D2');
        sheet.mergeCells('A3:B3');
        sheet.mergeCells('C3:D3');

        const left1 = sheet.getCell('A1');
        left1.value = 'มหาวิทยาลัยอุบลราชธานี';
        left1.font = { name: FONT_NAME, size: FONT_SIZE, bold: true };
        left1.alignment = {
          vertical: 'middle',
          horizontal: 'left',
          // เว้นพื้นที่ให้โลโก้ ไม่ให้ทับข้อความ
          indent: logoImageId !== null ? 9 : 0
        };

        const right1 = sheet.getCell('C1');
        right1.value = 'รายชื่อนักศึกษาทำงานระหว่างเรียน';
        right1.font = { name: FONT_NAME, size: FONT_SIZE, bold: true };
        right1.alignment = {
          vertical: 'middle',
          horizontal: 'right'
        };

        const left2 = sheet.getCell('A2');
        left2.value = 'ระบบรับสมัคร';
        left2.font = { name: FONT_NAME, size: FONT_SIZE, bold: false };
        left2.alignment = {
          vertical: 'middle',
          horizontal: 'left',
          indent: logoImageId !== null ? 9 : 0
        };

        const right2 = sheet.getCell('C2');
        right2.value = 'ครั้งที่ 1 ปีงบประมาณ 2570';
        right2.font = { name: FONT_NAME, size: FONT_SIZE, bold: false };
        right2.alignment = {
          vertical: 'middle',
          horizontal: 'right'
        };

        const left3 = sheet.getCell('A3');
        left3.value = '';
        left3.font = { name: FONT_NAME, size: FONT_SIZE };

        const right3 = sheet.getCell('C3');
        right3.value = '(1 ตุลาคม 2569 - 30 กันยายน 2570)';
        right3.font = { name: FONT_NAME, size: FONT_SIZE, bold: false };
        right3.alignment = {
          vertical: 'middle',
          horizontal: 'right'
        };

        // หัวกระดาษแบบกะทัดรัด ไม่แบ่งพื้นที่ใหญ่
        sheet.getRow(1).height = 21;
        sheet.getRow(2).height = 19;
        sheet.getRow(3).height = 19;

        // โลโก้เล็กลงและวางชิดซ้าย เพื่อไม่ทับตัวหนังสือ
        if (logoImageId !== null) {
          sheet.addImage(logoImageId, {
            tl: {
              col: 0.10,
              row: 0.18
            },
            ext: {
              width: 34,
              height: 42
            },
            editAs: 'oneCell'
          });
        }

        // เส้นคั่นบาง ๆ
        sheet.getRow(4).height = 5;
        for (let col = 1; col <= 4; col++) {
          sheet.getCell(4, col).border = {
            bottom: {
              style: 'thin',
              color: { argb: 'FF595959' }
            }
          };
        }

        // ====================================================
        // หัวตาราง - แถว 5
        // ====================================================
        const tableHeaderRow = sheet.getRow(5);
        tableHeaderRow.values = [
          'ลำดับที่',
          'รหัสนักศึกษา',
          'ชื่อ-สกุล',
          'คณะ'
        ];
        tableHeaderRow.height = 23;

        tableHeaderRow.eachCell(cell => {
          cell.font = {
            name: FONT_NAME,
            size: FONT_SIZE,
            bold: true
          };
          cell.alignment = {
            horizontal: 'center',
            vertical: 'middle',
            wrapText: true
          };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE7E9ED' }
          };
          applyBorder(cell);
        });

        // ====================================================
        // รายชื่อ - ต่อเนื่องยาวครั้งเดียว ไม่แบ่งหน้าใน Sheet
        // Excel จะจัดหน้าเองเมื่อ Print
        // ====================================================
        rows.forEach((item, index) => {
          const row = sheet.addRow([
            index + 1,
            String(item.studentId || ''),
            String(item.fullName || ''),
            String(item.faculty || '')
          ]);

          const longText =
            String(item.fullName || '').length > 32 ||
            String(item.faculty || '').length > 30;

          row.height = longText ? 29 : 21;

          row.eachCell((cell, colNumber) => {
            cell.font = {
              name: FONT_NAME,
              size: FONT_SIZE,
              bold: false
            };
            cell.alignment = {
              vertical: 'middle',
              horizontal: colNumber <= 2 ? 'center' : 'left',
              wrapText: true
            };
            applyBorder(cell);
          });

          row.getCell(2).numFmt = '@';
        });

        // จำนวนรวมท้ายรายการ
        const totalRowNumber = 6 + rows.length;
        sheet.mergeCells(`A${totalRowNumber}:D${totalRowNumber}`);

        const totalCell = sheet.getCell(`A${totalRowNumber}`);
        totalCell.value = `จำนวนทั้งสิ้น ${rows.length.toLocaleString('th-TH')} คน`;
        totalCell.font = {
          name: FONT_NAME,
          size: FONT_SIZE,
          bold: true
        };
        totalCell.alignment = {
          horizontal: 'right',
          vertical: 'middle'
        };
        sheet.getRow(totalRowNumber).height = 22;

        // ความกว้างคอลัมน์
        sheet.getColumn(1).width = 10;
        sheet.getColumn(2).width = 19;
        sheet.getColumn(3).width = 38;
        sheet.getColumn(4).width = 34;

        // เวลาเปิด Excel ให้หัวตารางติดด้านบน
        sheet.views = [{
          state: 'frozen',
          ySplit: 5,
          xSplit: 0,
          topLeftCell: 'A6',
          activeCell: 'A6'
        }];

        // ====================================================
        // การพิมพ์
        // แถว 1-5 จะแสดงซ้ำเฉพาะตอน Print / Print Preview
        // ไม่สร้างหัวซ้ำอยู่กลาง Sheet
        // ====================================================
        sheet.pageSetup.printTitlesRow = '1:5';
        sheet.pageSetup.printArea = `A1:D${totalRowNumber}`;
        sheet.pageSetup.fitToPage = true;
        sheet.pageSetup.fitToWidth = 1;
        sheet.pageSetup.fitToHeight = 0;
        sheet.pageSetup.paperSize = 9;
        sheet.pageSetup.orientation = 'portrait';
        sheet.pageSetup.horizontalCentered = true;

        // ส่วนท้ายทุกหน้า: TH SarabunPSK 14 pt
        const footerFont = '&"TH SarabunPSK,Regular"&14';

        sheet.headerFooter = {
          oddFooter:
            `&L${footerFont}${reportDateTime}` +
            `&R${footerFont}หน้า &P/&N`
        };

        const buffer = await workbook.xlsx.writeBuffer();

        downloadArrayBufferFile(
          buffer,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          announcementExcelFilename(generatedAt)
        );

        closeLoading();

        Swal.fire({
          icon: 'success',
          title: 'ดาวน์โหลด Excel สำเร็จ',
          html:
            `จัดทำรายชื่อจำนวน <strong>${rows.length.toLocaleString('th-TH')} คน</strong><br>` +
            `<span class="text-muted">หัวรายงานแสดงครั้งเดียวใน Sheet และทำซ้ำอัตโนมัติเฉพาะตอนพิมพ์</span>`,
          timer: 2600,
          showConfirmButton: false
        });

      } catch (error) {
        closeLoading();

        Swal.fire({
          icon: 'error',
          title: 'จัดทำ Excel ไม่สำเร็จ',
          text: error && error.message
            ? error.message
            : 'เกิดข้อผิดพลาดในการจัดทำไฟล์ Excel'
        });
      }
    }

    async function renderAnnouncementExcel() {
      setHeader('จัดทำประกาศรายชื่อ', 'ดาวน์โหลดรายชื่อเป็นไฟล์ Excel (.xlsx)');
      const content = document.getElementById('content');
      content.innerHTML = `<div class="dashboard-skeleton"><div class="spinner-border text-primary"></div><div>กำลังเตรียมรายชื่อผู้ผ่านการคัดเลือก</div></div>`;

      const source = await serverCall('getAdminAnnouncementApplicants', state.token) || [];
      const rows = sortAnnouncementApplicants(source);
      window.announcementApplicants = rows;
      const facultyCount = new Set(rows.map(item => String(item.studentId || '').slice(2, 4)).filter(Boolean)).size;
      const yearCount = new Set(rows.map(item => String(item.studentId || '').slice(0, 2)).filter(Boolean)).size;

      content.innerHTML = `
        <div class="announcement-hero">
          <div>
            <div class="page-title">จัดทำประกาศรายชื่อผู้ผ่านการคัดเลือก</div>
            <div class="page-subtitle">ดาวน์โหลดรายชื่อนักศึกษาทำงานระหว่างเรียนเป็น Excel พร้อมรูปแบบสำหรับพิมพ์ทุกหน้า</div>
          </div>
          <button class="btn btn-success" onclick="generateAnnouncementExcel()" ${rows.length ? '' : 'disabled'}>
            <i class="fa-solid fa-file-excel me-2"></i>ดาวน์โหลด Excel
          </button>
        </div>

        <div class="row g-3 mb-4">
          ${watermarkStatCard('fa-circle-check', rows.length, 'ผู้ผ่านการคัดเลือก', 'green')}
          ${watermarkStatCard('fa-building-columns', facultyCount, 'รหัสคณะในรายชื่อ', 'blue')}
          ${watermarkStatCard('fa-calendar-days', yearCount, 'ปีรหัสนักศึกษาในรายชื่อ', 'yellow')}
        </div>

        <div class="panel announcement-settings-panel mb-4">
          <div class="panel-header">
            <div>
              <h2 class="panel-title">รูปแบบไฟล์ Excel</h2>
              <div class="small text-muted mt-1">ระบบจะสร้าง .xlsx ด้วย TH SarabunPSK 14 pt โดยหัวรายงานแสดงครั้งเดียวใน Sheet และทำซ้ำอัตโนมัติเฉพาะตอนพิมพ์</div>
            </div>
            <span class="badge text-bg-light border">
              <i class="fa-solid fa-print me-1"></i>A4 · หัวซ้ำทุกหน้า · หน้า x/y
            </span>
          </div>
          <div class="p-3 pt-0">
            <div class="alert alert-light border mb-0">
              <i class="fa-solid fa-wand-magic-sparkles me-2 text-primary"></i>
              เมื่อกดดาวน์โหลด ระบบจะดึงเฉพาะรายชื่อที่หน่วยงานยืนยัน <strong>พิจารณารับ</strong> จากทุกหน่วยงาน
              แล้วสร้าง Excel พร้อมหัวรายงานตามรูปแบบมหาวิทยาลัย ตาราง <strong>ลำดับที่ | รหัสนักศึกษา | ชื่อ-สกุล | คณะ</strong>
              และส่วนท้ายวันที่/เวลา/เลขหน้าอัตโนมัติ
            </div>
          </div>
        </div>

        <div class="panel announcement-preview-panel">
          <div class="panel-header announcement-preview-header">
            <div>
              <h2 class="panel-title">ตัวอย่างลำดับรายชื่อ</h2>
              <div class="small text-muted mt-1">ตาราง Excel: ลำดับที่ · รหัสนักศึกษา · ชื่อ-สกุล · คณะ</div>
            </div>
            <div class="search-box announcement-search-box">
              <i class="fa-solid fa-magnifying-glass"></i>
              <input id="announcementSearch" class="form-control form-control-sm" placeholder="ค้นหารหัส ชื่อ หรือคณะ" oninput="filterAnnouncementPreview()">
            </div>
          </div>
          <div class="announcement-sort-note">
            <i class="fa-solid fa-arrow-down-wide-short"></i>
            <span>หลักการเรียง: รหัสคณะ = ตัวที่ 3-4 ↑ &nbsp;→&nbsp; ปีเข้าศึกษา = ตัวที่ 1-2 ↑ &nbsp;→&nbsp; รหัสนักศึกษาทั้งหมด ↑</span>
          </div>
          <div class="table-wrap">
            <table class="table table-hover align-middle announcement-table">
              <thead><tr><th class="text-center" style="width:90px">ลำดับที่</th><th>รหัสนักศึกษา</th><th>ชื่อ-สกุล</th><th>คณะ</th></tr></thead>
              <tbody id="announcementPreviewBody">
                ${rows.length ? rows.map((item, index) => `
                  <tr data-search="${escapeHtml(`${item.studentId || ''} ${item.fullName || ''} ${item.faculty || ''}`.toLowerCase())}">
                    <td class="text-center fw-semibold">${index + 1}</td>
                    <td><span class="candidate-code">${escapeHtml(item.studentId || '-')}</span><div class="announcement-code-hint">คณะ ${escapeHtml(String(item.studentId || '').slice(2,4) || '-')} · ปี ${escapeHtml(String(item.studentId || '').slice(0,2) || '-')}</div></td>
                    <td><span class="candidate-name">${escapeHtml(item.fullName || '-')}</span></td>
                    <td>${escapeHtml(item.faculty || '-')}</td>
                  </tr>`).join('') : `<tr><td colspan="4" class="text-center text-muted py-5">ยังไม่มีรายชื่อผู้ผ่านการคัดเลือก</td></tr>`}
              </tbody>
            </table>
          </div>
          <div class="announcement-count-footer">จำนวนทั้งสิ้น <strong>${rows.length.toLocaleString('th-TH')}</strong> คน</div>
        </div>
      `;
    }

    function getBasketCount() {
      return (state.departmentRows || []).filter(item => item.inBasket && !item.finalized).length;
    }

    function updateBasketIndicator() {
      const count = getBasketCount();
      const badge = document.getElementById('basketCountBadge');
      const countText = document.getElementById('basketCountText');
      const sidebarBadge = document.getElementById('sidebarBasketBadge');

      if (badge) badge.textContent = count;
      if (countText) countText.textContent = `${count} รายชื่อ`;
      if (sidebarBadge) {
        sidebarBadge.textContent = count;
        sidebarBadge.classList.toggle('d-none', count === 0);
      }
    }

    function openBasketPage() {
      navigate('departmentBasket');
    }

    async function renderDepartmentSelection() {
      setHeader('คัดเลือกรายชื่อลงตะกร้า', state.selectionUnit || state.department);
      const content=document.getElementById('content'); content.innerHTML=`<div class="dashboard-skeleton"><div class="spinner-border text-primary"></div><div>กำลังโหลดรายชื่อ</div></div>`;
      const rows=await serverCall('getDepartmentApplicants',state.token); state.departmentRows=Array.isArray(rows)?rows:[]; state.departmentPage=1;
      const activeRows=state.departmentRows.filter(item=>!item.finalized); const basketCount=getBasketCount();
      content.innerHTML=`
        <div class="d-flex flex-wrap justify-content-between gap-3 align-items-end mb-4"><div><div class="page-title">คัดเลือกรายชื่อลงตะกร้า</div><div class="page-subtitle">แสดงเฉพาะรายชื่อที่ส่งมายัง ${escapeHtml(state.selectionUnit||state.department)}</div></div><button class="basket-widget basket-widget-light" onclick="openBasketPage()"><span class="basket-icon-wrap"><i class="fa-solid fa-basket-shopping"></i><span id="basketCountBadge" class="basket-badge">${basketCount}</span></span><span><span class="basket-widget-label d-block">ตะกร้ารายชื่อ</span><span id="basketCountText" class="basket-widget-count d-block">${basketCount} รายชื่อ</span></span><i class="fa-solid fa-chevron-right ms-auto"></i></button></div>
        <div class="row g-3 mb-4">${watermarkStatCard('fa-inbox',state.departmentRows.length,'รายชื่อที่ได้รับ','yellow')}${watermarkStatCard('fa-user-clock',activeRows.filter(x=>!x.inBasket).length,'รอพิจารณา','blue')}${watermarkStatCard('fa-basket-shopping',basketCount,'อยู่ในตะกร้า','orange')}${watermarkStatCard('fa-circle-check',state.departmentRows.filter(x=>x.finalized&&x.departmentDecision==='selected').length,'คัดเลือกแล้ว','green')}</div>
        <div class="panel premium-panel"><div class="panel-header filter-toolbar"><div><h2 class="panel-title">รายชื่อสำหรับพิจารณา</h2><div class="text-muted small">จำกัดจำนวนแถวเพื่อให้ใช้งานได้ลื่นขึ้น</div></div><div class="filter-grid department-filter-grid"><select id="departmentPageSize" class="form-select form-select-sm" onchange="changeDepartmentPageSize(this.value)"><option value="20">20 แถว</option><option value="50">50 แถว</option><option value="100">100 แถว</option></select><div class="search-box"><i class="fa-solid fa-magnifying-glass"></i><input id="departmentSearch" class="form-control form-control-sm" placeholder="ค้นหารหัส / ชื่อ / คณะ" oninput="filterDepartmentTable()"></div></div></div><div class="table-wrap"><table class="table table-hover align-middle"><thead><tr><th class="text-center">เลือก</th><th>รูปภาพ</th><th>รหัสนักศึกษา</th><th>ชื่อ-สกุล</th><th>คณะ</th><th>ตำแหน่งงาน</th><th>ข้อมูลผู้สมัคร</th><th>สถานะ</th></tr></thead><tbody id="departmentBody"></tbody></table></div><div id="departmentPager" class="pager-shell"></div></div>`;
      refreshDepartmentPage(); updateBasketIndicator();
    }


    function getDepartmentFilteredRows(){const value=(document.getElementById('departmentSearch')?.value||'').trim().toLowerCase();return (state.departmentRows||[]).filter(item=>!item.finalized).filter(item=>!value||`${item.studentId} ${item.fullName} ${item.faculty} ${item.job}`.toLowerCase().includes(value));}
    function refreshDepartmentPage(){const rows=getDepartmentFilteredRows();const size=Number(state.departmentPageSize)||20;const pages=Math.max(1,Math.ceil(rows.length/size));state.departmentPage=Math.min(Math.max(1,state.departmentPage),pages);const start=(state.departmentPage-1)*size;const pageRows=rows.slice(start,start+size);const body=document.getElementById('departmentBody');if(body)body.innerHTML=pageRows.length?pageRows.map(item=>`<tr class="${item.inBasket&&!item.finalized?'selected-row':''}"><td class="text-center"><span class="basket-checkbox-wrap"><input class="form-check-input" type="checkbox" ${item.inBasket?'checked':''} onchange="toggleBasket('${escapeHtml(item.applicationId)}',this)"></span></td><td><div id="photo-wrap-${escapeHtml(item.applicationId)}" class="photo-placeholder"><i class="fa-regular fa-user"></i></div></td><td><span class="candidate-code">${escapeHtml(item.studentId)}</span></td><td><span class="candidate-name">${escapeHtml(item.fullName)}</span></td><td>${escapeHtml(item.faculty)}</td><td>${escapeHtml(item.job)}</td><td><button class="btn btn-outline-primary btn-sm" onclick="showApplicantDetail('${escapeHtml(item.applicationId)}')"><i class="fa-regular fa-eye me-1"></i>ดูข้อมูล</button></td><td class="row-status">${statusBadge(item.inBasket?'อยู่ในตะกร้า':'รอพิจารณา')}</td></tr>`).join(''):`<tr><td colspan="8" class="text-center text-muted py-5">ไม่พบรายชื่อ</td></tr>`;renderSimplePager('departmentPager',rows.length,state.departmentPage,size,'goDepartmentPage');loadDepartmentPhotos(pageRows);}
    function filterDepartmentTable(){state.departmentPage=1;refreshDepartmentPage();}
    function changeDepartmentPageSize(v){state.departmentPageSize=[20,50,100].includes(Number(v))?Number(v):20;state.departmentPage=1;refreshDepartmentPage();}
    function goDepartmentPage(p){state.departmentPage=Number(p)||1;refreshDepartmentPage();}


    async function loadDepartmentPhotos(rows) {
      // FAST V7: รูปเป็น low priority + cache ใน browser; โหลดจาก Drive เฉพาะรูปที่ยังไม่เคยเห็น
      const generation = Number(state.photoLoadGeneration || 0);
      const allIds = rows.filter(item => item.hasPhoto).map(item => item.applicationId);
      if (!allIds.length) return;

      const missingIds = [];
      allIds.forEach(id => {
        const cached = applicantPhotoCache.get(id);
        const wrapper = document.getElementById(`photo-wrap-${id}`);
        if (cached && wrapper) {
          wrapper.outerHTML = `<img id="photo-wrap-${escapeHtml(id)}" class="photo-thumb" src="${cached}" alt="รูปผู้สมัคร">`;
        } else if (!cached) {
          missingIds.push(id);
        }
      });

      if (!missingIds.length) return;
      await new Promise(resolve => setTimeout(resolve, 450));
      if (generation !== Number(state.photoLoadGeneration || 0) || state.currentView !== 'departmentSelection') return;

      const chunks=[];
      for (let i=0;i<missingIds.length;i+=6) chunks.push(missingIds.slice(i,i+6));
      for (const chunk of chunks) {
        if (generation !== Number(state.photoLoadGeneration || 0) || state.currentView !== 'departmentSelection') return;
        try {
          const photos=await serverCall('getStaffApplicantPhotoBatch',state.token,chunk);
          if (generation !== Number(state.photoLoadGeneration || 0)) return;
          chunk.forEach(id=>{
            const data=photos&&photos[id];
            const wrapper=document.getElementById(`photo-wrap-${id}`);
            if (!data) return;
            rememberApplicantPhoto(id, data);
            if (!wrapper) return;
            wrapper.outerHTML=`<img id="photo-wrap-${escapeHtml(id)}" class="photo-thumb" src="${data}" alt="รูปผู้สมัคร">`;
          });
        } catch(error) { console.warn('ไม่สามารถโหลดรูปผู้สมัครบางรายการได้',error); }
        await new Promise(resolve=>setTimeout(resolve,80));
      }
    }

    async function toggleBasket(applicationId, checkbox) {
      const row = checkbox.closest('tr');
      const previous = !checkbox.checked;
      const nextSelected = !!checkbox.checked;
      checkbox.disabled = true;

      // FAST V10: optimistic UI — หน้าจอตอบสนองทันที ไม่รอ Apps Script round-trip
      const item = (state.departmentRows || []).find(record => record.applicationId === applicationId);
      if (item) item.inBasket = nextSelected;
      if (row) {
        row.classList.toggle('selected-row', nextSelected);
        const statusCell = row.querySelector('.row-status');
        if (statusCell) statusCell.innerHTML = statusBadge(nextSelected ? 'อยู่ในตะกร้า' : 'รอพิจารณา');
      }
      updateBasketIndicator();

      try {
        await serverCall('setDepartmentBasket', state.token, applicationId, nextSelected);
      } catch (error) {
        checkbox.checked = previous;
        if (item) item.inBasket = previous;
        if (row) {
          row.classList.toggle('selected-row', previous);
          const statusCell = row.querySelector('.row-status');
          if (statusCell) statusCell.innerHTML = statusBadge(previous ? 'อยู่ในตะกร้า' : 'รอพิจารณา');
        }
        updateBasketIndicator();
        Swal.fire({ icon: 'error', title: 'ทำรายการไม่สำเร็จ', text: error.message });
      } finally {
        checkbox.disabled = false;
      }
    }


    async function renderDepartmentBasket() {
      setHeader('ยืนยันส่งรายชื่อ', 'เมื่อยืนยันส่งแล้วจะไม่สามารถแก้ไขรายชื่อได้');
      const content = document.getElementById('content');
      content.innerHTML = `<div class="text-center py-5"><div class="spinner-border text-primary"></div></div>`;
      
      const rows = await serverCall('getDepartmentApplicants', state.token);
      state.departmentRows = rows || [];
      const basket = rows.filter(item => item.inBasket && !item.finalized);
      const completed = rows.filter(item => item.finalized && item.departmentDecision === 'selected');

      content.innerHTML = `
        <div class="d-flex flex-wrap justify-content-between gap-3 align-items-end mb-4">
          <div>
            <div class="page-title">ยืนยันส่งรายชื่อ</div>
            <div class="page-subtitle">รายชื่อในตะกร้าที่จะส่งไปจัดทำประกาศผู้ผ่านการพิจารณาหน่วยงานคัดเลือก</div>
          </div>
          <button class="btn btn-success" ${basket.length ? '' : 'disabled'} onclick="confirmDepartmentFinalization()">
            <i class="fa-solid fa-paper-plane me-1"></i>ยืนยันส่ง ${basket.length} รายชื่อ
          </button>
        </div>
        
        <div class="panel mb-4">
          <div class="panel-header">
            <h2 class="panel-title">รายชื่อในตะกร้า (${basket.length} คน)</h2>
          </div>
          <div class="table-wrap">
            <table class="table">
              <thead>
                <tr>
                  <th>ลำดับ</th>
                  <th>รหัสนักศึกษา</th>
                  <th>ชื่อ-สกุล</th>
                  <th>คณะ</th>
                  <th>ตำแหน่งงาน</th>
                  <th>ข้อมูล</th>
                </tr>
              </thead>
              <tbody>
                ${basket.length ? basket.map((item, index) => `
                  <tr>
                    <td>${index + 1}</td>
                    <td>${escapeHtml(item.studentId)}</td>
                    <td>${escapeHtml(item.fullName)}</td>
                    <td>${escapeHtml(item.faculty)}</td>
                    <td>${escapeHtml(item.job)}</td>
                    <td>
                      <button class="btn btn-outline-primary btn-sm" onclick="showApplicantDetail('${escapeHtml(item.applicationId)}')">ดูข้อมูล</button>
                    </td>
                  </tr>
                `).join('') : `<tr><td colspan="6" class="text-center text-muted py-4">ยังไม่มีรายชื่อในตะกร้า</td></tr>`}
              </tbody>
            </table>
          </div>
        </div>
        
        ${completed.length ? `
          <div class="panel">
            <div class="panel-header">
              <h2 class="panel-title">รายชื่อที่ยืนยันส่งแล้ว (${completed.length} คน)</h2>
            </div>
            <div class="table-wrap">
              <table class="table">
                <thead>
                  <tr>
                    <th>รหัสนักศึกษา</th>
                    <th>ชื่อ-สกุล</th>
                    <th>คณะ</th>
                    <th>สถานะ</th>
                  </tr>
                </thead>
                <tbody>
                  ${completed.map(item => `
                    <tr>
                      <td>${escapeHtml(item.studentId)}</td>
                      <td>${escapeHtml(item.fullName)}</td>
                      <td>${escapeHtml(item.faculty)}</td>
                      <td>${statusBadge('ส่งแล้ว')}</td>
                    </tr>
                  `).join('')}
                </tbody>
              </table>
            </div>
          </div>
        ` : ''}
      `;
      updateBasketIndicator();
    }

    async function confirmDepartmentFinalization() {
      const basket = state.departmentRows.filter(item => item.inBasket && !item.finalized);
      const confirm = await Swal.fire({
        icon: 'warning',
        title: 'ยืนยันส่งรายชื่อ?',
        html: `
          <div class="text-start">
            <p>ระบบจะส่งรายชื่อในตะกร้า <strong>${basket.length} คน</strong> ไปจัดทำประกาศจ้างงานในขั้นตอนถัดไป</p>
            <p class="text-danger mb-0"><strong>เมื่อส่งแล้วจะไม่สามารถแก้ไขรายชื่อได้</strong></p>
          </div>
        `,
        showCancelButton: true,
        confirmButtonText: 'ยืนยันส่งรายชื่อ',
        cancelButtonText: 'ยกเลิก',
        confirmButtonColor: '#198754'
      });

      if (!confirm.isConfirmed) return;

      showLoading('กำลังยืนยันส่งรายชื่อ');
      try {
        const result = await serverCall('finalizeDepartmentSelection', state.token);
        closeLoading();
        await Swal.fire({ icon: 'success', title: 'ส่งรายชื่อเรียบร้อย', text: result.message });
        await renderDepartmentBasket();
      } catch (error) {
        closeLoading();
        Swal.fire({ icon: 'error', title: 'ส่งรายชื่อไม่สำเร็จ', text: error.message });
      }
    }

    async function showApplicantDetail(applicationId) {
      showLoading('กำลังโหลดข้อมูลผู้สมัคร');
      try {
        const item = await serverCall('getStaffApplicantDetail', state.token, applicationId);
        closeLoading();
        
        const details = [
          ['เลขที่ใบสมัคร', item.applicationId],
          ['วันที่สมัคร', item.submittedAt],
          ['รหัสนักศึกษา', item.studentId],
          ['ชื่อ-สกุล', item.fullName],
          ['เลขประจำตัวประชาชน', item.idCard],
          ['คณะ', item.faculty],
          ['สาขาวิชา', item.major],
          ['ชั้นปี', item.year],
          ['GPAX ที่ผู้สมัครกรอก', item.gpax],
          ['GPAX ที่ตรวจพบ', item.verifiedGpax || '-'],
          ['ผลตรวจคุณสมบัติเกรดเฉลี่ย', item.qualificationResult === 'pass' ? 'ผ่านคุณสมบัติ' : item.qualificationResult === 'fail' ? 'ไม่ผ่านคุณสมบัติ' : 'ยังไม่ตรวจสอบ'],
          ['สาเหตุที่ไม่ผ่าน', item.qualificationReason || '-'],
          ['เบอร์โทรศัพท์', item.phone],
          ['อีเมล', item.email],
          ['Line ID', item.lineId],
          ['กลุ่มงาน', item.department],
          ['หน่วยคัดเลือก', item.selectionUnit || item.department],
          ['ตำแหน่งงาน', item.job],
          ['สถานะ', item.status],
          ['ภาษาอังกฤษ', item.englishLevel],
          ['ภาษาไทย', item.thaiLevel],
          ['Microsoft Word', item.wordLevel],
          ['Microsoft Excel', item.excelLevel],
          ['Microsoft PowerPoint', item.powerPointLevel],
          ['Microsoft Access', item.accessLevel],
          ['ประเภทงานที่สนใจ', item.interestTypes],
          ['ทักษะ / ประสบการณ์', item.skills, true],
          ['เหตุผลที่สมัคร', item.reason, true]
        ];

        const files = (item.attachmentFiles || []).length
          ? item.attachmentFiles.map(file => `
              <a class="attachment-chip" href="${escapeHtml(file.url)}" target="_blank" rel="noopener noreferrer">
                <i class="fa-solid fa-paperclip"></i>${escapeHtml(file.label)}
              </a>
            `).join('')
          : `<span class="text-muted">ไม่ได้แนบไฟล์</span>`;

        document.getElementById('detailModalBody').innerHTML = `
          <div class="profile-detail-grid">
            <div>
              <div class="detail-list">
                ${details.map(([label, value, full]) => `
                  <div class="detail-item ${full ? 'full' : ''}">
                    <div class="detail-label">${escapeHtml(label)}</div>
                    <div class="detail-value">${escapeHtml(value || '-').replace(/\n/g, '<br>')}</div>
                  </div>
                `).join('')}
                <div class="detail-item full">
                  <div class="detail-label">ไฟล์แนบ</div>
                  <div class="detail-value">${files}</div>
                </div>
              </div>
            </div>
            <div class="text-center">
              <div id="detailProfilePhotoFast">
                ${item.hasPhoto
                  ? `<div class="profile-detail-photo d-flex flex-column gap-2 align-items-center justify-content-center text-muted"><i class="fa-solid fa-spinner fa-spin fa-2x"></i><span class="small">กำลังโหลดรูป...</span></div>`
                  : `<div class="profile-detail-photo d-flex align-items-center justify-content-center text-muted"><i class="fa-regular fa-user fa-2x"></i></div>`
                }
              </div>
              <div class="small text-muted mt-2">รูปถ่ายผู้สมัคร</div>
            </div>
          </div>
        `;
        detailModal.show();
        if (item && item.hasPhoto) loadDetailProfilePhotoFast(applicationId);
      } catch (error) {
        closeLoading();
        Swal.fire({ icon: 'error', title: 'ไม่สามารถเปิดข้อมูลได้', text: error.message });
      }
    }

    const CLIENT_IDLE_TIMEOUT_MS = 10 * 60 * 1000;
    let clientIdleTimer = null;

    function resetClientIdleTimer() {
      if (!state.token) return;

      clearTimeout(clientIdleTimer);
      clientIdleTimer = setTimeout(() => {
        forceLogout('ไม่มีการใช้งานระบบเกิน 10 นาที กรุณาเข้าสู่ระบบใหม่');
      }, CLIENT_IDLE_TIMEOUT_MS);
    }

    ['click', 'keydown', 'touchstart', 'scroll'].forEach(eventName => {
      window.addEventListener(eventName, resetClientIdleTimer, { passive: true });
    });

    document.addEventListener('visibilitychange', () => {
      if (!document.hidden) resetClientIdleTimer();
    });

    async function logout() {
      const confirmResult = await Swal.fire({
        icon: 'question',
        title: 'ยืนยันการออกจากระบบ',
        text: 'คุณต้องการออกจากระบบเจ้าหน้าที่หรือไม่?',
        showCancelButton: true,
        confirmButtonText: '<i class="fa-solid fa-right-from-bracket me-1"></i> ออกจากระบบ',
        cancelButtonText: 'ยกเลิก',
        confirmButtonColor: '#dc3545',
        cancelButtonColor: '#6c757d',
        reverseButtons: true,
        focusCancel: true,
        allowOutsideClick: false
      });

      if (!confirmResult.isConfirmed) return;

      clearTimeout(clientIdleTimer);

      Swal.fire({
        title: 'กำลังออกจากระบบ',
        text: 'กรุณารอสักครู่',
        allowOutsideClick: false,
        allowEscapeKey: false,
        didOpen: () => Swal.showLoading()
      });

      try {
        if (state.token) await serverCall('staffLogout', state.token);
      } catch (error) {
        console.warn('staffLogout:', error);
      } finally {
        sessionStorage.removeItem('ubuStaffSession');
        sessionStorage.removeItem('ubuStaffView');
        sessionStorage.removeItem('ubuStaffPreloadReadyAt');
        clearTimeout(staffMenuPreloadRefreshTimer);
        staffMenuPreloadQueued = false;

        applySession({});
        state.currentView = '';

        const loginForm = document.getElementById('loginForm');
        if (loginForm) loginForm.reset();

        if (window.location.hash) {
          history.replaceState(null, '', window.location.pathname + window.location.search);
        }

        Swal.close();
        showLoginView();

        const usernameInput = document.getElementById('username');
        if (usernameInput) setTimeout(() => usernameInput.focus(), 50);
      }
    }

    async function forceLogout(message) {
      clearTimeout(clientIdleTimer);
      sessionStorage.removeItem('ubuStaffSession');
      sessionStorage.removeItem('ubuStaffView');
      sessionStorage.removeItem('ubuStaffPreloadReadyAt');
      clearTimeout(staffMenuPreloadRefreshTimer);
      staffMenuPreloadQueued = false;
      applySession({});

      await Swal.fire({
        icon: 'warning',
        title: 'ไม่ได้ทำรายการในเวลาที่กำหนด',
        text: message || 'กรุณาเข้าสู่ระบบเจ้าหน้าที่ใหม่เพื่อใช้งานต่อ',
        confirmButtonText: 'เข้าสู่ระบบใหม่',
        confirmButtonColor: '#00346f',
        allowOutsideClick: false
      });

      const loginForm = document.getElementById('loginForm');
      if (loginForm) loginForm.reset();
      showLoginView();
    }

    function restoreSession() {
      const raw = sessionStorage.getItem('ubuStaffSession');
      if (!raw) {
        showLoginView();
        return;
      }

      try {
        const saved = JSON.parse(raw);
        if (!saved || !saved.token || !saved.role) {
          sessionStorage.removeItem('ubuStaffSession');
          showLoginView();
          return;
        }

        applySession(saved);
        enterApp();
      } catch (_) {
        sessionStorage.removeItem('ubuStaffSession');
        sessionStorage.removeItem('ubuStaffView');
        applySession({});
        showLoginView();
      }
    }

    window.addEventListener('pagehide', () => {
      const modalBody = document.getElementById('detailModalBody');
      if (modalBody) modalBody.replaceChildren();

      state.qualificationRows = [];
      state.qualificationDrafts.clear();
      state.adminApplicantRows = [];
      state.civilRegistryRows = [];
      state.departmentRows = [];
      state.publicContentRows = [];
    });

    function takeSsoTicketFromUrl() {
      try {
        const rawHash = String(window.location.hash || '').replace(/^#/, '');
        if (!rawHash) return '';

        const params = new URLSearchParams(rawHash);
        const ticket = String(params.get('sso') || '').trim();

        if (ticket) {
          // ลบ Ticket ออกจาก Address Bar ทันที
          // Ticket เป็น one-time อยู่แล้ว แต่ไม่ควรค้างใน URL / history
          history.replaceState(
            null,
            '',
            window.location.pathname + window.location.search
          );
        }

        return ticket;
      } catch (_) {
        return '';
      }
    }

    async function bootstrapStaffAuthentication() {
      const ssoTicket = takeSsoTicketFromUrl();

      if (ssoTicket) {
        // SSO จาก Portal ต้องแทน session เดิมเสมอ เพื่อป้องกันสับสนข้ามบัญชี
        sessionStorage.removeItem('ubuStaffSession');
        sessionStorage.removeItem('ubuStaffView');
        applySession({});

        try {
          const result = await serverCall('staffSsoLogin', ssoTicket);

          if (!result || !result.success) {
            throw new Error(
              result && result.message
                ? result.message
                : 'Single Sign-On ไม่สำเร็จ'
            );
          }

          saveStaffSessionFromLoginResult(result);
          sessionStorage.removeItem('ubuStaffView');
          enterApp();
          return;
        } catch (error) {
          showLoginView();
          Swal.fire({
            icon: 'error',
            title: 'Single Sign-On ไม่สำเร็จ',
            text: error && error.message
              ? error.message
              : 'กรุณากลับไปที่ S-MIS PORTAL แล้วเปิดระบบใหม่อีกครั้ง',
            confirmButtonText: 'ตกลง'
          });
          return;
        }
      }

      if (!sessionStorage.getItem('ubuStaffSession')) {
        // เปิด dashboard โดยตรง: คง Login เดิมไว้เป็น fallback
        startStaffLoginWarmup();
      }

      restoreSession();
    }

    bootstrapStaffAuthentication();

    async function loadDetailProfilePhotoFast(applicationId) {
      const target = document.getElementById('detailProfilePhotoFast');
      if (!target) return;
      try {
        let data = applicantPhotoCache.get(applicationId) || '';
        if (!data) {
          const photos = await serverCall('getStaffApplicantPhotoBatch', state.token, [applicationId]);
          data = photos && photos[applicationId];
          if (data) rememberApplicantPhoto(applicationId, data);
        }
        if (!document.body.contains(target)) return;
        target.innerHTML = data
          ? `<img class="profile-detail-photo" src="${data}" alt="รูปผู้สมัคร">`
          : `<div class="profile-detail-placeholder"><i class="fa-regular fa-user fa-2x"></i><div>ไม่มีรูปภาพ</div></div>`;
      } catch (_) {
        if (document.body.contains(target)) {
          target.innerHTML = `<div class="profile-detail-placeholder"><i class="fa-regular fa-image fa-2x"></i><div>โหลดรูปไม่สำเร็จ</div></div>`;
        }
      }
    }

