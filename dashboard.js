      const BUILD_VERSION = '2026.09.05-SMIS-AUTH-V1';
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

    function serverCall(method, ...args) {
      return new Promise((resolve, reject) => {
        try {
          const runner = google.script.run
            .withSuccessHandler(resolve)
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

        applySession({
          token: result.token,
          role: result.role,
          department: result.department || '',
          selectionUnit: result.selectionUnit || result.department || '',
          displayName: result.displayName || '',
          username: result.username || document.getElementById('username').value.trim()
        });

        sessionStorage.setItem('ubuStaffSession', JSON.stringify({
          token: state.token,
          role: state.role,
          department: state.department,
          selectionUnit: state.selectionUnit,
          displayName: state.displayName,
          username: state.username
        }));

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
      const adminViews = ['adminDashboard', 'staffAccounts', 'contentManagement', 'applicantReview', 'civilRegistry', 'qualification', 'forwarding', 'adminResults'];
      const departmentViews = ['departmentDashboard', 'departmentSelection', 'departmentBasket'];

      const targetView = state.role === 'admin'
        ? (adminViews.includes(savedView) ? savedView : 'adminDashboard')
        : (departmentViews.includes(savedView) ? savedView : 'departmentDashboard');

      Promise.resolve().then(() => navigate(targetView));
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
        ['adminResults', 'fa-inbox', 'รับข้อมูลส่งกลับจากหน่วยงาน']
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
      if (!rows.length) return `<tr><td colspan="6" class="text-center text-muted py-5">ยังไม่มีข้อมูล</td></tr>`;

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
            <td>${item.active ? '<span class="status-badge badge-pass">แสดง</span>' : '<span class="status-badge badge-muted">ซ่อน</span>'}</td>
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
            <div><h2 class="intel-panel-title"><i class="fa-solid fa-building me-2"></i>สถิติยอดการสมัครแยกตามหน่วยงาน</h2><div class="intel-panel-subtitle">จัดอันดับยอดสมัคร พร้อมสัดส่วน ค้างตรวจ และความคืบหน้าการส่งรายชื่อ</div></div>
            <span class="analysis-badge">${stats.length.toLocaleString('th-TH')} หน่วยงาน</span>
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

    function queueApplicantReviewSearch(){ clearTimeout(applicantReviewSearchTimer); applicantReviewSearchTimer=setTimeout(()=>filterApplicantReviewTable(),280); }
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
    function queueCivilRegistrySearch(){clearTimeout(civilRegistrySearchTimer);civilRegistrySearchTimer=setTimeout(()=>filterCivilRegistryTable(),280);}
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

    async function renderQualification() {
      setHeader('ตรวจสอบคุณสมบัติเกรดเฉลี่ย', 'ดาวน์โหลดรายชื่อไปตรวจสอบ แล้วนำเข้าผลจากระบบภายนอกเพื่อประมวลผลอัตโนมัติ');
      const content = document.getElementById('content');
      content.innerHTML = `<div class="text-center py-5"><div class="spinner-border text-primary"></div></div>`;
      state.qualificationRows = await serverCall('getAdminQualificationApplicants', state.token) || [];

      const total = state.qualificationRows.length;
      const passed = state.qualificationRows.filter(item => item.qualificationResult === 'pass').length;
      const failed = state.qualificationRows.filter(item => item.qualificationResult === 'fail').length;
      const pending = total - passed - failed;

      content.innerHTML = `
        <div class="d-flex flex-wrap justify-content-between gap-3 align-items-end mb-3">
          <div>
            <div class="page-title">ตรวจสอบคุณสมบัติเกรดเฉลี่ย</div>
            <div class="page-subtitle">เกณฑ์ผ่าน GPAX ตั้งแต่ ${QUALIFICATION_THRESHOLD.toFixed(2)} ขึ้นไป · รองรับการปรับสถานะแบบกลุ่มและนำเข้าผลตรวจจากไฟล์</div>
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
          <div class="result-summary-item"><div class="result-summary-value">${total.toLocaleString('th-TH')}</div><div class="result-summary-label">ทั้งหมด</div></div>
          <div class="result-summary-item"><div class="result-summary-value result-pass">${passed.toLocaleString('th-TH')}</div><div class="result-summary-label">ผ่าน</div></div>
          <div class="result-summary-item"><div class="result-summary-value result-fail">${failed.toLocaleString('th-TH')}</div><div class="result-summary-label">ไม่ผ่าน</div></div>
          <div class="result-summary-item"><div class="result-summary-value result-wait">${pending.toLocaleString('th-TH')}</div><div class="result-summary-label">ยังไม่ตรวจสอบ</div></div>
        </div>

        <div class="panel mb-3">
          <div class="panel-header">
            <div>
              <h2 class="panel-title mb-1">ปรับสถานะแบบกลุ่ม</h2>
              <div class="text-muted small">เลือกผู้สมัครหลายคน แล้วกำหนดผลตรวจพร้อมกัน</div>
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
            <h2 class="panel-title">รายชื่อผู้สมัครทั้งหมด</h2>
            <div class="d-flex flex-wrap gap-2 align-items-center">
              <select id="qualificationStatusFilter" class="form-select form-select-sm" style="width:175px;" onchange="filterQualificationTable()">
                <option value="all">ทุกสถานะ</option>
                <option value="pass">ผ่าน</option>
                <option value="fail">ไม่ผ่าน</option>
                <option value="pending">ยังไม่ตรวจสอบ</option>
              </select>
              <input id="qualificationSearch" class="form-control form-control-sm" style="width:min(300px, 70vw);" placeholder="ค้นหารหัสนักศึกษา / ชื่อ / คณะ" oninput="filterQualificationTable()">
            </div>
          </div>
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
              <tbody id="qualificationBody"></tbody>
            </table>
          </div>
        </div>
      `;
      renderQualificationRows(state.qualificationRows);
    }

    function qualificationResultLabel(result) {
      if (result === 'pass') return 'ผ่านคุณสมบัติ';
      if (result === 'fail') return 'ไม่ผ่านคุณสมบัติ';
      return 'ยังไม่ตรวจสอบ';
    }

    function renderQualificationRows(rows) {
      const body = document.getElementById('qualificationBody');
      body.innerHTML = rows.length ? rows.map((item, index) => {
        const locked = !!item.forwarded;
        const result = ['pass', 'fail'].includes(item.qualificationResult) ? item.qualificationResult : 'pending';
        const verified = item.verifiedGpax !== '' && item.verifiedGpax !== null && item.verifiedGpax !== undefined
          ? Number(item.verifiedGpax).toFixed(2)
          : '-';
        return `
          <tr data-search="${escapeHtml(`${item.studentId} ${item.fullName} ${item.faculty} ${item.department}`.toLowerCase())}"
              data-app-id="${escapeHtml(item.applicationId)}"
              data-qualification-status="${escapeHtml(result)}">
            <td class="text-center">
              <input class="form-check-input qual-select" type="checkbox" ${locked ? 'disabled' : ''}>
            </td>
            <td>${index + 1}</td>
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
              <input class="form-control form-control-sm qual-reason" value="${escapeHtml(item.qualificationReason || '')}" ${locked ? 'disabled' : ''} placeholder="ระบุสาเหตุกรณีไม่ผ่าน">
            </td>
            <td class="text-center">
              <div class="form-check form-switch d-inline-block">
                <input class="form-check-input qual-visible" type="checkbox" ${item.qualificationVisible ? 'checked' : ''} ${locked ? 'disabled' : ''}>
              </div>
            </td>
            <td>${statusBadge(item.status)}</td>
          </tr>`;
      }).join('') : `<tr><td colspan="12" class="text-center text-muted py-4">ยังไม่มีผู้สมัคร</td></tr>`;
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
      filterQualificationTable();
    }

    function toggleQualificationSelectAll(checkbox) {
      document.querySelectorAll('#qualificationBody tr[data-app-id]').forEach(row => {
        if (row.style.display === 'none') return;
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
        Swal.fire({ icon: 'info', title: 'ยังไม่ได้เลือกรายชื่อ', text: 'กรุณาเลือกรายชื่ออย่างน้อย 1 คน' });
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
      const search = (document.getElementById('qualificationSearch')?.value || '').trim().toLowerCase();
      const status = document.getElementById('qualificationStatusFilter')?.value || 'all';
      document.querySelectorAll('#qualificationBody tr[data-search]').forEach(row => {
        const matchSearch = row.dataset.search.includes(search);
        const matchStatus = status === 'all' || row.dataset.qualificationStatus === status;
        row.style.display = matchSearch && matchStatus ? '' : 'none';
      });
      const selectAll = document.getElementById('qualificationSelectAll');
      if (selectAll) selectAll.checked = false;
    }

    async function saveQualification() {
      const updates = [...document.querySelectorAll('#qualificationBody tr[data-app-id]')].map(row => {
        const result = row.querySelector('.qual-result').value || 'pending';
        let reason = row.querySelector('.qual-reason').value.trim();
        const visible = row.querySelector('.qual-visible').checked;
        if (result === 'fail' && !reason) reason = QUALIFICATION_FAIL_REASON;
        if (result !== 'fail') reason = '';
        return { applicationId: row.dataset.appId, result, reason, visible };
      });

      const confirm = await Swal.fire({
        icon: 'question',
        title: 'บันทึกผลการตรวจสอบ',
        text: 'ผลที่เปิดให้แสดงจะปรากฏในหน้าตรวจสอบของนักศึกษา',
        showCancelButton: true,
        confirmButtonText: 'บันทึก',
        cancelButtonText: 'ยกเลิก',
        confirmButtonColor: '#165f9c'
      });
      if (!confirm.isConfirmed) return;

      showLoading('กำลังบันทึกผล');
      try {
        const result = await serverCall('saveQualificationReviews', state.token, updates);
        closeLoading();
        await Swal.fire({ icon: 'success', title: 'บันทึกเรียบร้อย', text: `อัปเดต ${result.updated || 0} รายการ` });
        await renderQualification();
      } catch (error) {
        closeLoading();
        Swal.fire({ icon: 'error', title: 'บันทึกไม่สำเร็จ', text: error.message });
      }
    }

    async function exportQualificationCheckFile() {
      const header = ['ประทับเวลา', 'คำนำหน้า', 'ชื่อ (ไม่ต้องมีคำนำหน้า)', 'นามสกุล', 'เลขบัตรประจำตัวประชาชน', 'รหัสนักศึกษา', 'สังกัดคณะ'];
      const rows = state.qualificationRows.map(item => [
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
        downloadCsv('รายชื่อสำหรับตรวจสอบคุณสมบัติ.csv', [header, ...rows]);
        return;
      }

      const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
      ws['!cols'] = [
        { wch: 22 }, { wch: 15 }, { wch: 28 }, { wch: 28 }, { wch: 22 }, { wch: 18 }, { wch: 35 }
      ];

      // ระบบตรวจสอบภายนอกใช้ตัวอ่าน Excel รุ่นเก่า
      // จึงต้องสร้างไฟล์ .xls แบบ BIFF8 จริง ไม่ใช่ .xlsx แล้วเปลี่ยนนามสกุล
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
      setHeader('ส่งรายชื่อให้หน่วยงานคัดเลือก', 'ส่งเฉพาะรายชื่อที่ผ่านคุณสมบัติ');
      const content=document.getElementById('content');content.innerHTML=`<div class="dashboard-skeleton"><div class="spinner-border text-primary"></div><div>กำลังโหลดข้อมูลการส่งรายชื่อ</div></div>`;
      const groups=await serverCall('getAdminDepartmentForwarding',state.token);window.forwardingGroups=Array.isArray(groups)?groups:[];
      const total=window.forwardingGroups.reduce((n,g)=>n+Number(g.total||0),0);const passed=window.forwardingGroups.reduce((n,g)=>n+Number(g.pass||0),0);const sent=window.forwardingGroups.reduce((n,g)=>n+Number(g.forwarded||0),0);const waiting=window.forwardingGroups.reduce((n,g)=>n+Number(g.unforwardedPass||0),0);
      content.innerHTML=`
        <div class="d-flex flex-wrap justify-content-between gap-3 align-items-end mb-4"><div><div class="page-title">ส่งรายชื่อให้หน่วยงานคัดเลือก</div><div class="page-subtitle">ตรวจสอบจำนวนผู้สมัคร เปิดดูรายชื่อ และดาวน์โหลดไฟล์แยกตามหน่วยงาน</div></div></div>
        <div class="row g-3 mb-4">${watermarkStatCard('fa-users',total,'ผู้สมัครในกลุ่มงาน','yellow')}${watermarkStatCard('fa-user-check',passed,'ผ่านคุณสมบัติ','green')}${watermarkStatCard('fa-paper-plane',sent,'ส่งแล้ว','purple')}${watermarkStatCard('fa-clock',waiting,'รอส่ง','orange')}</div>
        <div class="panel premium-panel"><div class="panel-header"><div><h2 class="panel-title">รายชื่อแยกตามหน่วยงาน</h2><div class="text-muted small mt-1">คลิกจำนวนผู้สมัครเพื่อเปิดรายชื่อ · ดาวน์โหลด Excel ได้รายหน่วยงาน</div></div><span class="analysis-badge">${window.forwardingGroups.length.toLocaleString('th-TH')} หน่วยงาน</span></div><div class="table-wrap"><table class="table table-hover align-middle"><thead><tr><th>หน่วยงาน / งานย่อย</th><th class="text-center">จำนวนผู้สมัคร</th><th class="text-center">ผ่านคุณสมบัติ</th><th class="text-center">ไม่ผ่านคุณสมบัติ</th><th class="text-center">สถานะการส่ง</th><th class="text-end">ดำเนินการ</th></tr></thead><tbody>${window.forwardingGroups.length?window.forwardingGroups.map((group,index)=>{const sentAll=group.pass>0&&group.unforwardedPass===0;const partial=group.forwarded>0&&group.unforwardedPass>0;return `<tr><td><div class="fw-semibold">${escapeHtml(group.label||group.selectionUnit||group.department)}</div><div class="text-muted table-subtext">${Number(group.forwarded||0).toLocaleString('th-TH')} ส่งแล้ว · ${Number(group.unforwardedPass||0).toLocaleString('th-TH')} รอส่ง</div></td><td class="text-center"><button class="count-link" onclick="showForwardingAllApplicants(${index})">${Number(group.total||0).toLocaleString('th-TH')} คน <i class="fa-solid fa-chevron-right ms-1"></i></button></td><td class="text-center"><button class="count-link success" onclick="showGroupList('forward',${index},'pass')">${Number(group.pass||0).toLocaleString('th-TH')} คน</button></td><td class="text-center"><button class="count-link danger" onclick="showGroupList('forward',${index},'fail')">${Number(group.fail||0).toLocaleString('th-TH')} คน</button></td><td class="text-center">${sentAll?`<span class="status-badge badge-pass">ส่งแล้ว</span>`:partial?`<span class="status-badge badge-info">ส่งแล้วบางส่วน</span>`:`<span class="status-badge badge-wait">ยังไม่ส่ง</span>`}</td><td class="text-end text-nowrap"><button class="btn btn-sm btn-outline-success me-1" onclick="downloadForwardingGroup(${index})"><i class="fa-solid fa-file-excel me-1"></i>Excel</button><button class="btn btn-sm ${group.unforwardedPass>0?'btn-primary':'btn-outline-secondary'}" ${group.unforwardedPass>0?'':'disabled'} onclick="sendToDepartment(${index})"><i class="fa-solid fa-paper-plane me-1"></i>${partial?`ส่งเพิ่ม ${group.unforwardedPass} คน`:group.unforwardedPass>0?'ส่งรายชื่อ':'ส่งแล้ว'}</button></td></tr>`;}).join(''):`<tr><td colspan="6" class="text-center text-muted py-5">ยังไม่มีข้อมูล</td></tr>`}</tbody></table></div></div>`;
    }

    function showForwardingAllApplicants(index) {
      const group=(window.forwardingGroups||[])[index]; if(!group) return;
      const rows=[...(group.passApplicants||[]),...(group.failApplicants||[])];
      listModalTitle.textContent=`รายชื่อผู้สมัคร · ${group.label||group.selectionUnit||group.department}`;
      listModalBody.innerHTML=rows.length?`<div class="table-responsive"><table class="table table-sm"><thead><tr><th>รหัสนักศึกษา</th><th>ชื่อ-สกุล</th><th>คณะ</th><th>ผลคุณสมบัติ</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${escapeHtml(x.studentId||'-')}</td><td>${escapeHtml(x.fullName||'-')}</td><td>${escapeHtml(x.faculty||'-')}</td><td>${x.reason?'<span class="status-badge badge-fail">ไม่ผ่าน</span>':'<span class="status-badge badge-pass">ผ่าน</span>'}</td></tr>`).join('')}</tbody></table></div>`:'<div class="text-muted">ไม่มีข้อมูล</div>';
      listModal.show();
    }
    async function downloadForwardingGroup(index){
      const group=(window.forwardingGroups||[])[index]; if(!group)return; const rows=[...(group.passApplicants||[]),...(group.failApplicants||[])];
      const aoa=[['รหัสนักศึกษา','ชื่อ-สกุล','คณะ','หน่วยงาน','ผลคุณสมบัติ'],...rows.map(x=>[String(x.studentId||''),x.fullName||'',x.faculty||'',group.label||group.selectionUnit||group.department||'',x.reason?'ไม่ผ่าน':'ผ่าน'])];
      const safe=(group.label||group.selectionUnit||group.department||'หน่วยงาน').replace(/[\\/:*?"<>|]/g,'_');
      try {
        await ensureXlsxLoaded();
        const ws=XLSX.utils.aoa_to_sheet(aoa); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'รายชื่อ'); XLSX.writeFile(wb,`รายชื่อ_${safe}.xlsx`);
      } catch (_) {
        downloadCsv(`รายชื่อ_${safe}.csv`,aoa);
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
    function showGroupList(mode, index, type) {
      const source = mode === 'forward' ? window.forwardingGroups : window.resultGroups;
      const group = source[index];
      let list = [];
      let title = '';

      if (mode === 'forward') {
        list = type === 'pass' ? group.passApplicants : group.failApplicants;
        title = `${group.label || group.selectionUnit || group.department} · ${type === 'pass' ? 'ผ่านคุณสมบัติ' : 'ไม่ผ่านคุณสมบัติ'}`;
      } else {
        if (type === 'selected') list = group.selectedApplicants;
        if (type === 'rejected') list = group.rejectedApplicants;
        if (type === 'pending') list = group.pendingApplicants;
        title = `${group.label || group.selectionUnit || group.department} · ${type === 'selected' ? 'พิจารณารับ' : type === 'rejected' ? 'ไม่พิจารณารับ' : 'รอผลการพิจารณา'}`;
      }

      document.getElementById('listModalTitle').textContent = title.replace(/\s+/g, ' ').trim();
      document.getElementById('listModalBody').innerHTML = list && list.length ? `
        <div class="table-responsive">
          <table class="table table-sm">
            <thead>
              <tr><th>รหัสนักศึกษา</th><th>ชื่อ-สกุล</th><th>คณะ</th><th>ตำแหน่งงาน</th></tr>
            </thead>
            <tbody>
              ${list.map(item => `
                <tr>
                  <td>${escapeHtml(item.studentId)}</td>
                  <td>${escapeHtml(item.fullName)}</td>
                  <td>${escapeHtml(item.faculty)}</td>
                  <td>${escapeHtml(item.job || '-')}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      ` : `<div class="text-center text-muted py-4">ไม่มีรายชื่อ</div>`;
      listModal.show();
    }

    function buildDepartmentResultRows(group) {
      const unit = group.label || group.selectionUnit || group.department || '';
      const rows = [];
      (group.selectedApplicants || []).forEach(item => rows.push([unit, item.studentId || '', item.fullName || '', item.faculty || '', item.job || '', 'พิจารณารับ', item.status || '']));
      (group.rejectedApplicants || []).forEach(item => rows.push([unit, item.studentId || '', item.fullName || '', item.faculty || '', item.job || '', 'ไม่พิจารณารับ', item.status || '']));
      (group.pendingApplicants || []).forEach(item => rows.push([unit, item.studentId || '', item.fullName || '', item.faculty || '', item.job || '', 'รอผลการพิจารณา', item.status || '']));
      return rows;
    }

    function safeFilenamePart(value) {
      return String(value || 'รายชื่อ').replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 90) || 'รายชื่อ';
    }

    function downloadDepartmentResult(index) {
      const group = (window.resultGroups || [])[index];
      if (!group) return;
      const rows = buildDepartmentResultRows(group);
      if (!rows.length) {
        Swal.fire({ icon: 'info', title: 'ยังไม่มีรายชื่อ', text: 'หน่วยงานนี้ยังไม่มีข้อมูลสำหรับดาวน์โหลด' });
        return;
      }
      const header = ['หน่วยงาน / งานย่อย', 'รหัสนักศึกษา', 'ชื่อ-สกุล', 'คณะ', 'ตำแหน่งงาน', 'ผลการพิจารณา', 'สถานะ'];
      const name = safeFilenamePart(group.label || group.selectionUnit || group.department);
      downloadCsv(`ผลการคัดเลือก_${name}.csv`, [header, ...rows]);
    }

    function downloadAllDepartmentResults() {
      const groups = window.resultGroups || [];
      const rows = groups.flatMap(buildDepartmentResultRows);
      if (!rows.length) {
        Swal.fire({ icon: 'info', title: 'ยังไม่มีข้อมูล', text: 'ยังไม่มีรายชื่อจากหน่วยงานสำหรับดาวน์โหลด' });
        return;
      }
      const header = ['หน่วยงาน / งานย่อย', 'รหัสนักศึกษา', 'ชื่อ-สกุล', 'คณะ', 'ตำแหน่งงาน', 'ผลการพิจารณา', 'สถานะ'];
      downloadCsv('ผลการคัดเลือกจากหน่วยงาน_ทั้งหมด.csv', [header, ...rows]);
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
              <h2 class="panel-title">ผลการพิจารณาแยกตามหน่วยงาน</h2>
              <div class="small text-muted mt-1">${groups.length} หน่วยงาน / งานย่อย</div>
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
                  <th>หน่วยงาน / งานย่อย</th>
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
                        ${group.department && group.selectionUnit && group.selectionUnit !== group.department ? `<div class="small text-muted mt-1">${escapeHtml(group.department)}</div>` : ''}
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
      // FAST V5: Drive images are low priority and never compete with current menu data.
      const generation = Number(state.photoLoadGeneration || 0);
      const ids = rows.filter(item => item.hasPhoto).map(item => item.applicationId);
      if (!ids.length) return;
      await new Promise(resolve => setTimeout(resolve, 700));
      if (generation !== Number(state.photoLoadGeneration || 0) || state.currentView !== 'departmentSelection') return;

      const chunks=[];
      for (let i=0;i<ids.length;i+=6) chunks.push(ids.slice(i,i+6));
      for (const chunk of chunks) {
        if (generation !== Number(state.photoLoadGeneration || 0) || state.currentView !== 'departmentSelection') return;
        try {
          const photos=await serverCall('getStaffApplicantPhotoBatch',state.token,chunk);
          if (generation !== Number(state.photoLoadGeneration || 0)) return;
          chunk.forEach(id=>{
            const data=photos&&photos[id];
            const wrapper=document.getElementById(`photo-wrap-${id}`);
            if (!wrapper||!data) return;
            wrapper.outerHTML=`<img id="photo-wrap-${escapeHtml(id)}" class="photo-thumb" src="${data}" alt="รูปผู้สมัคร">`;
          });
        } catch(error) { console.warn('ไม่สามารถโหลดรูปผู้สมัครบางรายการได้',error); }
        await new Promise(resolve=>setTimeout(resolve,120));
      }
    }

    async function toggleBasket(applicationId, checkbox) {
      const row = checkbox.closest('tr');
      const previous = !checkbox.checked;
      checkbox.disabled = true;

      try {
        await serverCall('setDepartmentBasket', state.token, applicationId, checkbox.checked);
        const item = (state.departmentRows || []).find(record => record.applicationId === applicationId);
        if (item) item.inBasket = checkbox.checked;

        if (row) {
          row.classList.toggle('selected-row', checkbox.checked);
          const statusCell = row.querySelector('.row-status');
          if (statusCell) {
            statusCell.innerHTML = statusBadge(checkbox.checked ? 'อยู่ในตะกร้า' : 'รอพิจารณา');
          }
        }
        updateBasketIndicator();
      } catch (error) {
        checkbox.checked = previous;
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
          ['หน่วยงาน/งานย่อย', item.selectionUnit || item.department],
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
      state.adminApplicantRows = [];
      state.civilRegistryRows = [];
      state.departmentRows = [];
      state.publicContentRows = [];
    });

    if (!sessionStorage.getItem('ubuStaffSession')) {
      // เริ่มปลุก Apps Script + preload Users cache ทันทีที่หน้า Login เปิด
      // เวลาผู้ใช้พิมพ์ชื่อ/รหัสผ่าน ระบบฝั่ง Server จะพร้อมรับ Login แล้ว
      startStaffLoginWarmup();
    }

    restoreSession();

    async function loadDetailProfilePhotoFast(applicationId) {
      const target = document.getElementById('detailProfilePhotoFast');
      if (!target) return;
      try {
        const photos = await serverCall('getStaffApplicantPhotoBatch', state.token, [applicationId]);
        const data = photos && photos[applicationId];
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

