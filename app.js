const categoryConfig = {
      dev: { title: 'กลุ่มงานพัฒนานักศึกษา' },
      welfare: { title: 'กลุ่มงานสวัสดิการนักศึกษา' },
      alumni: { title: 'กลุ่มงานศิษย์เก่าสัมพันธ์' },
      office: { title: 'หน่วยงานในสำนักงานมหาวิทยาลัย' }
    };

    let jobsData = [];
    let jobsLoaded = false;
    let jobsLoading = false;
    let jobsLoadWaiters = [];

    let selectedJobForApplication = null;
    let currentApplicantType = '';

    const facultyPrograms = {
      "คณะวิทยาศาสตร์": [
            "อาชีวอนามัยและความปลอดภัย วท.บ.สาขาวิชาอาชีวอนามัยและความปลอดภัย",
            "คณิตศาสตร์ วท.บ.สาขาวิชาคณิตศาสตร์ประยุกต์",
            "เคมี วท.บ.สาขาวิชาเคมี",
            "จุลชีววิทยา วท.บ.สาขาวิชาชีววิทยา",
            "ชีววิทยา วท.บ.สาขาวิชาจุลชีววิทยา",
            "ฟิสิกส์ วท.บ.สาขาวิชาฟิสิกส์ชีวการแพทย์",
            "วิทยาศาสตร์สิ่งแวดล้อม วท.บ.สาขาวิชาวิทยาศาสตร์สิ่งแวดล้อม",
            "เทคโนโลยีพอลิเมอร์ วท.บ.สาขาวิชาเทคโนโลยีและการออกแบบผลิตภัณฑ์ยางและพอลิเมอร์",
            "นวัตกรรมเทคโนโลยีวัสดุ วท.บ.สาขาวิชานวัตกรรมเทคโนโลยีวัสดุ",
            "เทคโนโลยีสารสนเทศ วท.บ.สาขาวิชาเทคโนโลยีสารสนเทศและการสื่อสาร (วิชาเอกเทคโนโลยีดิจิทัลและเกม (Digital and Game Technology), วิชาเอกเทคโนโลยีเครือข่ายและความมั่นคงปลอดภัยไซเบอร์ (Network Technology and Cyber Security))",
            "วิทยาการ วิทยาศาสตร์ข้อมูล วท.บ.สาขาวิชาวิทยาการข้อมูลและนวัตกรรมซอฟต์แวร์ (วิชาเอกวิทยาการข้อมูล (Data Sciences), วิชาเอกนวัตกรรมซอฟต์แวร์ (Software Innovations))"
      ],
      "คณะเกษตรศาสตร์": [
            "เทคโนโลยีการอาหาร วท.บ.สาขาวิชาเทคโนโลยีการอาหาร",
            "เกษตรศาสตร์ วท.บ.สาขาวิชาเกษตรศาสตร์ (วิชาเอกพืชไร่, วิชาเอกพืชสวน, วิชาเอกสัตวศาสตร์, วิชาอกประมง",
            "เกษตรศาสตร์ วท.บ.สาขาเกษตรวิศว์อัจฉริยะ"
      ],
      "คณะวิศวกรรมศาสตร์": [
            "วิศวกรรมทั่วไป วศ.บ.วิศวกรรมทั่วไป [ประกอบด้วย สาขาวิชาวิศวกรรมเครื่องกล, สาขาวิชาวิศวกรรม อุตสาหการ,สาขาวิชาวิศวกรรมโยธา, สาขาวิชาวิศวกรรมไฟฟ้า, สาขาวิชาวิศวกรรมเคมีและสาขาวิชาวิศวกรรมสิ่งแวดล้อม (เลือกสาขาวิชาเมื่อขึ้นชั้นปีที่ 2)]",
            "วิศวกรรมเครื่องกล วศ.บ.สาขาวิชาวิศวกรรมเครื่องกล",
            "วิศวกรรม วศ.บ.สาขาวิชาวิศวกรรม",
            "วิศวกรรมอุตสาหการ วศ.บ.อุตสาหการ",
            "วิศวกรรมโยธา วศ.บ.สาขาวิชาวิศวกรรมโยธา",
            "วิศวกรรมไฟฟ้า วศ.บ.สาขาวิชาวิศวกรรมไฟฟ้า",
            "วิศวกรรมเคมี วศ.บ.สาขาวิชาวิศวกรรมเคมี",
            "วิศวกรรมสิ่งแวดล้อม วศ.บ.สาขาวิชาวิศวกรรมเคมีและสาขาวิชาวิศวกรรมสิ่งแวดล้อม",
            "วิศวกรรมเมคคาทรอนิกส์ วศ.บ.สาขาวิชาวิศวกรรมเมคคาทรอนิกส์และระบบอัตโนมัติ",
            "วิศวกรรมปัญญาประดิษฐ์ วศ.บ.สาขาวิชาวิศวกรรมปัญญาประดิษฐ์และการสั่งการ"
      ],
      "คณะศิลปศาสตร์": [
            "การท่องเที่ยวและการโรงแรม ศศ.บ.สาขาวิชาการท่องเที่ยวและการบริการ",
            "นิเทศศาสตร์ นศ.บ.สาขาวิชานิเทศศาสตร์",
            "ประวัติศาสตร์ ศศ.บ.สาขาประวัติศาสตร์",
            "จีน ศศ.บ.สาขาวิชาภาษาจีนและการสื่อสาร",
            "ญี่ปุ่น ศศ.บ.สาขาวิชาภาษาญี่ปุ่นและการสื่อสาร",
            "ไทย ศ.บ.สาขาวิชาภาษาไทยและการสื่อสาร",
            "อังกฤษ ศศ.บ.สาขาวิชาภาษาอังกฤษและการสื่อสาร",
            "อังกฤษ ศศ.บ.สาขาวิชาภาษาอังกฤษเพื่อธุรกิจในยุคดิจิทัล (หลักสูตรภาษาอังกฤษ) (ภาษาต่างประเทศ)",
            "การพัฒนาชุมชน ศศ.บ.สาขาวิชาการพัฒนาสังคม"
      ],
      "คณะเภสัชศาสตร์": [
            "เภสัชกรรมอุตสาหการ ภ.บ.วิชาเอกเภสัชกรรมอุตสาหการ",
            "การบริบาลทางเภสัชกรรม ภ.บ.วิชาเอกการบริบาลทางเภสัชกรรม"
      ],
      "คณะบริหารศาสตร์": [
            "การเงิน บธ.บ.วิชาเอกการเงินและการลงทุน",
            "การจัดการ บธ.บ.วิชาเอกธุรกิจระหว่างประเทศ (ภาคภาษาอังกฤษ) (ภาษาต่างประเทศ)",
            "การจัดการ บธ.บ.วิชาเอกการจัดการธุรกิจ",
            "การจัดการ บธ.บ.สาขาวิชาการจัดการโลจิสติกส์และโซ่อุปทาน",
            "การบัญชี บช.บ.",
            "การตลาด บธ.บ.วิชาเอกการตลาด (กลุ่มการตลาดดิจิทัล)",
            "การตลาด บธ.บ.วิชาเอกการตลาด (กลุ่มการค้าสมัยใหม่ )",
            "การท่องเที่ยวและการโรงแรม บธ.บ.วิชาเอกการจัดการโรงแรมและการบริการ"
      ],
      "วิทยาลัยแพทยศาสตร์และการสาธารณสุข": [
            "แพทยศาสตร์ พ.บ.",
            "อนามัยสิ่งแวดล้อม วท.บ.สาขาวิชาอนามัยสิ่งแวดล้อม",
            "สาธารณสุข ส.บ.สาขาสาธารณสุขศาสตร์"
      ],
      "คณะศิลปประยุกต์และสถาปัตยกรรมศาสตร์": [
            "สถาปัตยกรรม สถ.บ.สาขาวิชาสถาปัตยกรรมศาสตร์",
            "การออกแบบ ศป.บ.สาขาวิชาดิจิทัลอาร์ตและการออกแบบ"
      ],
      "คณะนิติศาสตร์": [
            "นิติศาสตร์ น.บ."
      ],
      "คณะรัฐศาสตร์": [
            "การปกครอง ร.บ.สาขาวิชารัฐศาสตร์ (วิชาเอกการปกครอง, วิชาเอกความสัมพันธ์ระหว่างประเทศ)",
            "รัฐประศาสนศาสตร์ รป.บ.สาขาวิชารัฐประศาสนศาสตร์ (วิชาเอกการบริหารรัฐกิจ, วิชาเอกนวัตกรรมการบริหารท้องถิ่น)"
      ],
      "คณะพยาบาลศาสตร์": [
            "พยาบาลศาสตร์ พย.บ."
      ],
      "คณะศึกษาศาสตร์": [
            "คณิตศาสตร์ ศษ.บ.วิชาเอกคณิตศาสตร์",
            "ภาษาไทยและภาษาต่างประเทศ ศษ.บ.วิชาเอกภาษาไทย",
            "ภาษาไทยและภาษาต่างประเทศ ศษ.บ.วิชาเอกภาษาอังกฤษ",
            "ศิลปศึกษา ดนตรีศึกษาและนาฏศิลป์ศึกษา ศษ.บ.วิชาเอกดนตรีศึกษา",
            "สังคมศึกษา ศษ.บ.วิชาเอกสังคมศึกษา",
            "สุขศึกษา พลศึกษา และนันทนาการ ศษ.บ.วิชาเอกพลศึกษา",
            "วิทยาศาสตร์ทั่วไป ศษ.บ.วิชาเอกวิทยาศาสตร์ทั่วไป"
      ]
};

    const applicantTypeLabels = {
      returning2569: 'นักศึกษาเก่าจ้างงานจากปี 2569',
      new: 'นักศึกษาใหม่ยังไม่เคยสมัครจ้างงาน'
    };

    const interestOptionsByDepartment = {
      'กลุ่มงานพัฒนานักศึกษา': [
        'งานกิจกรรมและพัฒนานักศึกษา',
        'งานกีฬาและนันทนาการ',
        'งานจัดสถานที่และดูแลอุปกรณ์',
        'งานพิธีกร / ดำเนินกิจกรรม',
        'งานประชาสัมพันธ์และสื่อกิจกรรม',
        'งานเอกสารและประสานงาน'
      ],

      'กลุ่มงานสวัสดิการนักศึกษา': [
        'งานเอกสารและตรวจสอบข้อมูล',
        'งานบริการและให้ข้อมูลนักศึกษา',
        'งานประสานงาน',
        'งานระบบข้อมูล / คอมพิวเตอร์',
        'งานประชาสัมพันธ์',
        'งานสนับสนุนกิจกรรมและโครงการ'
      ],

      'กลุ่มงานศิษย์เก่าสัมพันธ์': [
        'งานฐานข้อมูลและข้อมูลศิษย์เก่า',
        'งานประสานงานศิษย์เก่า',
        'งานประชาสัมพันธ์ / Social Media',
        'งานกราฟิก / ครีเอทีฟ',
        'งานกิจกรรมและอีเวนต์',
        'งานเอกสารและธุรการ'
      ],

      'หน่วยงานในสำนักงานมหาวิทยาลัย': [
        'งานเอกสารและธุรการ',
        'งานบริการและประสานงาน',
        'งานข้อมูล / สถิติ / บันทึกข้อมูล',
        'งานไอที / ระบบคอมพิวเตอร์',
        'งานสื่อ / กราฟิก / ประชาสัมพันธ์',
        'งานจัดกิจกรรมและสถานที่'
      ]
    };

    const defaultInterestOptions = [
      'งานเอกสารและธุรการ',
      'งานบริการและประสานงาน',
      'งานข้อมูล / บันทึกข้อมูล',
      'งานไอที / ระบบคอมพิวเตอร์',
      'งานประชาสัมพันธ์และสื่อ',
      'งานกิจกรรมและจัดสถานที่'
    ];

    function getInterestOptionsForJob(job) {
      if (!job || typeof job !== 'object') return defaultInterestOptions;
      const department = String(job.department || '').trim();
      return interestOptionsByDepartment[department] || defaultInterestOptions;
    }

    function renderInterestOptionsForJob(job) {
      const container = document.getElementById('interestTypesGroup');
      const hint = document.getElementById('interestGroupHint');
      if (!container) return;

      const department = String((job && job.department) || '').trim();
      const options = getInterestOptionsForJob(job);

      container.innerHTML = '';

      options.forEach(item => {
        const label = document.createElement('label');
        label.className = 'interest-option';

        const input = document.createElement('input');
        input.className = 'form-check-input me-2 interest-type';
        input.type = 'checkbox';
        input.value = item;

        label.appendChild(input);
        label.appendChild(document.createTextNode(item));
        container.appendChild(label);
      });

      if (hint) {
        hint.innerHTML = department
          ? `ตัวเลือกที่แนะนำสำหรับ <strong>${escapeHtml(department)}</strong>`
          : 'ระบบแสดงประเภทงานที่เหมาะสมกับตำแหน่งที่สมัคร';
      }
    }

    function loadJobsFromSheet(onReady, showError = false) {
      if (typeof onReady === 'function') jobsLoadWaiters.push(onReady);
      if (jobsLoaded) {
        const waiters = jobsLoadWaiters.splice(0);
        waiters.forEach(fn => fn());
        return;
      }
      if (jobsLoading) return;
      jobsLoading = true;

      google.script.run
        .withSuccessHandler((rows) => {
          jobsData = Array.isArray(rows) ? rows : [];
          jobsLoaded = true;
          jobsLoading = false;
          const waiters = jobsLoadWaiters.splice(0);
          waiters.forEach(fn => fn());
        })
        .withFailureHandler((error) => {
          jobsLoading = false;
          jobsLoadWaiters = [];
          console.error('โหลดรายการงานไม่สำเร็จ', error);
          if (showError) {
            Swal.fire({
              icon: 'error',
              title: 'ไม่สามารถโหลดรายการงานได้',
              text: error && error.message ? error.message : 'กรุณาลองใหม่อีกครั้ง',
              confirmButtonColor: '#004a99'
            });
          }
        })
        .getJobs();
    }

    function getJobsForCategory(key) {
      const config = categoryConfig[key];
      if (!config) return [];
      return jobsData
        .filter(job => {
          if (!job || job.active === false) return false;
          const department = String(job.department || '').trim();
          return department === config.title || department.toLowerCase() === key.toLowerCase();
        })
        .sort((a, b) => {
          const orderA = Number(a.sortOrder || 0);
          const orderB = Number(b.sortOrder || 0);
          if (orderA !== orderB) return orderA - orderB;
          return String(a.name || '').localeCompare(String(b.name || ''), 'th');
        });
    }

    function closeMobileMenu() {
      const navbarCollapse = document.getElementById('topNavbar');
      if (navbarCollapse && navbarCollapse.classList.contains('show')) {
        const bsCollapse = bootstrap.Collapse.getInstance(navbarCollapse) || new bootstrap.Collapse(navbarCollapse);
        bsCollapse.hide();
      }
    }

    function hideAllSections() {
      document.getElementById('landing-content').style.display = 'none';
      document.querySelectorAll('.content-section').forEach(sec => sec.style.display = 'none');
    }

    function showLandingPage() {
      closeMobileMenu();
      hideAllSections();
      document.getElementById('landing-content').style.display = 'block';
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function showSection(sectionId) {
      closeMobileMenu();
      hideAllSections();
      const target = document.getElementById(sectionId);
      target.style.display = 'block';
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function showStatusSection() {
      showSection('status-section');
      document.getElementById('statusStepper').style.display = 'none';
      document.getElementById('applicationDetailCard').style.display = 'none';
      resetApplicationProfileImage();
      document.getElementById('applicationDetailGrid').innerHTML = '';
      document.getElementById('searchInput').value = '';
      const verifyId = document.getElementById('statusIdCardInput');
      if (verifyId) verifyId.value = '';
    }

    function publicServerCall(method, ...args) {
      if (window.UBUApi && typeof window.UBUApi.call === 'function') {
        return window.UBUApi.call(method, ...args);
      }

      return new Promise((resolve, reject) => {
        try {
          const runner = google.script.run
            .withSuccessHandler(resolve)
            .withFailureHandler(error => reject(new Error(
              error && error.message ? error.message : 'เกิดข้อผิดพลาดในการเชื่อมต่อระบบ'
            )));
          runner[method](...args);
        } catch (error) {
          reject(error);
        }
      });
    }

    function safePublicFileUrl(value) {
      try {
        const url = new URL(String(value || ''));
        const host = url.hostname.toLowerCase();
        if (url.protocol !== 'https:') return '';
        if (host === 'drive.google.com' || host === 'docs.google.com' || host.endsWith('.googleusercontent.com')) {
          return url.href;
        }
      } catch (_) {}
      return '';
    }

    function renderPublicContentItems(rows, type) {
      const data = Array.isArray(rows) ? rows : [];
      if (!data.length) {
        const emptyText = type === 'news'
          ? 'ยังไม่มีข่าวประกาศในขณะนี้'
          : type === 'rules'
            ? 'ยังไม่มีประกาศระเบียบการรับสมัคร'
            : 'ยังไม่มีไฟล์คู่มือการสมัคร';
        return `<div class="text-center text-muted py-5"><i class="fa-regular fa-folder-open fa-2x mb-3 d-block"></i>${emptyText}</div>`;
      }

      return data.map(item => {
        const fileUrl = safePublicFileUrl(item.fileUrl);
        const fileLabel = type === 'rules'
          ? 'เปิดไฟล์ประกาศ'
          : type === 'manual'
            ? 'เปิดไฟล์คู่มือ'
            : 'เปิดไฟล์แนบ';
        const details = escapeHtml(item.details || '').replace(/\r?\n/g, '<br>');

        return `
          <article class="public-content-card">
            <div class="public-content-date"><i class="fa-regular fa-calendar me-1"></i>วันที่ประกาศ ${escapeHtml(item.publishDate || '-')}</div>
            <div class="public-content-subject">${escapeHtml(item.subject || '-')}</div>
            ${details ? `<div class="public-content-details">${details}</div>` : ''}
            ${fileUrl ? `
              <div class="public-content-file">
                <a class="btn btn-outline-primary btn-sm" href="${escapeHtml(fileUrl)}" target="_blank" rel="noopener noreferrer">
                  <i class="fa-solid fa-paperclip me-1"></i>${fileLabel}
                </a>
              </div>` : ''}
          </article>`;
      }).join('');
    }

    async function loadPublicContent(type, containerId) {
      const container = document.getElementById(containerId);
      if (!container) return;
      container.innerHTML = '<div class="text-center py-5"><div class="spinner-border text-primary" role="status"></div><div class="text-muted mt-2">กำลังโหลดข้อมูล...</div></div>';

      try {
        const rows = await publicServerCall('getPublicContents', type);
        container.innerHTML = renderPublicContentItems(rows, type);
      } catch (error) {
        container.innerHTML = `<div class="alert alert-danger mb-0"><i class="fa-solid fa-triangle-exclamation me-2"></i>${escapeHtml(error.message || 'ไม่สามารถโหลดข้อมูลได้')}</div>`;
      }
    }

    async function showAnnouncementsSection() {
      showSection('announcements-section');
      await loadPublicContent('news', 'announcementsContentList');
    }

    async function showRulesSection() {
      showSection('rules-section');
      await loadPublicContent('rules', 'rulesContentList');
    }

    async function showManualSection() {
      showSection('manual-section');
      await loadPublicContent('manual', 'manualContentList');
    }

    function searchStatus() {
      const studentEl = document.getElementById('searchInput');
      const idCardEl = document.getElementById('statusIdCardInput');

      const studentId = String(studentEl.value || '').replace(/\D/g, '').slice(0, 11);
      const idCard = String(idCardEl.value || '').replace(/\D/g, '').slice(0, 13);

      studentEl.value = studentId;
      idCardEl.value = idCard;

      if (!/^\d{11}$/.test(studentId) || !isValidThaiId(idCard)) {
        Swal.fire({
          icon: 'warning',
          title: 'ข้อมูลยืนยันไม่ถูกต้อง',
          text: 'กรุณากรอกรหัสนักศึกษา 11 หลักและเลขประจำตัวประชาชน 13 หลักให้ถูกต้อง',
          confirmButtonColor: '#004a99'
        });
        return;
      }

      Swal.fire({
        title: 'กำลังตรวจสอบข้อมูล',
        text: 'กรุณารอสักครู่',
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
      });

      google.script.run
        .withSuccessHandler((result) => {
          Swal.close();

          if (!result || !result.found) {
            document.getElementById('statusStepper').style.display = 'none';
            document.getElementById('applicationDetailCard').style.display = 'none';
            resetApplicationProfileImage();

            Swal.fire({
              icon: 'info',
              title: 'ไม่พบข้อมูลการสมัคร',
              text: result && result.message
                ? result.message
                : 'ข้อมูลรหัสนักศึกษาและเลขประจำตัวประชาชนไม่ตรงกับใบสมัคร',
              confirmButtonColor: '#004a99'
            });
            return;
          }

          idCardEl.value = idCard;
          renderApplicationStatus(result);
        })
        .withFailureHandler((error) => {
          document.getElementById('statusStepper').style.display = 'none';
          document.getElementById('applicationDetailCard').style.display = 'none';
          resetApplicationProfileImage();

          Swal.fire({
            icon: 'error',
            title: 'ไม่สามารถตรวจสอบสถานะได้',
            text: error && error.message
              ? error.message
              : 'เกิดข้อผิดพลาดในการเชื่อมต่อระบบ',
            confirmButtonColor: '#004a99'
          });
        })
        .getApplicationStatus(studentId, idCard);
    }

    function maskThaiIdForDetail(value) {
      const id = String(value || '').replace(/\D/g, '');
      if (id.length !== 13) return '-';
      return `${id.slice(0, 1)}-XXXX-XXXXX-${id.slice(-3)}`;
    }

    function renderApplicationStatus(result) {
      const stepper = document.getElementById('statusStepper');
      const steps = [...stepper.querySelectorAll('.step-item')];
      const status = String(result.status || 'รอตรวจสอบ').trim();
      const normalized = status.toLowerCase();

      steps.forEach((step, index) => {
        step.classList.remove('completed', 'active');
        const counter = step.querySelector('.step-counter');
        counter.innerHTML = String(index + 1);
      });

      steps[0].classList.add('completed');
      steps[0].querySelector('.step-counter').innerHTML = '<i class="fa-solid fa-check"></i>';

      let activeIndex = 1;
      let messageIcon = 'fa-hourglass-half';
      let bg = '#f0f5fa';
      let color = '#004a99';
      let border = '#cce0ff';

      if (normalized.includes('ไม่ผ่าน') || normalized.includes('ปฏิเสธ') || normalized.includes('ยกเลิก')) {
        if (normalized.includes('คัดเลือก')) {
          steps[1].classList.add('completed');
          steps[1].querySelector('.step-counter').innerHTML = '<i class="fa-solid fa-check"></i>';
          steps[2].classList.add('completed');
          steps[2].querySelector('.step-counter').innerHTML = '<i class="fa-solid fa-check"></i>';
          steps[3].classList.add('active');
          steps[3].querySelector('.step-counter').innerHTML = '<i class="fa-solid fa-xmark"></i>';
        } else {
          steps[1].classList.add('active');
          steps[1].querySelector('.step-counter').innerHTML = '<i class="fa-solid fa-xmark"></i>';
        }
        activeIndex = -1;
        messageIcon = 'fa-circle-xmark';
        bg = '#fff5f5'; color = '#b02a37'; border = '#f1aeb5';
      } else if (normalized.includes('ผ่านการคัดเลือก') || normalized === 'อนุมัติ') {
        steps.forEach(step => {
          step.classList.add('completed');
          step.querySelector('.step-counter').innerHTML = '<i class="fa-solid fa-check"></i>';
        });
        activeIndex = -1;
        messageIcon = 'fa-circle-check';
        bg = '#ecfdf3'; color = '#146c43'; border = '#a3cfbb';
      } else if (normalized.includes('รอพิจารณา')) {
        steps[1].classList.add('completed');
        steps[1].querySelector('.step-counter').innerHTML = '<i class="fa-solid fa-check"></i>';
        activeIndex = 2;
      } else if (normalized.includes('ผ่านคุณสมบัติ')) {
        steps[1].classList.add('completed');
        steps[1].querySelector('.step-counter').innerHTML = '<i class="fa-solid fa-check"></i>';
        activeIndex = 2;
        messageIcon = 'fa-circle-check';
        bg = '#ecfdf3'; color = '#146c43'; border = '#a3cfbb';
      } else {
        activeIndex = 1;
      }

      if (activeIndex >= 0) steps[activeIndex].classList.add('active');

      const statusMessage = document.getElementById('statusMessage');
      statusMessage.style.backgroundColor = bg;
      statusMessage.style.color = color;
      statusMessage.style.borderColor = border;
      statusMessage.innerHTML = `
        <div><i class="fa-solid ${messageIcon} me-1"></i> สถานะปัจจุบัน: <strong>${escapeHtml(status)}</strong></div>
        <div class="mt-2" style="font-size:13px; font-weight:400;">
          ${escapeHtml(result.fullName || '')} · ${escapeHtml(result.job || '')}
          ${result.submittedAt ? ' · สมัครเมื่อ ' + escapeHtml(result.submittedAt) : ''}
        </div>
      `;

      renderApplicationDetails(result);
      stepper.style.display = 'block';
    }

    function renderAttachmentLinks(files, fallbackSummary) {
      const items = Array.isArray(files) ? files : [];

      if (!items.length) {
        return escapeHtml(fallbackSummary || 'ไม่ได้แนบไฟล์');
      }

      const links = items.map((file, index) => {
        const rawUrl = String((file && file.url) || '').trim();
        let safeUrl = '';

        try {
          const parsed = new URL(rawUrl);
          if (
            parsed.protocol === 'https:' &&
            (parsed.hostname === 'drive.google.com' ||
             parsed.hostname.endsWith('.google.com') ||
             parsed.hostname.endsWith('.googleusercontent.com'))
          ) {
            safeUrl = parsed.href;
          }
        } catch (_) {}

        if (!safeUrl) return '';

        const label = `ไฟล์แนบ ${index + 1}`;

        return `
          <a
            class="attachment-link"
            href="${escapeHtml(safeUrl)}"
            target="_blank"
            rel="noopener noreferrer"
            aria-label="เปิดไฟล์แนบ ${index + 1}"
          >
            <i class="fa-solid fa-paperclip"></i>
            ${escapeHtml(label)}
          </a>
        `;
      }).filter(Boolean);

      if (!links.length) {
        return escapeHtml(fallbackSummary || 'ไม่พบลิงก์ไฟล์แนบ');
      }

      return `<div class="attachment-links">${links.join('')}</div>`;
    }

    function renderApplicationDetails(result) {
      const card = document.getElementById('applicationDetailCard');
      const grid = document.getElementById('applicationDetailGrid');

      const items = [
        ['เลขที่ใบสมัคร', result.applicationId ? `<span class="application-id-badge">${escapeHtml(result.applicationId)}</span>` : '-'],
        ['วันที่สมัคร', escapeHtml(result.submittedAt || '-')],
        ['ตำแหน่งงานที่สมัคร', escapeHtml(result.job || '-')],
        ['ประเภทผู้สมัคร', escapeHtml(result.applicantTypeLabel || result.applicantType || '-')],
        ['ชื่อ-นามสกุล', escapeHtml(result.fullName || '-')],
        ['เลขประจำตัวประชาชน', escapeHtml(maskThaiIdForDetail(result.idCard))],
        ['รหัสนักศึกษา', escapeHtml(result.studentId || '-')],
        ['คณะ', escapeHtml(result.faculty || '-')],
        ['สาขาวิชา', escapeHtml(result.major || '-')],
        ['ชั้นปี', escapeHtml(result.year || '-')],
        ['GPAX', escapeHtml(result.gpax || '-')],
        ['เบอร์โทรศัพท์', escapeHtml(result.phone || '-')],
        ['อีเมล', escapeHtml(result.email || '-')],
        ['Line ID', escapeHtml(result.lineId || '-')],
        ['ภาษาอังกฤษ', escapeHtml(result.englishLevel || '-')],
        ['ภาษาไทย', escapeHtml(result.thaiLevel || '-')],
        ['Microsoft Word', escapeHtml(result.wordLevel || '-')],
        ['Microsoft Excel', escapeHtml(result.excelLevel || '-')],
        ['Microsoft PowerPoint', escapeHtml(result.powerPointLevel || '-')],
        ['Microsoft Access', escapeHtml(result.accessLevel || '-')],
        ['ประเภทงานที่สนใจ', escapeHtml(result.interestTypes || '-')],
        ['ไฟล์แนบ', renderAttachmentLinks(result.attachmentFiles, result.attachmentSummary || 'ไม่ได้แนบไฟล์')],
        ...(result.qualificationVisible
          ? [[
              'ผลตรวจสอบคุณสมบัติ',
              result.qualificationResult === 'pass'
                ? '<span class="badge text-bg-success">ผ่านคุณสมบัติ</span>'
                : result.qualificationResult === 'fail'
                  ? '<span class="badge text-bg-danger">ไม่ผ่านคุณสมบัติ</span>'
                  : '<span class="badge text-bg-secondary">รอตรวจสอบ</span>'
            ]]
          : []),
        ...(result.qualificationVisible && result.qualificationResult === 'fail' && result.qualificationReason
          ? [['สาเหตุที่ไม่ผ่าน', escapeHtml(result.qualificationReason), true]]
          : []),
        ['ทักษะความสามารถพิเศษ / ประสบการณ์', escapeHtml(result.skills || '-'), true],
        ['เหตุผลที่สนใจสมัครงานนี้', escapeHtml(result.reason || '-').replace(/\n/g, '<br>'), true]
      ];

      grid.innerHTML = items.map(([label, value, full]) => `
        <div class="application-detail-item${full ? ' full' : ''}">
          <div class="application-detail-label">${escapeHtml(label)}</div>
          <div class="application-detail-value">${value}</div>
        </div>
      `).join('');

      const profileImage = document.getElementById('applicationProfileImage');
      const profilePlaceholder = document.getElementById('applicationProfilePlaceholder');

      if (result.profileImageData) {
        profileImage.src = result.profileImageData;
        profileImage.style.display = 'block';
        profilePlaceholder.style.display = 'none';
      } else {
        profileImage.removeAttribute('src');
        profileImage.style.display = 'none';
        profilePlaceholder.style.display = 'block';
      }

      card.style.display = 'block';
    }

    function resetApplicationProfileImage() {
      const profileImage = document.getElementById('applicationProfileImage');
      const profilePlaceholder = document.getElementById('applicationProfilePlaceholder');

      if (profileImage) {
        profileImage.removeAttribute('src');
        profileImage.style.display = 'none';
      }

      if (profilePlaceholder) profilePlaceholder.style.display = 'block';
    }



    document.getElementById('searchInput').addEventListener('keydown', function (event) {
      if (event.key === 'Enter') {
        event.preventDefault();
        searchStatus();
      }
    });

    function escapeHtml(value) {
      return String(value ?? '')
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#039;');
    }

    let currentCategoryKey = null;
    let currentGroup = null;

    function createJobButton(jobTitle, onClick) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'list-group-item w-100 text-start d-flex align-items-center';
      btn.innerHTML = `
        <div class="icon-box-list"><i class="fa-solid fa-briefcase"></i></div>
        <span class="fw-bold flex-grow-1" style="color: #333; font-size: 15px;">${escapeHtml(jobTitle)}</span>
        <i class="fa-solid fa-chevron-right text-primary"></i>
      `;
      btn.onclick = onClick;
      return btn;
    }

    function setSubCategoryBack(mode) {
      const text = document.getElementById('subCategoryBackText');
      if (!text) return;
      text.innerText = mode === 'group'
        ? 'ย้อนกลับไปเลือกงานหลัก'
        : 'ย้อนกลับไปเลือกกลุ่มงานหลัก';
    }

    function handleSubCategoryBack() {
      if (currentCategoryKey && currentGroup) {
        showSubCategories(currentCategoryKey);
        return;
      }
      showLandingPage();
    }

    function showSubCategories(key) {
      const config = categoryConfig[key];
      if (!config) return;

      currentCategoryKey = key;
      currentGroup = null;
      document.getElementById('subCategoryTitle').innerText = config.title;
      setSubCategoryBack('category');

      const listContainer = document.getElementById('subJobList');
      listContainer.innerHTML = `
        <div class="text-center text-muted py-4">
          <i class="fa-solid fa-spinner fa-spin me-2"></i>กำลังโหลดรายการงาน...
        </div>
      `;
      showSection('sub-categories-section');

      loadJobsFromSheet(() => {
        if (currentCategoryKey === key) renderCategoryJobs(key);
      }, true);
    }

    function renderCategoryJobs(key) {
      const rows = getJobsForCategory(key);
      const listContainer = document.getElementById('subJobList');
      listContainer.innerHTML = '';

      if (!rows.length) {
        listContainer.innerHTML = `
          <div class="text-center text-muted py-4">
            <i class="fa-regular fa-folder-open me-2"></i>ยังไม่มีตำแหน่งงานที่เปิดรับสมัคร
          </div>
        `;
        return;
      }

      const groups = new Map();
      const directJobs = [];

      rows.forEach(job => {
        const parent = String(job.category || '').trim();
        if (parent) {
          if (!groups.has(parent)) groups.set(parent, []);
          groups.get(parent).push(job);
        } else {
          directJobs.push(job);
        }
      });

      groups.forEach((childJobs, groupTitle) => {
        listContainer.appendChild(
          createJobButton(groupTitle, () => showJobGroup(key, groupTitle, childJobs))
        );
      });

      directJobs.forEach(job => {
        listContainer.appendChild(
          createJobButton(job.name, () => showJobDetails(job))
        );
      });
    }

    function showJobGroup(categoryKey, groupTitle, childJobs) {
      currentCategoryKey = categoryKey;
      currentGroup = groupTitle;
      document.getElementById('subCategoryTitle').innerText = groupTitle;
      setSubCategoryBack('group');

      const listContainer = document.getElementById('subJobList');
      listContainer.innerHTML = '';

      childJobs
        .slice()
        .sort((a, b) => Number(a.sortOrder || 0) - Number(b.sortOrder || 0))
        .forEach(job => {
          listContainer.appendChild(
            createJobButton(job.name, () => showJobDetails(job))
          );
        });

      showSection('sub-categories-section');
    }

    function formatJobText(value) {
      const text = String(value || '').trim();
      return text ? escapeHtml(text).replace(/\n/g, '<br>') : '<span class="text-muted">ไม่ระบุ</span>';
    }

    function showJobDetails(job) {
      if (!job || typeof job !== 'object') return;

      const jobName = String(job.name || '').trim();
      document.getElementById('jobModalTitle').innerText = jobName;
      document.getElementById('jobModalDesc').innerHTML = `
        <div class="mb-3">
          <div class="fw-bold mb-1"><i class="fa-solid fa-list-check text-primary me-2"></i>ลักษณะงานเบื้องต้น</div>
          <div class="ms-4 text-muted">${formatJobText(job.description)}</div>
        </div>
        <div class="mb-3">
          <div class="fw-bold mb-1"><i class="fa-solid fa-clock text-warning me-2"></i>เวลาปฏิบัติงาน</div>
          <div class="ms-4 text-muted">${formatJobText(job.workTime)}</div>
        </div>
        <div class="mb-3">
          <div class="fw-bold mb-1"><i class="fa-solid fa-sack-dollar text-success me-2"></i>อัตราค่าตอบแทน</div>
          <div class="ms-4 text-muted">${formatJobText(job.compensation)}</div>
        </div>
      `;

      document.getElementById('btnApplyJob').onclick = function() {
        const myModalEl = document.getElementById('jobModal');
        const modal = bootstrap.Modal.getInstance(myModalEl);
        if (modal) modal.hide();
        setTimeout(() => showForm(job), 300);
      };

      const jobModal = bootstrap.Modal.getOrCreateInstance(document.getElementById('jobModal'));
      jobModal.show();
    }


    function showForm(job) {
      const jobName = typeof job === 'object' ? String(job.name || '') : String(job || '');
      selectedJobForApplication = job;
      document.getElementById('displaySelectedJob').innerText = jobName;
      resetApplicationFlow(jobName);
      renderInterestOptionsForJob(job);
      showSection('apply-section');
    }

    function resetApplicationFlow(jobName) {
      const form = document.getElementById('applicationForm');
      if (form) form.reset();

      currentApplicantType = '';
      document.getElementById('applicantType').value = '';
      document.getElementById('applicantTypeStep').style.display = 'block';
      document.getElementById('applicationFormStep').style.display = 'none';
      document.getElementById('legacyLookupBox').style.display = 'none';
      document.getElementById('legacyIdCard').value = '';
      document.getElementById('applicantTypeBadge').innerText = '';
      document.getElementById('idCard').readOnly = false;
      document.getElementById('major').innerHTML = '<option value="">กรุณาเลือกคณะก่อน...</option>';
      document.getElementById('major').disabled = true;
      document.querySelectorAll('.interest-type').forEach(el => el.checked = false);
      document.querySelectorAll('.reason-option').forEach(el => el.checked = false);
      document.getElementById('selectedFilesList').innerHTML = '';
      document.getElementById('jobSelected').value = jobName || '';
    }

    function chooseApplicantType(type) {
      if (!applicantTypeLabels[type]) return;

      currentApplicantType = type;
      document.getElementById('applicantType').value = type;

      if (type === 'returning2569') {
        document.getElementById('legacyLookupBox').style.display = 'block';
        setTimeout(() => document.getElementById('legacyIdCard').focus(), 100);
        return;
      }

      openApplicationForm('new');
    }

    function openApplicationForm(type) {
      currentApplicantType = type;
      document.getElementById('applicantType').value = type;
      document.getElementById('applicantTypeStep').style.display = 'none';
      document.getElementById('applicationFormStep').style.display = 'block';
      document.getElementById('applicantTypeBadge').innerText = applicantTypeLabels[type] || '';
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function backToApplicantType() {
      const jobName = document.getElementById('jobSelected').value;
      document.getElementById('applicationForm').reset();
      document.querySelectorAll('.interest-type').forEach(el => el.checked = false);
      document.querySelectorAll('.reason-option').forEach(el => el.checked = false);
      document.getElementById('selectedFilesList').innerHTML = '';
      document.getElementById('idCard').readOnly = false;
      document.getElementById('jobSelected').value = jobName;
      currentApplicantType = '';
      document.getElementById('applicantType').value = '';
      document.getElementById('applicationFormStep').style.display = 'none';
      document.getElementById('applicantTypeStep').style.display = 'block';
      document.getElementById('legacyLookupBox').style.display = 'none';
      document.getElementById('legacyIdCard').value = '';
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function isValidThaiId(id) {
      const value = String(id || '').replace(/\D/g, '');
      if (!/^\d{13}$/.test(value)) return false;

      let sum = 0;
      for (let i = 0; i < 12; i++) {
        sum += Number(value.charAt(i)) * (13 - i);
      }
      const checkDigit = (11 - (sum % 11)) % 10;
      return checkDigit === Number(value.charAt(12));
    }

    function lookupLegacyApplicant() {
      const input = document.getElementById('legacyIdCard');
      const btn = document.getElementById('legacyLookupBtn');
      const idCard = input.value.replace(/\D/g, '');
      input.value = idCard;

      if (!isValidThaiId(idCard)) {
        Swal.fire({
          icon: 'warning',
          title: 'เลขประจำตัวประชาชนไม่ถูกต้อง',
          text: 'กรุณาตรวจสอบเลขประจำตัวประชาชน 13 หลักอีกครั้ง',
          confirmButtonColor: '#004a99'
        });
        return;
      }

      const original = btn.innerHTML;
      btn.disabled = true;
      btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin me-1"></i> กำลังตรวจสอบข้อมูล';

      google.script.run
        .withSuccessHandler((result) => {
          btn.disabled = false;
          btn.innerHTML = original;

          if (result && result.duplicate) {
            Swal.fire({
              icon: 'warning',
              title: 'ไม่สามารถสมัครซ้ำได้',
              text: result.message || 'เลขประจำตัวประชาชนนี้ส่งใบสมัครแล้ว',
              confirmButtonColor: '#004a99'
            });
            return;
          }

          if (!result || !result.found) {
            Swal.fire({
              icon: 'info',
              title: 'ไม่พบข้อมูลนักศึกษาเก่าปี 2569',
              text: result && result.message ? result.message : 'กรุณาตรวจสอบเลขประจำตัวประชาชน หรือเลือกประเภทนักศึกษาใหม่',
              confirmButtonColor: '#004a99'
            });
            return;
          }

          fillLegacyApplicant(result.data || {});
          openApplicationForm('returning2569');

          Swal.fire({
            toast: true,
            position: 'top-end',
            icon: 'success',
            title: 'ดึงข้อมูลเดิมเรียบร้อยแล้ว',
            text: 'กรุณาตรวจสอบและอัปเดตข้อมูลก่อนส่งใบสมัคร',
            showConfirmButton: false,
            timer: 2800
          });
        })
        .withFailureHandler((error) => {
          btn.disabled = false;
          btn.innerHTML = original;
          Swal.fire({
            icon: 'error',
            title: 'ไม่สามารถค้นหาข้อมูลได้',
            text: error && error.message ? error.message : 'เกิดข้อผิดพลาดในการเชื่อมต่อระบบ',
            confirmButtonColor: '#004a99'
          });
        })
        .getLegacyApplicantByIdCard(idCard);
    }

    function fillLegacyApplicant(data) {
      document.getElementById('idCard').value = String(data.idCard || '');
      document.getElementById('idCard').readOnly = true;
      document.getElementById('prefix').value = String(data.prefix || '');
      document.getElementById('fullName').value = String(data.fullName || '');
      document.getElementById('phone').value = String(data.phone || '').replace(/\D/g, '');
      document.getElementById('email').value = String(data.email || '');
      document.getElementById('lineId').value = String(data.lineId || '');
      document.getElementById('studentId').value = String(data.studentId || '').replace(/\D/g, '');
      document.getElementById('faculty').value = String(data.faculty || '');
      updateMajorOptions(String(data.major || ''));
      document.getElementById('year').value = String(data.year || '');
      document.getElementById('gpax').value = String(data.gpax || '');
      document.getElementById('englishLevel').value = String(data.englishLevel || '');
      document.getElementById('thaiLevel').value = String(data.thaiLevel || '');
      document.getElementById('wordLevel').value = String(data.wordLevel || '');
      document.getElementById('excelLevel').value = String(data.excelLevel || '');
      document.getElementById('powerPointLevel').value = String(data.powerPointLevel || '');
      document.getElementById('accessLevel').value = String(data.accessLevel || '');
      document.getElementById('skills').value = String(data.skills || '');
    }

    function updateMajorOptions(selectedValue = '') {
      const faculty = document.getElementById('faculty').value;
      const major = document.getElementById('major');
      const items = Array.isArray(facultyPrograms[faculty]) ? facultyPrograms[faculty] : [];

      major.innerHTML = '';
      if (!faculty) {
        major.innerHTML = '<option value="">กรุณาเลือกคณะก่อน...</option>';
        major.disabled = true;
        return;
      }

      major.disabled = false;
      major.appendChild(new Option('เลือกสาขาวิชา...', ''));

      items.forEach(item => {
        major.appendChild(new Option(item, item));
      });

      if (selectedValue) {
        const exists = items.includes(selectedValue);
        if (!exists) {
          major.appendChild(new Option(selectedValue, selectedValue));
        }
        major.value = selectedValue;
      }
    }

    function getSelectedInterestTypes() {
      return [...document.querySelectorAll('.interest-type:checked')].map(el => el.value);
    }

    function getSelectedReasons() {
      return [...document.querySelectorAll('.reason-option:checked')].map(el => el.value);
    }

    function getSelectedStudentPhoto() {
      const input = document.getElementById('studentPhoto');
      return input && input.files && input.files.length ? input.files[0] : null;
    }

    function validateStudentPhoto(file) {
      if (!file) return 'กรุณาอัปโหลดรูปถ่ายนักศึกษา 1 รูป';

      const allowedMime = new Set(['image/jpeg', 'image/png']);
      const allowedExt = /\.(jpe?g|png)$/i;
      const maxBytes = 2 * 1024 * 1024;

      if (!allowedMime.has(file.type) && !allowedExt.test(file.name)) {
        return 'รูปถ่ายต้องเป็นไฟล์ JPG/JPEG หรือ PNG เท่านั้น';
      }

      if (file.size > maxBytes) {
        return 'รูปถ่ายนักศึกษามีขนาดเกิน 2 MB';
      }

      return '';
    }

    function renderStudentPhotoPreview() {
      const file = getSelectedStudentPhoto();
      const input = document.getElementById('studentPhoto');
      const preview = document.getElementById('studentPhotoPreview');
      const placeholder = document.getElementById('studentPhotoPreviewPlaceholder');
      const error = validateStudentPhoto(file);

      if (error) {
        if (file) {
          Swal.fire({
            icon: 'warning',
            title: 'รูปถ่ายนักศึกษาไม่ถูกต้อง',
            text: error,
            confirmButtonColor: '#004a99'
          });
        }

        input.value = '';
        preview.removeAttribute('src');
        preview.style.display = 'none';
        placeholder.style.display = 'block';
        return;
      }

      const reader = new FileReader();
      reader.onload = () => {
        preview.src = String(reader.result || '');
        preview.style.display = 'block';
        placeholder.style.display = 'none';
      };
      reader.readAsDataURL(file);
    }

    function resetStudentPhotoPreview() {
      const preview = document.getElementById('studentPhotoPreview');
      const placeholder = document.getElementById('studentPhotoPreviewPlaceholder');

      if (preview) {
        preview.removeAttribute('src');
        preview.style.display = 'none';
      }

      if (placeholder) placeholder.style.display = 'block';
    }

    function getSelectedFiles() {
      return [...document.getElementById('resumeFiles').files];
    }

    function validateSelectedFiles(files) {
      if (files.length > 5) return 'แนบไฟล์ได้สูงสุด 5 ไฟล์';

      const allowedMime = new Set(['application/pdf', 'image/jpeg']);
      const allowedExt = /\.(pdf|jpe?g)$/i;
      const maxBytes = 5 * 1024 * 1024;

      for (const file of files) {
        if ((!allowedMime.has(file.type) && !allowedExt.test(file.name)) || file.size > maxBytes) {
          if (file.size > maxBytes) return `ไฟล์ ${file.name} มีขนาดเกิน 5 MB`;
          return `ไฟล์ ${file.name} ไม่ใช่ PDF หรือ JPG/JPEG`;
        }
      }
      return '';
    }

    function renderSelectedFiles() {
      const files = getSelectedFiles();
      const list = document.getElementById('selectedFilesList');
      list.innerHTML = '';

      files.forEach(file => {
        const li = document.createElement('li');
        li.textContent = `${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`;
        list.appendChild(li);
      });

      const error = validateSelectedFiles(files);
      if (error) {
        Swal.fire({
          icon: 'warning',
          title: 'ไฟล์แนบไม่ถูกต้อง',
          text: error,
          confirmButtonColor: '#004a99'
        });
        document.getElementById('resumeFiles').value = '';
        list.innerHTML = '';
      }
    }

    function readFileAsPayload(file) {
      return new Promise((resolve, reject) => {
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
        reader.onerror = () => reject(new Error(`อ่านไฟล์ ${file.name} ไม่สำเร็จ`));
        reader.readAsDataURL(file);
      });
    }

    function readStudentPhotoAsPayload(file) {
      return new Promise((resolve, reject) => {
        const img = new Image();
        const objectUrl = URL.createObjectURL(file);
        img.onload = () => {
          try {
            const maxWidth = 450;
            const maxHeight = 600;
            const ratio = Math.min(1, maxWidth / img.width, maxHeight / img.height);
            const width = Math.max(1, Math.round(img.width * ratio));
            const height = Math.max(1, Math.round(img.height * ratio));
            const canvas = document.createElement('canvas');
            canvas.width = width;
            canvas.height = height;
            canvas.getContext('2d').drawImage(img, 0, 0, width, height);
            canvas.toBlob((blob) => {
              URL.revokeObjectURL(objectUrl);
              if (!blob) return reject(new Error('ไม่สามารถปรับขนาดรูปถ่ายนักศึกษาได้'));
              const reader = new FileReader();
              reader.onload = () => {
                const result = String(reader.result || '');
                const comma = result.indexOf(',');
                const baseName = String(file.name || 'student-photo').replace(/\.[^.]+$/, '');
                resolve({
                  name: `${baseName}.jpg`,
                  mimeType: 'image/jpeg',
                  size: blob.size,
                  data: comma >= 0 ? result.slice(comma + 1) : result
                });
              };
              reader.onerror = () => reject(new Error('อ่านรูปถ่ายนักศึกษาไม่สำเร็จ'));
              reader.readAsDataURL(blob);
            }, 'image/jpeg', 0.82);
          } catch (error) {
            URL.revokeObjectURL(objectUrl);
            reject(error);
          }
        };
        img.onerror = () => {
          URL.revokeObjectURL(objectUrl);
          reject(new Error('ไม่สามารถเปิดรูปถ่ายนักศึกษาได้'));
        };
        img.src = objectUrl;
      });
    }

    async function uploadFilesAfterSubmit(result, data, files) {
      if (!files.length) return;

      try {
        const payloadFiles = await Promise.all(files.map(readFileAsPayload));

        google.script.run
          .withSuccessHandler(() => {
            Swal.fire({
              toast: true,
              position: 'top-end',
              icon: 'success',
              title: 'อัปโหลดไฟล์แนบเรียบร้อย',
              showConfirmButton: false,
              timer: 2200,
              timerProgressBar: true
            });
          })
          .withFailureHandler((error) => {
            Swal.fire({
              toast: true,
              position: 'top-end',
              icon: 'warning',
              title: 'บันทึกใบสมัครแล้ว แต่ไฟล์แนบอัปโหลดไม่สำเร็จ',
              text: error && error.message ? error.message : '',
              showConfirmButton: true
            });
          })
          .uploadApplicationFiles({
            applicationRow: result.applicationRow,
            applicationId: result.applicationId,
            studentId: data.studentId,
            fullName: data.fullName,
            files: payloadFiles
          });
      } catch (error) {
        Swal.fire({
          toast: true,
          position: 'top-end',
          icon: 'warning',
          title: 'บันทึกใบสมัครแล้ว แต่ไม่สามารถอ่านไฟล์แนบได้',
          text: error && error.message ? error.message : '',
          showConfirmButton: true
        });
      }
    }

    document.getElementById('searchInput').addEventListener('input', function () {
      this.value = this.value.replace(/\D/g, '').slice(0, 11);
    });

    document.getElementById('statusIdCardInput').addEventListener('input', function () {
      this.value = this.value.replace(/\D/g, '').slice(0, 13);
    });

    document.getElementById('statusIdCardInput').addEventListener('keydown', function (event) {
      if (event.key === 'Enter') searchStatus();
    });

    document.getElementById('studentPhoto').addEventListener('change', renderStudentPhotoPreview);
    document.getElementById('resumeFiles').addEventListener('change', renderSelectedFiles);

    document.getElementById('legacyIdCard').addEventListener('input', function () {
      this.value = this.value.replace(/\D/g, '').slice(0, 13);
    });

    document.getElementById('idCard').addEventListener('input', function () {
      this.value = this.value.replace(/\D/g, '').slice(0, 13);
    });

    document.getElementById('studentId').addEventListener('input', function () {
      this.value = this.value.replace(/\D/g, '').slice(0, 11);
    });

    document.getElementById('phone').addEventListener('input', function () {
      this.value = this.value.replace(/\D/g, '');
    });

    document.getElementById('applicationForm').addEventListener('submit', async function (e) {
      e.preventDefault();

      const form = e.currentTarget;
      const btn = document.getElementById('submitBtn');
      const originalText = btn.innerHTML;

      if (!form.checkValidity()) {
        form.reportValidity();
        return;
      }

      const idCard = document.getElementById('idCard').value.trim();
      const studentId = document.getElementById('studentId').value.trim();
      const phone = document.getElementById('phone').value.trim();
      const gpax = Number(document.getElementById('gpax').value);
      const interests = getSelectedInterestTypes();
      const reasons = getSelectedReasons();
      const studentPhoto = getSelectedStudentPhoto();
      const photoError = validateStudentPhoto(studentPhoto);
      const files = getSelectedFiles();
      const fileError = validateSelectedFiles(files);

      if (!isValidThaiId(idCard)) {
        Swal.fire({ icon: 'warning', title: 'เลขประจำตัวประชาชนไม่ถูกต้อง', text: 'กรุณาตรวจสอบเลขประจำตัวประชาชน 13 หลักอีกครั้ง' });
        return;
      }

      if (!/^\d{11}$/.test(studentId)) {
        Swal.fire({ icon: 'warning', title: 'รหัสนักศึกษาไม่ถูกต้อง', text: 'รหัสนักศึกษาไม่ถูกต้อง' });
        return;
      }

      if (!/^\d+$/.test(phone)) {
        Swal.fire({ icon: 'warning', title: 'เบอร์โทรศัพท์ไม่ถูกต้อง', text: 'กรุณากรอกเบอร์โทรศัพท์เป็นตัวเลขเท่านั้น' });
        return;
      }

      if (!Number.isFinite(gpax) || gpax < 0 || gpax > 4) {
        Swal.fire({ icon: 'warning', title: 'GPAX ไม่ถูกต้อง', text: 'กรุณากรอก GPAX ระหว่าง 0.00 - 4.00' });
        return;
      }

      if (!reasons.length) {
        Swal.fire({
          icon: 'warning',
          title: 'กรุณาเลือกเหตุผลที่สนใจสมัครงานนี้',
          text: 'เลือกอย่างน้อย 1 รายการ',
          confirmButtonColor: '#004a99'
        });
        return;
      }

      if (!interests.length) {
        Swal.fire({ icon: 'warning', title: 'กรุณาเลือกประเภทงานที่สนใจ', text: 'เลือกอย่างน้อย 1 รายการ' });
        return;
      }

      if (photoError) {
        Swal.fire({
          icon: 'warning',
          title: 'กรุณาอัปโหลดรูปถ่ายนักศึกษา',
          text: photoError,
          confirmButtonColor: '#004a99'
        });
        return;
      }

      if (fileError) {
        Swal.fire({ icon: 'warning', title: 'ไฟล์แนบไม่ถูกต้อง', text: fileError });
        return;
      }

      if (!document.getElementById('pdpaConsent').checked) {
        Swal.fire({ icon: 'warning', title: 'กรุณายอมรับข้อตกลง PDPA', text: 'ต้องยอมรับข้อตกลงก่อนส่งใบสมัคร' });
        return;
      }

      let studentPhotoPayload;

      try {
        studentPhotoPayload = await readStudentPhotoAsPayload(studentPhoto);
      } catch (error) {
        Swal.fire({
          icon: 'error',
          title: 'ระบบไม่สามารถอ่านรูปถ่ายนักศึกษาได้',
          text: error && error.message ? error.message : 'กรุณาเลือกรูปใหม่แล้วลองอีกครั้ง',
          confirmButtonColor: '#004a99'
        });
        return;
      }

      const data = {
        applicantType: document.getElementById('applicantType').value,
        idCard: idCard,
        prefix: document.getElementById('prefix').value,
        fullName: document.getElementById('fullName').value.trim(),
        phone: phone,
        email: document.getElementById('email').value.trim(),
        lineId: document.getElementById('lineId').value.trim(),
        studentId: studentId,
        faculty: document.getElementById('faculty').value,
        major: document.getElementById('major').value,
        year: document.getElementById('year').value,
        gpax: document.getElementById('gpax').value,
        englishLevel: document.getElementById('englishLevel').value,
        thaiLevel: document.getElementById('thaiLevel').value,
        wordLevel: document.getElementById('wordLevel').value,
        excelLevel: document.getElementById('excelLevel').value,
        powerPointLevel: document.getElementById('powerPointLevel').value,
        accessLevel: document.getElementById('accessLevel').value,
        skills: document.getElementById('skills').value.trim(),
        reason: reasons,
        interestTypes: interests,
        pdpaConsent: true,
        jobSelected: document.getElementById('jobSelected').value,
        jobDepartment: selectedJobForApplication && selectedJobForApplication.department
          ? String(selectedJobForApplication.department)
          : '',
        website: document.getElementById('websiteField').value,
        profileImage: studentPhotoPayload
      };

      btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin me-2"></i>ระบบกำลังบันทึกใบสมัคร';
      btn.disabled = true;

      google.script.run
        .withSuccessHandler((result) => {
          btn.innerHTML = originalText;
          btn.disabled = false;

          if (!result || !result.success) {
            Swal.fire({
              icon: 'error',
              title: 'ส่งใบสมัครไม่สำเร็จ',
              text: result && result.message ? result.message : 'ไม่สามารถบันทึกข้อมูลได้',
              confirmButtonColor: '#004a99'
            });
            return;
          }

          if (files.length) uploadFilesAfterSubmit(result, data, files);

          Swal.fire({
            icon: 'success',
            title: 'บันทึกใบสมัครสำเร็จ',
            html: `
              ระบบบันทึกใบสมัครตำแหน่ง <strong>${escapeHtml(result.job || '')}</strong> เรียบร้อยแล้ว
              <br><small class="text-muted">เลขประจำตัวประชาชนนี้ไม่สามารถส่งใบสมัครซ้ำได้อีก</small>
              <br><small class="text-muted">รูปถ่ายนักศึกษาถูกบันทึกเรียบร้อยแล้ว</small>
              ${files.length ? '<br><small class="text-muted">ไฟล์ Resume/Portfolio กำลังอัปโหลดต่อ</small>' : ''}
            `,
            confirmButtonColor: '#004a99'
          }).then(() => {
            form.reset();
            document.querySelectorAll('.interest-type').forEach(el => el.checked = false);
      document.querySelectorAll('.reason-option').forEach(el => el.checked = false);
            document.getElementById('selectedFilesList').innerHTML = '';
            resetStudentPhotoPreview();
            showLandingPage();
          });
        })
        .withFailureHandler((error) => {
          btn.innerHTML = originalText;
          btn.disabled = false;

          Swal.fire({
            icon: 'error',
            title: 'ส่งใบสมัครไม่สำเร็จ',
            text: error && error.message ? error.message : 'ไม่สามารถเชื่อมต่อกับฐานข้อมูลระบบกลางได้',
            confirmButtonColor: '#004a99'
          });
        })
        .submitApplication(data);
    });

    /* =========================================================
       Site visitor counter - FAST / REALTIME
       - แสดงค่าล่าสุดจาก localStorage ทันที ไม่รอเครือข่าย
       - Apps Script ใช้ Script Properties เป็นแหล่งอ่านหลัก
       - นับ 1 ครั้งต่อเบราว์เซอร์ในช่วงเวลา 30 นาที
       - refresh ยอดเบา ๆ ทุก 60 วินาทีเมื่อแท็บเปิดอยู่
       ========================================================= */
    const SITE_VISIT_STORAGE_KEY = 'ubuStudentJobsLastVisitAt';
    const SITE_VISIT_STATS_CACHE_KEY = 'ubuStudentJobsVisitStatsCacheV2';
    const SITE_VISIT_WINDOW_MS = 30 * 60 * 1000;
    const SITE_VISIT_LIVE_REFRESH_MS = 60 * 1000;
    let siteVisitRequestInFlight = false;
    let siteVisitRefreshTimer = null;

    function formatSiteVisitNumber(value) {
      const number = Number(value);
      if (!Number.isFinite(number) || number < 0) return '…';
      try {
        return new Intl.NumberFormat('th-TH').format(Math.floor(number));
      } catch (_) {
        return String(Math.floor(number));
      }
    }

    function getBangkokDateKey() {
      try {
        const parts = new Intl.DateTimeFormat('en-CA', {
          timeZone: 'Asia/Bangkok',
          year: 'numeric',
          month: '2-digit',
          day: '2-digit'
        }).formatToParts(new Date());

        const map = {};
        parts.forEach(part => {
          if (part.type !== 'literal') map[part.type] = part.value;
        });
        return `${map.year}-${map.month}-${map.day}`;
      } catch (_) {
        const now = new Date(Date.now() + (7 * 60 * 60 * 1000));
        return now.toISOString().slice(0, 10);
      }
    }

    function normalizeSiteVisitStats(stats) {
      if (!stats || typeof stats !== 'object') return null;

      const total = Number(stats.total);
      const today = Number(stats.today);
      if (!Number.isFinite(total) || total < 0 || !Number.isFinite(today) || today < 0) {
        return null;
      }

      return {
        total: Math.floor(total),
        today: Math.floor(today),
        date: String(stats.date || getBangkokDateKey()),
        updatedAt: Date.now()
      };
    }

    function renderSiteVisitStats(stats) {
      const totalEl = document.getElementById('siteVisitorTotal');
      const todayEl = document.getElementById('siteVisitorToday');
      if (!stats) return;

      if (totalEl) totalEl.textContent = formatSiteVisitNumber(stats.total);
      if (todayEl) todayEl.textContent = formatSiteVisitNumber(stats.today);
    }

    function loadCachedSiteVisitStats() {
      try {
        const raw = localStorage.getItem(SITE_VISIT_STATS_CACHE_KEY);
        if (!raw) return null;

        const cached = JSON.parse(raw);
        const normalized = normalizeSiteVisitStats(cached);
        if (!normalized) return null;

        // หากเป็นวันใหม่ ให้ยอด "วันนี้" เริ่มที่ 0 แต่ยังใช้ยอดรวมเดิมแสดงได้ทันที
        if (normalized.date !== getBangkokDateKey()) {
          normalized.today = 0;
          normalized.date = getBangkokDateKey();
        }
        return normalized;
      } catch (_) {
        return null;
      }
    }

    function saveCachedSiteVisitStats(stats) {
      const normalized = normalizeSiteVisitStats(stats);
      if (!normalized) return;

      try {
        localStorage.setItem(SITE_VISIT_STATS_CACHE_KEY, JSON.stringify(normalized));
      } catch (_) {}
    }

    function hasSiteVisitCookieConsent() {
      try {
        return localStorage.getItem('ubuStudentJobsCookieConsent') === '1.0';
      } catch (_) {
        return false;
      }
    }

    function getLastRecordedVisitAt() {
      try {
        const value = Number(localStorage.getItem(SITE_VISIT_STORAGE_KEY) || 0);
        return Number.isFinite(value) ? value : 0;
      } catch (_) {
        return 0;
      }
    }

    function saveLastRecordedVisitAt(timestamp) {
      try {
        localStorage.setItem(SITE_VISIT_STORAGE_KEY, String(timestamp));
      } catch (_) {}
    }

    async function refreshSiteVisitCounter(options = {}) {
      if (siteVisitRequestInFlight) return;
      if (document.visibilityState === 'hidden' && !options.force) return;

      siteVisitRequestInFlight = true;
      let baseline = loadCachedSiteVisitStats();
      let optimistic = null;

      try {
        const now = Date.now();
        const lastVisitAt = getLastRecordedVisitAt();
        const consentGranted = hasSiteVisitCookieConsent();
        const canRecord = consentGranted &&
          (!lastVisitAt || (now - lastVisitAt) >= SITE_VISIT_WINDOW_MS);

        // แสดงค่าล่าสุดทันทีโดยไม่ต้องรอ Apps Script
        if (baseline) renderSiteVisitStats(baseline);

        // ถ้ามี cache และเป็นการเข้าชมใหม่ ให้บวกเลขบนหน้าจอทันที
        // จากนั้นข้อมูลจริงจาก server จะเข้ามายืนยันอีกครั้ง
        if (canRecord && baseline) {
          optimistic = {
            total: baseline.total + 1,
            today: baseline.today + 1,
            date: getBangkokDateKey()
          };
          renderSiteVisitStats(optimistic);
        }

        const method = canRecord ? 'recordSiteVisit' : 'getSiteVisitStats';
        const stats = await publicServerCall(method);
        const normalized = normalizeSiteVisitStats(stats);

        if (normalized && stats.success !== false) {
          renderSiteVisitStats(normalized);
          saveCachedSiteVisitStats(normalized);

          if (canRecord) {
            saveLastRecordedVisitAt(now);
          }
        }
      } catch (error) {
        // หาก optimistic update ไม่สำเร็จ ให้กลับมาใช้ค่าที่ cache ไว้
        if (optimistic && baseline) {
          renderSiteVisitStats(baseline);
        }
        console.warn('ไม่สามารถโหลดสถิติผู้เข้าชมได้:', error);
      } finally {
        siteVisitRequestInFlight = false;
      }
    }

    function initSiteVisitorCounter() {
      // 1) แสดงค่าจาก cache ทันทีตั้งแต่วินาทีแรก
      const cached = loadCachedSiteVisitStats();
      if (cached) renderSiteVisitStats(cached);

      // 2) Sync กับ server แบบเบื้องหลัง
      refreshSiteVisitCounter({ force: true });

      // 3) หลังยอมรับคุกกี้ ให้บันทึกการเข้าชมทันที
      window.addEventListener('ubu:cookie-consent-accepted', function () {
        refreshSiteVisitCounter({ force: true });
      });

      // 4) อัปเดตยอดทุก 60 วินาทีเฉพาะขณะผู้ใช้เปิดหน้าเว็บอยู่
      if (!siteVisitRefreshTimer) {
        siteVisitRefreshTimer = window.setInterval(function () {
          if (document.visibilityState === 'visible') {
            refreshSiteVisitCounter();
          }
        }, SITE_VISIT_LIVE_REFRESH_MS);
      }

      // 5) กลับมาเปิดแท็บอีกครั้ง ให้ sync ทันที
      document.addEventListener('visibilitychange', function () {
        if (document.visibilityState === 'visible') {
          refreshSiteVisitCounter({ force: true });
        }
      });
    }

    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', initSiteVisitorCounter, { once: true });
    } else {
      initSiteVisitorCounter();
    }


    function goToDashboard() {
      window.location.href = './dashboard.html';
    }

    loadJobsFromSheet();
