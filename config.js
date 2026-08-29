/*
 * GitHub Pages frontend configuration
 *
 * หลัง Deploy Apps Script เป็น Web App แล้ว ให้นำ URL ที่ลงท้ายด้วย /exec
 * มาใส่ใน API_URL ด้านล่าง
 *
 * ตัวอย่าง:
 * https://script.google.com/macros/s/AKfycbxxxxxxxxxxxxxxxx/exec
 */
window.UBU_APP_CONFIG = Object.freeze({
  API_URL: 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_EXEC_URL_HERE',
  REQUEST_TIMEOUT_MS: 120000
});


/*
 * SECURITY NOTE
 * ห้ามใส่รหัสผ่าน, token, secret key หรือข้อมูลลับในไฟล์นี้
 * เพราะไฟล์บน GitHub Pages ผู้ใช้งานสามารถเปิดดู source ได้
 */
