// =====================================================
// GLOBAL CONSTANTS
// =====================================================
// TABLE COLUMN INDEX
const COL_STT = 0;
const COL_UNIT = 1;
const COL_ACTION = 2;
const COL_STATUS = 3;
const COL_UPDATED = 4;
// REPORT STATUS
const STATUS_ALL = "Tất cả";
const STATUS_SUBMITTED = "Đã gửi";
const STATUS_EDITING = "Đang nhập";
const STATUS_NOT_STARTED = "Chưa nhập";
// FETCH INTERCEPTOR
// =====================================================
/* Tự gắn JWT vào mọi fetch request */
const _originalFetch = window.fetch;
window.fetch = function (url, options = {}) {
   const u = JSON.parse(localStorage.getItem("user") || "{}");
   if (u.token) {
      options.headers = options.headers || {};
      options.headers["Authorization"] = "Bearer " + u.token;
   }
   return _originalFetch(url, options);
};
// =====================================================
// AUTHORIZATION
// =====================================================
const user = JSON.parse(localStorage.getItem("user"));
const token = user.token;
const headers = {
   Authorization: "Bearer " + token
};
if (!user || (user.role !== "admin" && user.role !== "viewer")) {
   window.location = "login.html";
}