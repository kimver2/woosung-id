/**
 * 우성아이디 대시보드 분리형 데이터 스냅샷 & 비동기 GAS 동기화 모듈
 * - HTML(App Shell)과 Google Sheets 데이터를 완전히 분리하여 0ms 즉시 렌더링 보장
 * - Stale-While-Revalidate 전략: Snapshot/localStorage로 즉시 렌더 후 백그라운드 비동기 동기화
 */
(function (global) {
  "use strict";

  var GAS_API_URL = "https://script.google.com/macros/s/AKfycbzd6pJp0oDfhN55dv6xM1HsAoQikA2Y_ApMz-H-GpaYqf3EKh3ffh6ULz-bU9n24dfTbg/exec";
  var CACHE_KEYS = {
    dashboard: "ws_dashboard_data_v3",
    todo: "ws_todo_data_v3",
    finance: "ws_finance_data_v3",
    settings: "ws_settings_data_v3",
    syncMeta: "ws_sync_meta_v3"
  };

  global.WS_SNAPSHOT_DASHBOARD = {"kpi": {"totalSites": 0, "activeSites": 0, "totalTasks": 0, "todayCount": 0, "delayedCount": 0}, "scopeCounts": {"site": 0, "siteDone": 0, "single": 0, "singleDone": 0, "as": 0, "asDone": 0, "newSite": 0, "newSiteDone": 0}, "todayStr": "", "members": ["차정률", "조항만", "송지은", "전예린", "한송이"], "tasks": ["착공", "설비", "철거", "누수보수", "전기", "전기입선", "전기마감", "조명", "에어컨", "목공", "도어설치", "중문", "창호", "창호실측", "타일", "줄눈", "도기", "SMC", "필름", "도장", "도배", "마루", "바닥", "가구", "주방", "보양", "청소", "마감", "상담", "현장실측", "AS", "방통", "양생", "전기타공", "도기세팅 천장", "준공청소", "현장촬영", "코킹", "전출", "홈파기", "잔마감", "전입", "마루철거", "S/A", "입주", "천장경량", "마루시공", "계약상담", "현장정리", "도기세팅", "기타", "현관부속", "수납가구", "장비 시연회(힐티)", "욕실 문의", "자재양중"], "partners": [], "partnerPhoneMap": {}, "memberPhoneMap": {}, "memberRoleMap": {"차정률": "현장소장", "조항만": "대표", "송지은": "과장", "전예린": "대리", "한송이": "주임"}, "siteSuperintendents": [], "holidays": {"2024-01-01": "신정", "2024-02-09": "설날연휴", "2024-02-10": "설날", "2024-02-11": "설날연휴", "2024-02-12": "대체공휴일", "2024-03-01": "삼일절", "2024-04-10": "총선", "2024-05-05": "어린이날", "2024-05-06": "대체공휴일", "2024-05-15": "부처님오신날", "2024-06-06": "현충일", "2024-08-15": "광복절", "2024-09-16": "추석연휴", "2024-09-17": "추석", "2024-09-18": "추석연휴", "2024-10-03": "개천절", "2024-10-09": "한글날", "2024-12-25": "성탄절", "2025-01-01": "신정", "2025-01-28": "설날연휴", "2025-01-29": "설날", "2025-01-30": "설날연휴", "2025-03-01": "삼일절", "2025-03-03": "대체공휴일", "2025-05-05": "어린이날", "2025-05-06": "대체공휴일", "2025-06-06": "현충일", "2025-08-15": "광복절", "2025-10-03": "개천절", "2025-10-05": "추석연휴", "2025-10-06": "추석", "2025-10-07": "추석연휴", "2025-10-08": "대체공휴일", "2025-10-09": "한글날", "2025-12-25": "성탄절", "2026-01-01": "신정", "2026-02-16": "설날연휴", "2026-02-17": "설날", "2026-02-18": "설날연휴", "2026-03-01": "삼일절", "2026-03-02": "대체공휴일", "2026-05-05": "어린이날", "2026-05-24": "부처님오신날", "2026-05-25": "대체공휴일", "2026-06-03": "지방선거", "2026-06-06": "현충일", "2026-08-15": "광복절", "2026-08-17": "대체공휴일", "2026-09-24": "추석연휴", "2026-09-25": "추석", "2026-09-26": "추석연휴", "2026-09-28": "대체공휴일", "2026-10-03": "개천절", "2026-10-05": "대체공휴일", "2026-10-09": "한글날", "2026-12-25": "성탄절", "2027-01-01": "신정", "2027-02-06": "설날연휴", "2027-02-07": "설날", "2027-02-08": "설날연휴", "2027-02-09": "대체공휴일", "2027-03-01": "삼일절", "2027-05-05": "어린이날", "2027-05-13": "부처님오신날", "2027-06-06": "현충일", "2027-06-07": "대체공휴일", "2027-08-15": "광복절", "2027-08-16": "대체공휴일", "2027-09-14": "추석연휴", "2027-09-15": "추석", "2027-09-16": "추석연휴", "2027-10-03": "개천절", "2027-10-04": "대체공휴일", "2027-10-09": "한글날", "2027-10-11": "대체공휴일", "2027-12-25": "성탄절"}, "sites": [], "siteInfoMap": {}};
  global.WS_SNAPSHOT_TODO = {"success": true, "list": []};
  global.WS_SNAPSHOT_FINANCE = {"updatedAt": "", "summary": {"totalContract": 0, "totalCollected": 0, "totalUncollected": 0, "totalExpense": 0}, "ledger": [], "collection": [], "expense": [], "fixedMaster": []};

  function safeReadStorage(key) {
    try {
      var raw = global.localStorage && global.localStorage.getItem(key);
      if (!raw) return null;
      return JSON.parse(raw);
    } catch (e) {
      return null;
    }
  }

  function safeWriteStorage(key, val) {
    try {
      if (global.localStorage) {
        global.localStorage.setItem(key, JSON.stringify(val));
      }
    } catch (e) {}
  }

  global.WSDataService = {
    GAS_API_URL: GAS_API_URL,

    getInitialDashboardData: function () {
      var cached = safeReadStorage(CACHE_KEYS.dashboard);
      if (cached && Array.isArray(cached.sites) && cached.sites.length > 0) {
        return cached;
      }
      return global.WS_SNAPSHOT_DASHBOARD;
    },

    getInitialTodoData: function () {
      var cached = safeReadStorage(CACHE_KEYS.todo);
      if (cached && Array.isArray(cached.list)) {
        return cached;
      }
      return global.WS_SNAPSHOT_TODO;
    },

    getInitialFinanceData: function () {
      var cached = safeReadStorage(CACHE_KEYS.finance);
      if (cached && (Array.isArray(cached.collection) || Array.isArray(cached.ledger))) {
        return cached;
      }
      return global.WS_SNAPSHOT_FINANCE;
    },

    saveFinanceCache: function (data) {
      if (data && typeof data === "object") {
        safeWriteStorage(CACHE_KEYS.finance, data);
      }
    },

    saveDashboardCache: function (data) {
      if (data && Array.isArray(data.sites)) {
        safeWriteStorage(CACHE_KEYS.dashboard, data);
      }
    },

    saveTodoCache: function (data) {
      if (data && Array.isArray(data.list)) {
        safeWriteStorage(CACHE_KEYS.todo, data);
      }
    },

    getLastSyncMeta: function () {
      return safeReadStorage(CACHE_KEYS.syncMeta) || { source: "snapshot", updatedAt: "2026-11-13 15:40" };
    },

    syncLiveFromGAS: function (optsOrCb, onUpdated) {
      if (typeof fetch !== "function") {
        return Promise.resolve({ skipped: true });
      }
      var opts = (typeof optsOrCb === "object" && optsOrCb !== null) ? optsOrCb : {};
      var cb = (typeof optsOrCb === "function") ? optsOrCb : onUpdated;
      var t = Date.now();
      var dashReq = fetch(GAS_API_URL + "?action=getDashboardData&_t=" + t, { method: "GET", mode: "cors" })
        .then(function (r) { return r.json(); })
        .catch(function () { return null; });
      var todoReq = fetch(GAS_API_URL + "?action=getTodoList&_t=" + t, { method: "GET", mode: "cors" })
        .then(function (r) { return r.json(); })
        .catch(function () { return null; });
      var settingsReq = fetch(GAS_API_URL + "?action=getSettingsData&format=json&_t=" + t, { method: "GET", mode: "cors" })
        .then(function (r) { return r.json(); })
        .catch(function () { return null; });

      var promises = [dashReq, todoReq, settingsReq];
      var syncFinance = !!opts.syncFinance && !!opts.authUser;
      if (syncFinance) {
        promises.push(this.getFinanceDbData(opts.authUser));
      }

      var self = this;
      return Promise.all(promises).then(function (results) {
        var liveDash = results[0];
        var liveTodo = results[1];
        var liveSettings = results[2];
        var liveFinance = syncFinance ? results[3] : null;
        var normalizedSettings = null;
        var changed = false;
        if (liveDash && Array.isArray(liveDash.sites) && liveDash.sites.length > 0) {
          if (Array.isArray(liveDash.tasks) && liveDash.tasks.indexOf("착공") === -1) {
            liveDash.tasks.unshift("착공");
          }
          safeWriteStorage(CACHE_KEYS.dashboard, liveDash);
          changed = true;
        }
        if (liveTodo && Array.isArray(liveTodo.list)) {
          safeWriteStorage(CACHE_KEYS.todo, liveTodo);
          changed = true;
        }
        if (liveSettings && liveSettings.success && (Array.isArray(liveSettings.members) || Array.isArray(liveSettings.partners) || Array.isArray(liveSettings.tasks))) {
          var syncedTasks = (liveSettings.tasks || []).map(function (t) {
            return Object.assign({}, t, { name: t.name || t.taskName || "" });
          });
          if (!syncedTasks.some(function (t) { return t && String(t.name || t.taskName || "").trim() === "착공"; })) {
            syncedTasks.unshift({ row: 1, name: "착공", taskName: "착공", category: "사전/준비", note: "공사 착공 시작" });
          }
          normalizedSettings = {
            partners: (liveSettings.partners || []).map(function (p) {
              var fullName = p.fullName || p.dropdownVal || (p.biz ? p.biz + "_" + p.name : p.name);
              return Object.assign({}, p, { fullName: fullName });
            }),
            tasks: syncedTasks,
            members: (liveSettings.members || []).map(function (m) {
              return Object.assign({}, m, { name: m.name || "", role: m.role || "" });
            })
          };
          safeWriteStorage(CACHE_KEYS.settings, normalizedSettings);
          changed = true;
        }
        if (liveFinance && liveFinance.success) {
          var cleanFinance = {
            collection: Array.isArray(liveFinance.collection) ? liveFinance.collection : [],
            expense: Array.isArray(liveFinance.expense) ? liveFinance.expense : [],
            fixedMaster: Array.isArray(liveFinance.fixedMaster) ? liveFinance.fixedMaster : [],
            updatedAt: liveFinance.updatedAt || ""
          };
          self.saveFinanceCache(cleanFinance);
          changed = true;
        }
        var now = new Date();
        var hh = String(now.getHours()).padStart(2, "0");
        var mm = String(now.getMinutes()).padStart(2, "0");
        var ss = String(now.getSeconds()).padStart(2, "0");
        var meta = { source: changed ? "live_gas" : "snapshot", updatedAt: hh + ":" + mm + ":" + ss };
        safeWriteStorage(CACHE_KEYS.syncMeta, meta);
        var resObj = {
          dashData: liveDash && Array.isArray(liveDash.sites) ? liveDash : null,
          todoData: liveTodo && Array.isArray(liveTodo.list) ? liveTodo : null,
          settingsData: normalizedSettings,
          financeData: (liveFinance && liveFinance.success) ? liveFinance : null,
          meta: meta
        };
        if (typeof cb === "function") {
          cb(resObj);
        }
        return resObj;
      });
    },

    callGasApi: function (action, params) {
      if (typeof fetch !== "function" || typeof location === "undefined" || !String(location.protocol || "").startsWith("http")) {
        return Promise.resolve({ skipped: true, ok: true, offline: true });
      }
      var query = ["action=" + encodeURIComponent(action)];
      var hasTimestamp = false;
      if (params && typeof params === "object") {
        Object.keys(params).forEach(function (k) {
          var v = params[k];
          if (k === "_t") hasTimestamp = true;
          if (v !== undefined && v !== null) {
            query.push(encodeURIComponent(k) + "=" + encodeURIComponent(typeof v === "object" ? JSON.stringify(v) : String(v)));
          }
        });
      }
      if (!hasTimestamp) {
        query.push("_t=" + Date.now());
      }
      if (!/^get/i.test(String(action || ""))) {
        try { global.localStorage.setItem("ws_last_mutation_start_ms", String(Date.now())); } catch (eM) {}
      }
      var url = GAS_API_URL + "?" + query.join("&");
      return fetch(url, { method: "GET", mode: "cors", redirect: "follow", keepalive: true })
        .then(function (r) { return r.text(); })
        .then(function (txt) {
          try {
            var parsed = JSON.parse(txt);
            if (parsed && parsed.ok === undefined && parsed.success !== undefined) parsed.ok = !!parsed.success;
            return parsed;
          } catch (e) {
            var isHtml = /<html|<body|<!doctype/i.test(String(txt || ""));
            return { ok: !isHtml, htmlFallback: isHtml, raw: String(txt || "").slice(0, 200) };
          }
        })
        .catch(function (err) {
          return { ok: false, offlineFallback: true, error: err && err.message ? err.message : String(err) };
        });
    },

    getInitialSettingsData: function () {
      var snap = global.WS_SNAPSHOT_DASHBOARD || {};
      var defaultRoleMap = Object.assign(
        { "차정률": "현장소장", "조항만": "대표", "송지은": "과장", "전예린": "대리", "한송이": "주임" },
        snap.memberRoleMap || {}
      );
      var cached = safeReadStorage(CACHE_KEYS.settings);
      if (cached && (Array.isArray(cached.partners) || Array.isArray(cached.tasks) || Array.isArray(cached.members))) {
        if (Array.isArray(cached.members)) {
          cached.members = cached.members.map(function (m) {
            if (!m || !m.name) return m;
            var curRole = String(m.role || "").trim();
            if ((!curRole || curRole === "소장") && defaultRoleMap[m.name]) {
              return Object.assign({}, m, { role: defaultRoleMap[m.name] });
            }
            return m;
          });
        }
        if (!Array.isArray(cached.tasks)) cached.tasks = [];
        if (!cached.tasks.some(function (t) { return t && String(t.name || t.taskName || "").trim() === "착공"; })) {
          cached.tasks.unshift({ row: 1, name: "착공", taskName: "착공", category: "사전/준비", note: "공사 착공 시작" });
        }
        return cached;
      }
      var pMap = snap.partnerPhoneMap || {};
      var mMap = snap.memberPhoneMap || {};
      var partners = (snap.partners || []).map(function (full, i) {
        var idx = full.indexOf("_");
        var cat = idx > -1 ? full.slice(0, idx) : "";
        var name = idx > -1 ? full.slice(idx + 1) : full;
        return {
          row: i + 2,
          category: cat,
          biz: cat,
          name: name,
          phone: pMap[full] || pMap[name] || "",
          fullName: full,
          priority: "일반",
          manager: "",
          status: "운영 중",
          feature: "",
          orderMethod: "",
          bankAccount: "",
          note: ""
        };
      });
      var members = (snap.members || []).map(function (m, mi) {
        return {
          row: mi + 2,
          name: m,
          role: defaultRoleMap[m] || "소장",
          phone: mMap[m] || "",
          email: "",
          receive: true,
          notifyEnabled: true,
          note: ""
        };
      });
      var tasks = (snap.tasks || []).map(function (t, ti) {
        return {
          row: ti + 2,
          name: t,
          taskName: t,
          category: t === "착공" ? "사전/준비" : "기타",
          note: t === "착공" ? "공사 착공 시작" : ""
        };
      });
      if (!tasks.some(function (t) { return t && String(t.name || t.taskName || "").trim() === "착공"; })) {
        tasks.unshift({ row: 1, name: "착공", taskName: "착공", category: "사전/준비", note: "공사 착공 시작" });
      }
      return { partners: partners, tasks: tasks, members: members };
    },

    saveSettingsCache: function (data) {
      if (data && typeof data === "object") {
        safeWriteStorage(CACHE_KEYS.settings, data);
      }
    },

    syncSettingsFromGAS: function () {
      return this.callGasApi("getSettingsData", {});
    },

    saveTask: function (isEdit, payload, dashData) {
      if (dashData) this.saveDashboardCache(dashData);
      if (isEdit) {
        return this.callGasApi("updateScheduleStatus", Object.assign({ format: "json", id: (payload && (payload.id || payload.taskId)) || "" }, payload));
      }
      return this.callGasApi("addScheduleTask", payload);
    },

    deleteTask: function (payload, dashData) {
      if (dashData) this.saveDashboardCache(dashData);
      return this.callGasApi("deleteScheduleTask", payload);
    },

    moveTaskDate: function (payload, dashData) {
      if (dashData) this.saveDashboardCache(dashData);
      return this.callGasApi("moveScheduleTaskDate", payload);
    },

    cascadeShiftTasks: function (payload, dashData) {
      if (dashData) this.saveDashboardCache(dashData);
      return this.callGasApi("cascadeShiftSiteTasks", payload);
    },

    deleteSite: function (payload, dashData) {
      if (dashData) this.saveDashboardCache(dashData);
      return this.callGasApi("deleteSite", payload);
    },

    saveSiteInfo: function (payload, dashData) {
      if (dashData) this.saveDashboardCache(dashData);
      return this.callGasApi("saveSiteInfo", payload);
    },

    toggleChecklist: function (payload, dashData) {
      if (dashData) this.saveDashboardCache(dashData);
      return this.callGasApi("toggleSitePreChecklist", payload);
    },

    sortScheduleSheet: function () {
      return this.callGasApi("sortScheduleSheet", {});
    },

    archiveCompletedSites: function () {
      return this.callGasApi("archiveCompletedSites", {});
    },

    changeSharedPassword: function (newPin, authUser) {
      return this.callGasApi("changeSharedPassword", { newPin: newPin, authUser: authUser });
    },

    changeAdminPassword: function (newPin, authUser) {
      return this.callGasApi("changeAdminPassword", { newPin: newPin, authUser: authUser });
    },

    callGasPost: function (action, payload) {
      if (typeof fetch !== "function" || typeof location === "undefined" || !String(location.protocol || "").startsWith("http")) {
        return Promise.resolve({ success: true, offline: true });
      }
      var body = Object.assign({ action: action }, payload || {});
      try { global.localStorage.setItem("ws_last_mutation_start_ms", String(Date.now())); } catch (eMp) {}
      return fetch(GAS_API_URL, {
        method: "POST",
        mode: "cors",
        redirect: "follow",
        keepalive: true,
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(body)
      }).then(function (r) {
        return r.json();
      }).catch(function (err) {
        console.warn("GAS POST error for " + action + ":", err);
        return { success: false, error: String(err && err.message || err) };
      });
    },

    saveTodo: function (isEdit, payload, todoData) {
      if (todoData) this.saveTodoCache(todoData);
      if (isEdit) {
        return this.callGasApi("updateTodo", { id: payload.id, data: payload });
      }
      return this.callGasApi("addTodo", { data: payload });
    },

    toggleTodoStatus: function (payload, todoData) {
      if (todoData) this.saveTodoCache(todoData);
      return this.callGasApi("updateTodoStatus", payload);
    },

    deleteTodo: function (id, todoData) {
      if (todoData) this.saveTodoCache(todoData);
      return this.callGasApi("deleteTodo", { id: id });
    },

    saveEstimate: function (isEdit, payload, dashData) {
      if (dashData) this.saveDashboardCache(dashData);
      if (isEdit) {
        return this.callGasApi("updateEstimateRow", payload);
      }
      return this.callGasApi("addEstimateRow", payload);
    },

    promoteEstimateToSite: function (payload, dashData) {
      if (dashData) this.saveDashboardCache(dashData);
      if (typeof google !== "undefined" && google.script && google.script.run) {
        return new Promise(function (resolve, reject) {
          google.script.run
            .withSuccessHandler(function (res) { resolve(res || { success: true }); })
            .withFailureHandler(function (err) { reject(err); })
            .promoteEstimateToSite(payload);
        });
      }
      return this.callGasApi("promoteEstimateToSite", payload);
    },

    saveSettings: function (settingsData) {
      this.saveSettingsCache(settingsData);
      if (typeof google !== "undefined" && google.script && google.script.run) {
        return new Promise(function (resolve, reject) {
          google.script.run
            .withSuccessHandler(function (res) { resolve(res || { success: true }); })
            .withFailureHandler(function (err) { reject(err); })
            .saveAllSettingsBatchFromWeb(settingsData);
        });
      }
      return this.callGasPost("saveAllSettingsBatch", { data: settingsData });
    },

    uploadSiteFiles: function (siteName, filesPayload, clientDriveUrl) {
      if (typeof fetch !== "function" || typeof location === "undefined" || !String(location.protocol || "").startsWith("http")) {
        return Promise.resolve({ success: true, count: (filesPayload || []).length, offline: true });
      }
      var body = JSON.stringify({
        action: "uploadSiteFiles",
        siteName: siteName,
        clientDriveUrl: clientDriveUrl || "",
        files: filesPayload
      });
      return fetch(GAS_API_URL, {
        method: "POST",
        mode: "cors",
        redirect: "follow",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: body
      }).then(function (r) {
        if (!r.ok) throw new Error("HTTP " + r.status);
        return r.json();
      });
    },

    getFinanceDbData: function (authUser) {
      var user = authUser || "";
      if (typeof fetch !== "function" || typeof location === "undefined" || !String(location.protocol || "").startsWith("http")) {
        return Promise.resolve({ success: true, offline: true, collection: [], expense: [], fixedMaster: [] });
      }
      return this.callGasApi("getFinanceDbData", { authUser: user, _t: Date.now() });
    },

    saveCollectionRow: function (row, authUser) {
      var user = authUser || "";
      if (typeof google !== "undefined" && google.script && google.script.run) {
        return new Promise(function (resolve, reject) {
          google.script.run
            .withSuccessHandler(function (res) { resolve(res || { success: true }); })
            .withFailureHandler(function (err) { reject(err); })
            .saveCollectionRow({ row: row, authUser: user });
        });
      }
      return this.callGasPost("saveCollectionRow", { row: row, authUser: user });
    },

    deleteCollectionRow: function (id, authUser) {
      var user = authUser || "";
      if (typeof google !== "undefined" && google.script && google.script.run) {
        return new Promise(function (resolve, reject) {
          google.script.run
            .withSuccessHandler(function (res) { resolve(res || { success: true }); })
            .withFailureHandler(function (err) { reject(err); })
            .deleteCollectionRow({ id: id, authUser: user });
        });
      }
      return this.callGasPost("deleteCollectionRow", { id: id, authUser: user });
    },

    saveExpenseRow: function (row, authUser) {
      var user = authUser || "";
      if (typeof google !== "undefined" && google.script && google.script.run) {
        return new Promise(function (resolve, reject) {
          google.script.run
            .withSuccessHandler(function (res) { resolve(res || { success: true }); })
            .withFailureHandler(function (err) { reject(err); })
            .saveExpenseRow({ row: row, authUser: user });
        });
      }
      return this.callGasPost("saveExpenseRow", { row: row, authUser: user });
    },

    deleteExpenseRow: function (id, authUser) {
      var user = authUser || "";
      if (typeof google !== "undefined" && google.script && google.script.run) {
        return new Promise(function (resolve, reject) {
          google.script.run
            .withSuccessHandler(function (res) { resolve(res || { success: true }); })
            .withFailureHandler(function (err) { reject(err); })
            .deleteExpenseRow({ id: id, authUser: user });
        });
      }
      return this.callGasPost("deleteExpenseRow", { id: id, authUser: user });
    },

    generateFixedExpensesForMonth: function (settleMonth, authUser) {
      var user = authUser || "";
      if (typeof google !== "undefined" && google.script && google.script.run) {
        return new Promise(function (resolve, reject) {
          google.script.run
            .withSuccessHandler(function (res) { resolve(res || { success: true }); })
            .withFailureHandler(function (err) { reject(err); })
            .generateFixedExpensesForMonth({ settleMonth: settleMonth, authUser: user });
        });
      }
      return this.callGasPost("generateFixedExpensesForMonth", { settleMonth: settleMonth, authUser: user });
    },

    saveExpenseBatchRows: function (rows, authUser) {
      var user = authUser || "";
      if (typeof google !== "undefined" && google.script && google.script.run) {
        return new Promise(function (resolve, reject) {
          google.script.run
            .withSuccessHandler(function (res) { resolve(res || { success: true }); })
            .withFailureHandler(function (err) { reject(err); })
            .saveExpenseBatchRows({ rows: rows, authUser: user });
        });
      }
      return this.callGasPost("saveExpenseBatchRows", { rows: rows, authUser: user });
    },

    saveFixedMasterRow: function (row, authUser) {
      var user = authUser || "";
      if (typeof google !== "undefined" && google.script && google.script.run) {
        return new Promise(function (resolve, reject) {
          google.script.run
            .withSuccessHandler(function (res) { resolve(res || { success: true }); })
            .withFailureHandler(function (err) { reject(err); })
            .saveFixedMasterRow({ row: row, authUser: user });
        });
      }
      return this.callGasPost("saveFixedMaster", { row: row, authUser: user });
    },

    deleteFixedMasterRow: function (id, authUser) {
      var user = authUser || "";
      if (typeof google !== "undefined" && google.script && google.script.run) {
        return new Promise(function (resolve, reject) {
          google.script.run
            .withSuccessHandler(function (res) { resolve(res || { success: true }); })
            .withFailureHandler(function (err) { reject(err); })
            .deleteFixedMasterRow({ id: id, authUser: user });
        });
      }
      return this.callGasPost("deleteFixedMaster", { id: id, authUser: user });
    },

    cleanupCorruptedFixedExpenses: function (settleMonth, authUser) {
      var user = authUser || "";
      var m = settleMonth || "2026-11";
      if (typeof google !== "undefined" && google.script && google.script.run) {
        return new Promise(function (resolve, reject) {
          google.script.run
            .withSuccessHandler(function (res) { resolve(res || { success: true }); })
            .withFailureHandler(function (err) { reject(err); })
            .cleanupCorruptedFixedExpenses({ settleMonth: m, authUser: user });
        });
      }
      return this.callGasPost("cleanupCorruptedFixedExpenses", { settleMonth: m, authUser: user });
    },

    mapPartnerBankAccounts: function () {
      if (typeof google !== "undefined" && google.script && google.script.run) {
        return new Promise(function (resolve, reject) {
          google.script.run
            .withSuccessHandler(function (res) { resolve(res || { success: true }); })
            .withFailureHandler(function (err) { reject(err); })
            .mapPartnerBankAccountsFromFinanceDb();
        });
      }
      return this.callGasPost("mapPartnerBankAccounts", {});
    }
  };
})(typeof window !== "undefined" ? window : globalThis);
