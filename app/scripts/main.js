/**
 * DVIR Dashboard — MyGeotab Add-In
 *
 * Provides visibility into outstanding DVIRs (Driver Vehicle Inspection Reports)
 * and DVIRs where defects were marked "repair not necessary for safe operation."
 */

geotab.addin.dvirDashboard = function () {
  "use strict";

  // ── State ──────────────────────────────────────────────────────────────
  var api;
  var allDevices = [];
  var allGroups = {};
  var deviceMap = {};        // deviceId -> device
  var driverMap = {};        // driverId -> driver
  var abortController = null;
  var firstFocus = true;
  var activeTab = "fleet";

  // Computed data (populated on Apply)
  var dvirData = {
    logs: [],           // raw DVIRLog entities
    fleetRows: [],      // per-DVIR summary rows
    defectRows: []      // per-defect detail rows
  };

  // Sort state per table
  var sortState = {
    fleet: { col: "date", dir: "desc" },
    defects: { col: "date", dir: "desc" }
  };

  // ── DOM refs (set during initialize) ───────────────────────────────────
  var els = {};

  // ── Helpers ────────────────────────────────────────────────────────────

  function $(id) {
    return document.getElementById(id);
  }

  function escapeHtml(str) {
    var div = document.createElement("div");
    div.textContent = str || "";
    return div.innerHTML;
  }

  function formatDate(d) {
    if (!d) return "--";
    var dt = new Date(d);
    return (dt.getMonth() + 1) + "/" + dt.getDate() + "/" + dt.getFullYear();
  }

  function formatDateTime(d) {
    if (!d) return "--";
    var dt = new Date(d);
    return (dt.getMonth() + 1) + "/" + dt.getDate() + "/" + dt.getFullYear() + " " +
      String(dt.getHours()).padStart(2, "0") + ":" + String(dt.getMinutes()).padStart(2, "0");
  }

  function getDateRange() {
    var now = new Date();
    var preset = document.querySelector(".dvir-preset.active");
    var key = preset ? preset.dataset.preset : "7days";
    var from, to;

    to = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);

    switch (key) {
      case "yesterday":
        from = new Date(now);
        from.setDate(from.getDate() - 1);
        from.setHours(0, 0, 0, 0);
        to = new Date(from);
        to.setHours(23, 59, 59);
        break;
      case "7days":
        from = new Date(now);
        from.setDate(from.getDate() - 7);
        from.setHours(0, 0, 0, 0);
        break;
      case "30days":
        from = new Date(now);
        from.setDate(from.getDate() - 30);
        from.setHours(0, 0, 0, 0);
        break;
      case "alltime":
        from = new Date(now);
        from.setDate(from.getDate() - 365);
        from.setHours(0, 0, 0, 0);
        break;
      default:
        from = new Date(now);
        from.setDate(from.getDate() - 7);
        from.setHours(0, 0, 0, 0);
        break;
    }

    return { from: from.toISOString(), to: to.toISOString() };
  }

  function isAborted() {
    return abortController && abortController.signal && abortController.signal.aborted;
  }

  function showLoading(show, text) {
    els.loading.style.display = show ? "flex" : "none";
    els.empty.style.display = "none";
    if (text) els.loadingText.textContent = text;
  }

  function showEmpty(show) {
    els.empty.style.display = show ? "flex" : "none";
  }

  function setProgress(pct) {
    els.progressBar.style.width = Math.min(100, Math.round(pct)) + "%";
  }

  function showWarning(msg) {
    els.warning.style.display = msg ? "block" : "none";
    els.warning.textContent = msg || "";
  }

  // ── API Helpers ────────────────────────────────────────────────────────

  function delay(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
  }

  function apiCall(method, params) {
    return new Promise(function (resolve, reject) {
      api.call(method, params, resolve, reject);
    });
  }

  function apiMultiCall(calls) {
    return new Promise(function (resolve, reject) {
      api.multiCall(calls, resolve, reject);
    });
  }

  // ── Dropdown Population ────────────────────────────────────────────────

  function populateGroupDropdown() {
    var current = els.group.value;
    els.group.innerHTML = '<option value="all">All Groups</option>';

    var skipIds = { GroupCompanyId: true, GroupNothingId: true };
    var groupList = [];
    Object.keys(allGroups).forEach(function (gid) {
      var g = allGroups[gid];
      if (skipIds[gid]) return;
      if (!g.name || g.name === "CompanyGroup" || g.name === "**Nothing**") return;
      groupList.push(g);
    });
    groupList.sort(function (a, b) { return (a.name || "").localeCompare(b.name || ""); });

    groupList.forEach(function (g) {
      var opt = document.createElement("option");
      opt.value = g.id;
      opt.textContent = g.name || g.id;
      els.group.appendChild(opt);
    });
    if (current && els.group.querySelector('option[value="' + current + '"]')) {
      els.group.value = current;
    }
  }

  function populateVehicleDropdown() {
    var current = els.vehicle.value;
    els.vehicle.innerHTML = '<option value="all">All Vehicles</option>';
    var sorted = allDevices.slice().sort(function (a, b) {
      return (a.name || "").localeCompare(b.name || "");
    });
    sorted.forEach(function (d) {
      var opt = document.createElement("option");
      opt.value = d.id;
      opt.textContent = d.name || d.id;
      els.vehicle.appendChild(opt);
    });
    if (current && els.vehicle.querySelector('option[value="' + current + '"]')) {
      els.vehicle.value = current;
    }
  }

  // ── Filtered Devices ───────────────────────────────────────────────────

  function filteredDeviceIds() {
    var vehicleId = els.vehicle.value;
    var groupId = els.group.value;

    if (vehicleId !== "all") {
      var set = {};
      set[vehicleId] = true;
      return set;
    }

    var set = {};
    allDevices.forEach(function (dev) {
      if (groupId !== "all") {
        var devGroups = dev.groups || [];
        var inGroup = false;
        for (var i = 0; i < devGroups.length; i++) {
          if (devGroups[i].id === groupId) { inGroup = true; break; }
        }
        if (!inGroup) return;
      }
      set[dev.id] = true;
    });
    return set;
  }

  // ── DVIR Data Fetch ────────────────────────────────────────────────────

  function fetchDVIRLogs(dateRange, onProgress) {
    var CHUNK_DAYS = 7;
    var fromMs = new Date(dateRange.from).getTime();
    var toMs = new Date(dateRange.to).getTime();
    var chunks = [];
    var cursor = fromMs;
    while (cursor < toMs) {
      var chunkEnd = Math.min(cursor + CHUNK_DAYS * 86400000, toMs);
      chunks.push({
        from: new Date(cursor).toISOString(),
        to: new Date(chunkEnd).toISOString()
      });
      cursor = chunkEnd;
    }

    var totalChunks = chunks.length;
    var completedChunks = 0;
    var allLogStubs = [];

    // Phase 1: Fetch all DVIRLog stubs (no dvirDefects in list queries)
    return chunks.reduce(function (chain, chunk, chunkIdx) {
      return chain.then(function () {
        if (isAborted()) return;
        var pause = chunkIdx > 0 ? delay(300) : Promise.resolve();
        return pause.then(function () {
          if (isAborted()) return;
          return apiCall("Get", {
            typeName: "DVIRLog",
            search: {
              fromDate: chunk.from,
              toDate: chunk.to
            }
          }).then(function (logs) {
            allLogStubs = allLogStubs.concat(logs || []);
            completedChunks++;
            if (onProgress) onProgress((completedChunks / totalChunks) * 50);
          });
        });
      });
    }, Promise.resolve()).then(function () {
      if (isAborted()) return [];
      if (allLogStubs.length === 0) return [];

      // Phase 2: Re-fetch each DVIRLog by ID to get dvirDefects populated
      // Batch into multiCall groups of 50
      var BATCH = 50;
      var calls = allLogStubs.map(function (log) {
        return ["Get", { typeName: "DVIRLog", search: { id: log.id } }];
      });
      var batches = [];
      for (var i = 0; i < calls.length; i += BATCH) {
        batches.push(calls.slice(i, i + BATCH));
      }

      var completedBatches = 0;
      var totalBatches = batches.length;
      var fullLogs = [];

      return batches.reduce(function (chain, batch, batchIdx) {
        return chain.then(function () {
          if (isAborted()) return;
          var pause = batchIdx > 0 ? delay(300) : Promise.resolve();
          return pause.then(function () {
            if (isAborted()) return;
            return apiMultiCall(batch).then(function (results) {
              results.forEach(function (arr) {
                if (Array.isArray(arr) && arr.length > 0) {
                  fullLogs.push(arr[0]);
                }
              });
              completedBatches++;
              if (onProgress) onProgress(50 + (completedBatches / totalBatches) * 50);
            });
          });
        });
      }, Promise.resolve()).then(function () {
        return fullLogs;
      });
    });
  }

  function fetchDrivers(driverIds) {
    if (driverIds.length === 0) return Promise.resolve();

    // Batch fetch unique driver IDs
    var unique = [];
    var seen = {};
    driverIds.forEach(function (id) {
      if (!seen[id] && !driverMap[id]) {
        seen[id] = true;
        unique.push(id);
      }
    });

    if (unique.length === 0) return Promise.resolve();

    var calls = unique.map(function (id) {
      return ["Get", { typeName: "User", search: { id: id } }];
    });

    // Batch in groups of 50
    var BATCH = 50;
    var batches = [];
    for (var i = 0; i < calls.length; i += BATCH) {
      batches.push(calls.slice(i, i + BATCH));
    }

    return batches.reduce(function (chain, batch) {
      return chain.then(function () {
        return apiMultiCall(batch).then(function (results) {
          results.forEach(function (arr) {
            if (Array.isArray(arr) && arr.length > 0) {
              var user = arr[0];
              driverMap[user.id] = user;
            }
          });
        });
      });
    }, Promise.resolve());
  }

  // ── DVIR Classification ────────────────────────────────────────────────

  function getDefects(log) {
    // Geotab API casing quirk: property may be dVIRDefects, dvirDefects, or DVIRDefects
    var list = log.dVIRDefects || log.dvirDefects || log.DVIRDefects || [];
    if (!Array.isArray(list)) return [];
    return list;
  }

  function getRepairStatus(defect) {
    // DVIRDefect.repairStatus is a string: "NotRepaired", "NotNecessary", or "Repaired"
    var status = defect.repairStatus || "";
    if (typeof status === "string") return status;
    return "";
  }

  function isOutstanding(log) {
    var defects = getDefects(log);
    for (var i = 0; i < defects.length; i++) {
      if (getRepairStatus(defects[i]) === "NotRepaired") return true;
    }
    return false;
  }

  function hasNotNecessary(log) {
    var defects = getDefects(log);
    for (var i = 0; i < defects.length; i++) {
      if (getRepairStatus(defects[i]) === "NotNecessary") return true;
    }
    return false;
  }

  function hasDefects(log) {
    return getDefects(log).length > 0;
  }

  function getDeviceName(log) {
    if (log.device && log.device.id && deviceMap[log.device.id]) {
      return deviceMap[log.device.id].name || log.device.id;
    }
    if (log.device && log.device.name) return log.device.name;
    if (log.device && log.device.id) return log.device.id;
    return "--";
  }

  function getDriverName(log) {
    var driverId = null;
    if (log.driver && log.driver.id) driverId = log.driver.id;

    if (driverId && driverMap[driverId]) {
      var d = driverMap[driverId];
      var name = (d.firstName || "") + " " + (d.lastName || "");
      return name.trim() || d.name || driverId;
    }
    if (log.driver && log.driver.name) return log.driver.name;
    if (driverId && driverId !== "UnknownDriverId") return driverId;
    return "--";
  }

  function getLogType(log) {
    // logType: "PreTrip", "PostTrip", or other
    return log.logType || log.type || "--";
  }

  // ── Build Rows ─────────────────────────────────────────────────────────

  function buildFleetRows(logs) {
    var deviceIds = filteredDeviceIds();

    return logs.filter(function (log) {
      // Filter to selected devices
      var did = log.device ? log.device.id : null;
      if (did && !deviceIds[did]) return false;
      return true;
    }).map(function (log) {
      var defects = getDefects(log);
      var outstanding = 0, notNecessary = 0, repaired = 0;

      defects.forEach(function (d) {
        var status = getRepairStatus(d);
        if (status === "NotRepaired") outstanding++;
        else if (status === "NotNecessary") notNecessary++;
        else if (status === "Repaired") repaired++;
      });

      return {
        id: log.id,
        vehicle: getDeviceName(log),
        deviceId: log.device ? log.device.id : null,
        driver: getDriverName(log),
        date: log.dateTime || log.logDate,
        logType: getLogType(log),
        safeToOperate: log.isSafeToOperate !== false,
        totalDefects: defects.length,
        outstandingDefects: outstanding,
        notNecessary: notNecessary,
        repaired: repaired
      };
    });
  }

  function buildDefectRows(logs) {
    var deviceIds = filteredDeviceIds();
    var rows = [];

    logs.forEach(function (log) {
      var did = log.device ? log.device.id : null;
      if (did && !deviceIds[did]) return;

      var defects = getDefects(log);
      defects.forEach(function (defect) {
        var status = getRepairStatus(defect);
        var statusLabel, statusKey;
        if (status === "NotRepaired") { statusLabel = "Outstanding"; statusKey = "outstanding"; }
        else if (status === "NotNecessary") { statusLabel = "Not Necessary"; statusKey = "notNecessary"; }
        else if (status === "Repaired") { statusLabel = "Repaired"; statusKey = "repaired"; }
        else { statusLabel = status || "--"; statusKey = "other"; }

        // Get defect part and description
        var part = "--";
        var defectName = "--";
        var severity = "--";

        if (defect.defect) {
          defectName = defect.defect.name || defect.defect.description || "--";
          severity = defect.defect.severity || "--";
        }
        if (defect.part) {
          part = typeof defect.part === "object" ? (defect.part.name || "--") : (defect.part || "--");
        }

        // Repair details — API uses "repairUser" not "repairedBy"
        var repairedBy = "--";
        var repairDate = null;
        var repairUserObj = defect.repairUser || null;
        if (repairUserObj) {
          if (typeof repairUserObj === "object") {
            if (repairUserObj.id && driverMap[repairUserObj.id]) {
              var u = driverMap[repairUserObj.id];
              repairedBy = ((u.firstName || "") + " " + (u.lastName || "")).trim() || u.name || repairUserObj.id;
            } else {
              repairedBy = repairUserObj.name || repairUserObj.id || "--";
            }
          } else {
            repairedBy = repairUserObj;
          }
        }
        if (defect.repairDateTime) {
          repairDate = defect.repairDateTime;
        }

        // Remarks come from defectRemarks array
        var remarks = "--";
        if (Array.isArray(defect.defectRemarks) && defect.defectRemarks.length > 0) {
          remarks = defect.defectRemarks.map(function (r) {
            return r.remark || r.comment || r.text || "";
          }).filter(function (r) { return r; }).join("; ") || "--";
        }

        rows.push({
          dvirLogId: log.id,
          vehicle: getDeviceName(log),
          deviceId: log.device ? log.device.id : null,
          driver: getDriverName(log),
          date: log.dateTime || log.logDate,
          part: part,
          defect: defectName,
          severity: severity,
          repairStatus: statusLabel,
          repairStatusKey: statusKey,
          repairedBy: repairedBy,
          repairDate: repairDate,
          remarks: remarks
        });
      });
    });

    return rows;
  }

  // ── Rendering ──────────────────────────────────────────────────────────

  function renderActiveTab() {
    switch (activeTab) {
      case "fleet": renderFleetTable(); break;
      case "defects": renderDefectsTable(); break;
    }
  }

  function renderKpis() {
    var fleetRows = dvirData.fleetRows;
    var outstandingCount = 0;
    var notNecessaryCount = 0;

    fleetRows.forEach(function (r) {
      if (r.outstandingDefects > 0) outstandingCount++;
      if (r.notNecessary > 0) notNecessaryCount++;
    });

    els.kpiOutstanding.textContent = outstandingCount;
    els.kpiNotNecessary.textContent = notNecessaryCount;
  }

  function renderFleetTable() {
    var rows = dvirData.fleetRows.slice();
    var searchTerm = els.fleetSearch.value.toLowerCase();

    if (searchTerm) {
      rows = rows.filter(function (r) {
        return r.vehicle.toLowerCase().indexOf(searchTerm) >= 0 ||
               r.driver.toLowerCase().indexOf(searchTerm) >= 0;
      });
    }

    sortRows(rows, sortState.fleet);

    var frag = document.createDocumentFragment();
    rows.forEach(function (r) {
      var tr = document.createElement("tr");
      tr.className = "dvir-clickable-row";
      tr.dataset.dvirId = r.id;
      tr.dataset.deviceId = r.deviceId || "";

      var safeClass = r.safeToOperate ? "dvir-badge-safe" : "dvir-badge-unsafe";
      var safeText = r.safeToOperate ? "Yes" : "No";
      var outstandingClass = r.outstandingDefects > 0 ? ' class="dvir-outstanding-count"' : '';

      tr.innerHTML = '<td>' + escapeHtml(r.vehicle) + '</td>' +
        '<td>' + escapeHtml(r.driver) + '</td>' +
        '<td>' + formatDateTime(r.date) + '</td>' +
        '<td>' + escapeHtml(r.logType) + '</td>' +
        '<td><span class="' + safeClass + '">' + safeText + '</span></td>' +
        '<td>' + r.totalDefects + '</td>' +
        '<td' + outstandingClass + '>' + r.outstandingDefects + '</td>' +
        '<td>' + r.notNecessary + '</td>' +
        '<td>' + r.repaired + '</td>';

      tr.addEventListener("click", function () {
        var hash = "dvir,device:" + r.deviceId + ",id:" + r.id + ",trailer:!n";
        window.top.location.hash = hash;
      });

      frag.appendChild(tr);
    });
    els.fleetBody.innerHTML = "";
    els.fleetBody.appendChild(frag);

    if (rows.length === 0) {
      els.fleetBody.innerHTML = '<tr><td colspan="9" style="text-align:center;color:#888;padding:20px;">No DVIRs found.</td></tr>';
    }
  }

  function renderDefectsTable() {
    var rows = dvirData.defectRows.slice();
    var filterVal = els.defectFilter.value;
    var searchTerm = els.defectSearch.value.toLowerCase();

    // Apply repair status filter
    if (filterVal !== "all") {
      rows = rows.filter(function (r) { return r.repairStatusKey === filterVal; });
    }

    if (searchTerm) {
      rows = rows.filter(function (r) {
        return r.vehicle.toLowerCase().indexOf(searchTerm) >= 0 ||
               r.driver.toLowerCase().indexOf(searchTerm) >= 0 ||
               r.part.toLowerCase().indexOf(searchTerm) >= 0 ||
               r.defect.toLowerCase().indexOf(searchTerm) >= 0 ||
               r.remarks.toLowerCase().indexOf(searchTerm) >= 0;
      });
    }

    sortRows(rows, sortState.defects);

    var frag = document.createDocumentFragment();
    rows.forEach(function (r) {
      var tr = document.createElement("tr");
      tr.className = "dvir-clickable-row";

      var badgeClass = "dvir-badge ";
      if (r.repairStatusKey === "outstanding") badgeClass += "dvir-badge-outstanding";
      else if (r.repairStatusKey === "notNecessary") badgeClass += "dvir-badge-not-necessary";
      else if (r.repairStatusKey === "repaired") badgeClass += "dvir-badge-repaired";

      tr.innerHTML = '<td>' + escapeHtml(r.vehicle) + '</td>' +
        '<td>' + escapeHtml(r.driver) + '</td>' +
        '<td>' + formatDateTime(r.date) + '</td>' +
        '<td>' + escapeHtml(r.part) + '</td>' +
        '<td>' + escapeHtml(r.defect) + '</td>' +
        '<td>' + escapeHtml(r.severity) + '</td>' +
        '<td><span class="' + badgeClass + '">' + escapeHtml(r.repairStatus) + '</span></td>' +
        '<td>' + escapeHtml(r.repairedBy) + '</td>' +
        '<td>' + formatDate(r.repairDate) + '</td>' +
        '<td>' + escapeHtml(r.remarks) + '</td>';

      tr.addEventListener("click", function () {
        var hash = "dvir,device:" + r.deviceId + ",id:" + r.dvirLogId + ",trailer:!n";
        window.top.location.hash = hash;
      });

      frag.appendChild(tr);
    });
    els.defectBody.innerHTML = "";
    els.defectBody.appendChild(frag);

    if (rows.length === 0) {
      els.defectBody.innerHTML = '<tr><td colspan="10" style="text-align:center;color:#888;padding:20px;">No defects found.</td></tr>';
    }
  }

  // ── Table Utilities ────────────────────────────────────────────────────

  function renderTableBody(tbody, rows, cellFn) {
    var frag = document.createDocumentFragment();
    rows.forEach(function (r) {
      var tr = document.createElement("tr");
      tr.innerHTML = cellFn(r);
      frag.appendChild(tr);
    });
    tbody.innerHTML = "";
    tbody.appendChild(frag);
  }

  function sortRows(rows, state) {
    var col = state.col;
    var dir = state.dir === "asc" ? 1 : -1;

    rows.sort(function (a, b) {
      var va = a[col], vb = b[col];
      if (va == null) va = "";
      if (vb == null) vb = "";
      if (typeof va === "boolean" && typeof vb === "boolean") return (va === vb ? 0 : va ? -1 : 1) * dir;
      if (typeof va === "number" && typeof vb === "number") return (va - vb) * dir;
      if (typeof va === "string" && typeof vb === "string") return va.localeCompare(vb) * dir;
      return String(va).localeCompare(String(vb)) * dir;
    });
  }

  function handleSort(tableId, th) {
    var col = th.dataset.col;
    if (!col) return;
    var state = sortState[tableId];
    if (state.col === col) {
      state.dir = state.dir === "asc" ? "desc" : "asc";
    } else {
      state.col = col;
      state.dir = "asc";
    }

    // Update arrow indicators
    var table = th.closest("table");
    table.querySelectorAll(".dvir-sortable").forEach(function (h) {
      h.classList.remove("dvir-sort-asc", "dvir-sort-desc");
    });
    th.classList.add("dvir-sort-" + state.dir);

    // Re-render
    switch (tableId) {
      case "fleet": renderFleetTable(); break;
      case "defects": renderDefectsTable(); break;
    }
  }

  // ── CSV Export ──────────────────────────────────────────────────────────

  function exportCsv(filename, headers, rows) {
    var lines = [headers.join(",")];
    rows.forEach(function (r) {
      var vals = headers.map(function (h) {
        var v = r[h] != null ? String(r[h]) : "";
        if (v.indexOf(",") >= 0 || v.indexOf('"') >= 0 || v.indexOf("\n") >= 0) {
          v = '"' + v.replace(/"/g, '""') + '"';
        }
        return v;
      });
      lines.push(vals.join(","));
    });

    var blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ── Main Load (Apply) ──────────────────────────────────────────────────

  function loadData() {
    if (abortController) abortController.abort();
    abortController = new AbortController();

    showLoading(true, "Fetching DVIR data...");
    showEmpty(false);
    showWarning(null);
    setProgress(0);

    var dateRange = getDateRange();

    els.progress.textContent = "Loading...";

    fetchDVIRLogs(dateRange, function (pct) {
      setProgress(pct * 0.8);
      els.loadingText.textContent = "Fetching DVIR data... " + Math.round(pct) + "%";
    }).then(function (logs) {
      if (isAborted()) return;

      dvirData.logs = logs;

      var logsWithDefects = logs.filter(function (l) { return getDefects(l).length > 0; });
      console.log("DVIR Dashboard:", logs.length, "DVIRLogs,", logsWithDefects.length, "with defects");

      els.loadingText.textContent = "Fetching driver info...";
      setProgress(85);

      // Collect driver IDs for lookup
      var driverIds = [];
      logs.forEach(function (log) {
        if (log.driver && log.driver.id && log.driver.id !== "UnknownDriverId") {
          driverIds.push(log.driver.id);
        }
        // Also collect repairUser IDs from defects
        var defects = getDefects(log);
        defects.forEach(function (defect) {
          if (defect.repairUser && typeof defect.repairUser === "object" && defect.repairUser.id) {
            driverIds.push(defect.repairUser.id);
          }
        });
      });

      return fetchDrivers(driverIds).then(function () {
        return logs;
      });
    }).then(function (logs) {
      if (isAborted()) return;
      if (!logs) return;

      els.loadingText.textContent = "Processing data...";
      setProgress(95);

      // Build rows
      dvirData.fleetRows = buildFleetRows(logs);
      dvirData.defectRows = buildDefectRows(logs);

      // Update KPIs
      renderKpis();

      // Update progress text
      var totalLogs = dvirData.fleetRows.length;
      var totalDefects = dvirData.defectRows.length;
      els.progress.textContent = totalLogs + " DVIRs" + (totalDefects > 0 ? ", " + totalDefects + " defects" : "");

      // Render active tab
      renderActiveTab();
      showLoading(false);

      if (totalLogs === 0) {
        showEmpty(true);
        els.empty.textContent = "No DVIRs found for the selected filters.";
      }
    }).catch(function (err) {
      if (!isAborted()) {
        console.error("DVIR Dashboard error:", err);
        showLoading(false);
        showEmpty(true);
        els.empty.textContent = "Error loading data. Please try again.";
      }
    });
  }

  // ── UI Event Handlers ──────────────────────────────────────────────────

  function onPresetClick(e) {
    var btn = e.target.closest(".dvir-preset");
    if (!btn) return;

    document.querySelectorAll(".dvir-preset").forEach(function (b) { b.classList.remove("active"); });
    btn.classList.add("active");
  }

  function onTabClick(e) {
    var btn = e.target.closest(".dvir-tab");
    if (!btn) return;

    document.querySelectorAll(".dvir-tab").forEach(function (t) { t.classList.remove("active"); });
    btn.classList.add("active");

    activeTab = btn.dataset.tab;

    // Show/hide panels
    document.querySelectorAll(".dvir-panel").forEach(function (p) { p.classList.remove("active"); });
    var panel = $("dvir-panel-" + activeTab);
    if (panel) panel.classList.add("active");

    // Show/hide KPI strip (only on fleet tab)
    els.kpiStrip.style.display = activeTab === "fleet" ? "flex" : "none";

    // Re-render active tab
    if (dvirData.fleetRows.length > 0 || dvirData.defectRows.length > 0) {
      renderActiveTab();
    }
  }

  // ── Add-In Lifecycle ───────────────────────────────────────────────────

  return {
    initialize: function (freshApi, state, callback) {
      api = freshApi;

      // Cache DOM refs
      els.group = $("dvir-group");
      els.vehicle = $("dvir-vehicle");
      els.apply = $("dvir-apply");
      els.progress = $("dvir-progress");
      els.loading = $("dvir-loading");
      els.loadingText = $("dvir-loading-text");
      els.progressBar = $("dvir-progress-bar");
      els.empty = $("dvir-empty");
      els.warning = $("dvir-warning");
      els.kpiStrip = $("dvir-kpi-strip");
      els.kpiOutstanding = $("dvir-kpi-outstanding");
      els.kpiNotNecessary = $("dvir-kpi-not-necessary");
      els.fleetSearch = $("dvir-fleet-search");
      els.fleetBody = $("dvir-fleet-body");
      els.defectFilter = $("dvir-defect-filter");
      els.defectSearch = $("dvir-defect-search");
      els.defectBody = $("dvir-defect-body");

      // Event listeners
      els.apply.addEventListener("click", loadData);
      document.querySelector(".dvir-presets").addEventListener("click", onPresetClick);
      $("dvir-tabs").addEventListener("click", onTabClick);

      // Table sort listeners
      $("dvir-fleet-table").addEventListener("click", function (e) {
        var th = e.target.closest(".dvir-sortable");
        if (th) handleSort("fleet", th);
      });
      $("dvir-defect-table").addEventListener("click", function (e) {
        var th = e.target.closest(".dvir-sortable");
        if (th) handleSort("defects", th);
      });

      // Search / filter listeners
      els.fleetSearch.addEventListener("input", renderFleetTable);
      els.defectFilter.addEventListener("change", renderDefectsTable);
      els.defectSearch.addEventListener("input", renderDefectsTable);

      // CSV export listeners
      $("dvir-fleet-export").addEventListener("click", function () {
        var headers = ["vehicle", "driver", "date", "logType", "safeToOperate", "totalDefects", "outstandingDefects", "notNecessary", "repaired"];
        exportCsv("dvir_fleet_summary.csv", headers, dvirData.fleetRows);
      });
      $("dvir-defect-export").addEventListener("click", function () {
        var headers = ["vehicle", "driver", "date", "part", "defect", "severity", "repairStatus", "repairedBy", "repairDate", "remarks"];
        exportCsv("dvir_defect_detail.csv", headers, dvirData.defectRows);
      });

      // Load foundation data: Devices + Groups
      apiMultiCall([
        ["Get", { typeName: "Device", resultsLimit: 5000 }],
        ["Get", { typeName: "Group", resultsLimit: 5000 }]
      ]).then(function (results) {
        var now = new Date();
        allDevices = (results[0] || []).filter(function (d) {
          if (!d.activeTo) return true;
          return new Date(d.activeTo) > now;
        });

        var groups = results[1] || [];

        // Build device map
        allDevices.forEach(function (d) {
          deviceMap[d.id] = d;
        });

        // Build group map
        groups.forEach(function (g) {
          allGroups[g.id] = g;
        });

        populateGroupDropdown();
        populateVehicleDropdown();
        callback();
      }).catch(function (err) {
        console.error("DVIR Dashboard init error:", err);
        callback();
      });
    },

    focus: function (freshApi, state) {
      api = freshApi;

      // Refresh devices
      apiCall("Get", { typeName: "Device", resultsLimit: 5000 }).then(function (devices) {
        var now = new Date();
        allDevices = (devices || []).filter(function (d) {
          if (!d.activeTo) return true;
          return new Date(d.activeTo) > now;
        });
        allDevices.forEach(function (d) {
          deviceMap[d.id] = d;
        });
        populateVehicleDropdown();
      }).catch(function () {});

      // Auto-load on first focus
      if (firstFocus) {
        firstFocus = false;
        loadData();
      }
    },

    blur: function () {
      if (abortController) {
        abortController.abort();
        abortController = null;
      }
      showLoading(false);
    }
  };
};
