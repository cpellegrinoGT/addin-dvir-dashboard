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
  var groupPicker = null;
  var vehiclePicker = null;
  var FLEET_ROW_LIMIT = 100;

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

  function dvirLink(dvirId, deviceId, label) {
    return '<a href="#" class="dvir-log-link" data-dvir-id="' + escapeHtml(dvirId) + '" data-device-id="' + escapeHtml(deviceId || "") + '">' + escapeHtml(label) + '</a>';
  }

  function goToDvir(dvirId, deviceId) {
    var hash = "dvir,device:" + deviceId + ",id:" + dvirId + ",trailer:!n";
    window.top.location.hash = hash;
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
      case "custom":
        from = els.fromDate.value ? new Date(els.fromDate.value + "T00:00:00") : new Date(now.getTime() - 30 * 86400000);
        to = els.toDate.value ? new Date(els.toDate.value + "T23:59:59") : to;
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

  // ── Multi-Select Widget ───────────────────────────────────────────────

  function closeAllDropdowns() {
    document.querySelectorAll("#dvir-root .dvir-ms-dropdown.open").forEach(function (d) {
      d.classList.remove("open");
    });
  }

  function initMultiSelect(cfg) {
    var container = $(cfg.id);
    var toggle = container.querySelector(".dvir-ms-toggle");
    var dropdown = container.querySelector(".dvir-ms-dropdown");
    var searchInput = container.querySelector(".dvir-ms-search");
    var selectAllCb = container.querySelector(".dvir-ms-select-all input");
    var clearBtn = container.querySelector(".dvir-ms-clear");
    var listEl = container.querySelector(".dvir-ms-list");

    var items = [];
    var selected = new Set();

    function render(filter) {
      var filt = (filter || "").toLowerCase();
      listEl.innerHTML = "";
      var visibleCount = 0;
      var checkedCount = 0;

      var sorted = items.filter(function (item) {
        return !filt || item.label.toLowerCase().indexOf(filt) >= 0;
      });
      sorted.sort(function (a, b) {
        var aChecked = selected.has(a.value) ? 0 : 1;
        var bChecked = selected.has(b.value) ? 0 : 1;
        if (aChecked !== bChecked) return aChecked - bChecked;
        return a.label.localeCompare(b.label);
      });

      sorted.forEach(function (item) {
        visibleCount++;
        var checked = selected.has(item.value);
        if (checked) checkedCount++;
        var label = document.createElement("label");
        label.className = "dvir-ms-item";
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = checked;
        cb.addEventListener("change", function () {
          if (cb.checked) selected.add(item.value);
          else selected.delete(item.value);
          updateToggleText();
          updateSelectAll();
          render(searchInput.value);
        });
        var span = document.createElement("span");
        span.textContent = item.label;
        label.appendChild(cb);
        label.appendChild(span);
        listEl.appendChild(label);
      });
      selectAllCb.checked = visibleCount > 0 && checkedCount === visibleCount;
    }

    function updateToggleText() {
      if (selected.size === 0) {
        toggle.textContent = cfg.placeholder || "Select...";
      } else if (selected.size <= 2) {
        var labels = [];
        items.forEach(function (it) {
          if (selected.has(it.value)) labels.push(it.label);
        });
        toggle.textContent = labels.join(", ");
      } else {
        toggle.textContent = selected.size + " selected";
      }
    }

    function updateSelectAll() {
      var filt = (searchInput.value || "").toLowerCase();
      var visibleCount = 0;
      var checkedCount = 0;
      items.forEach(function (item) {
        if (filt && item.label.toLowerCase().indexOf(filt) < 0) return;
        visibleCount++;
        if (selected.has(item.value)) checkedCount++;
      });
      selectAllCb.checked = visibleCount > 0 && checkedCount === visibleCount;
    }

    toggle.addEventListener("click", function (e) {
      e.stopPropagation();
      var isOpen = dropdown.classList.contains("open");
      closeAllDropdowns();
      if (!isOpen) {
        dropdown.classList.add("open");
        searchInput.value = "";
        render("");
        searchInput.focus();
      }
    });

    searchInput.addEventListener("input", function () {
      render(searchInput.value);
    });

    searchInput.addEventListener("click", function (e) { e.stopPropagation(); });

    selectAllCb.addEventListener("change", function () {
      var filt = (searchInput.value || "").toLowerCase();
      items.forEach(function (item) {
        if (filt && item.label.toLowerCase().indexOf(filt) < 0) return;
        if (selectAllCb.checked) selected.add(item.value);
        else selected.delete(item.value);
      });
      render(searchInput.value);
      updateToggleText();
    });

    clearBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      selected.clear();
      selectAllCb.checked = false;
      render(searchInput.value);
      updateToggleText();
    });

    dropdown.addEventListener("click", function (e) { e.stopPropagation(); });

    return {
      setItems: function (newItems) {
        items = newItems.slice().sort(function (a, b) {
          return a.label.localeCompare(b.label);
        });
        selected.clear();
        updateToggleText();
        render("");
      },
      getSelected: function () {
        return Array.from(selected);
      },
      hasSelection: function () {
        return selected.size > 0;
      },
      container: container,
      dropdown: dropdown
    };
  }

  // ── Dropdown Population ────────────────────────────────────────────────

  function populateGroupDropdown() {
    var skipIds = { GroupCompanyId: true, GroupNothingId: true };
    var items = [];
    Object.keys(allGroups).forEach(function (gid) {
      var g = allGroups[gid];
      if (skipIds[gid]) return;
      if (!g.name || g.name === "CompanyGroup" || g.name === "**Nothing**") return;
      items.push({ value: g.id, label: g.name || g.id });
    });
    groupPicker.setItems(items);
  }

  function populateVehicleDropdown() {
    var items = allDevices.map(function (d) {
      return { value: d.id, label: d.name || d.id };
    });
    vehiclePicker.setItems(items);
  }

  // ── Filtered Devices ───────────────────────────────────────────────────

  function filteredDeviceIds() {
    var selectedVehicles = vehiclePicker.getSelected();
    var selectedGroups = groupPicker.getSelected();

    // If specific vehicles selected, use those directly
    if (selectedVehicles.length > 0) {
      var set = {};
      selectedVehicles.forEach(function (vid) { set[vid] = true; });
      return set;
    }

    // Otherwise filter by groups (empty = all)
    var set = {};
    var groupSet = {};
    if (selectedGroups.length > 0) {
      selectedGroups.forEach(function (gid) { groupSet[gid] = true; });
    }

    allDevices.forEach(function (dev) {
      if (selectedGroups.length > 0) {
        var devGroups = dev.groups || [];
        var inGroup = false;
        for (var i = 0; i < devGroups.length; i++) {
          if (groupSet[devGroups[i].id]) { inGroup = true; break; }
        }
        if (!inGroup) return;
      }
      set[dev.id] = true;
    });
    return set;
  }

  // ── DVIR Data Fetch ────────────────────────────────────────────────────

  // Primary: GetFeed returns full DVIRLog objects with dVIRDefects populated.
  // Each GetFeed call counts as 1 API call regardless of result count.
  function fetchDVIRLogsViaFeed(dateRange, onProgress) {
    var LIMIT = 5000;
    var toMs = new Date(dateRange.to).getTime();
    var allLogs = [];
    var fromVersion = null;

    function nextPage() {
      if (isAborted()) return Promise.resolve(allLogs);

      var params = {
        typeName: "DVIRLog",
        search: { fromDate: dateRange.from },
        resultsLimit: LIMIT
      };
      if (fromVersion) {
        params.fromVersion = fromVersion;
      }

      return apiCall("GetFeed", params).then(function (feed) {
        if (!feed) return allLogs;

        var data = feed.data || [];
        fromVersion = feed.toVersion || null;

        // Filter to within our date range (GetFeed only supports fromDate)
        data.forEach(function (log) {
          var logTime = new Date(log.dateTime).getTime();
          if (logTime <= toMs) {
            allLogs.push(log);
          }
        });

        if (onProgress) onProgress(allLogs.length);

        // If we got a full page AND haven't passed our toDate, fetch more
        if (data.length >= LIMIT && fromVersion) {
          var lastLogTime = data.length > 0 ? new Date(data[data.length - 1].dateTime).getTime() : 0;
          if (lastLogTime <= toMs) {
            return delay(200).then(nextPage);
          }
        }

        return allLogs;
      });
    }

    return nextPage();
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

  function getDeviceGroups(log) {
    var deviceId = log.device ? log.device.id : null;
    if (!deviceId || !deviceMap[deviceId]) return "";
    var dev = deviceMap[deviceId];
    var devGroups = dev.groups || [];
    var names = [];
    for (var i = 0; i < devGroups.length; i++) {
      var gid = devGroups[i].id;
      if (gid === "GroupCompanyId" || gid === "GroupNothingId") continue;
      var g = allGroups[gid];
      if (g && g.name && g.name !== "CompanyGroup" && g.name !== "**Nothing**") {
        names.push(g.name);
      }
    }
    return names.join(", ");
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
        groups: getDeviceGroups(log),
        date: log.dateTime || log.logDate,
        logType: getLogType(log),
        safeToOperate: log.isSafeToOperate === true ? "yes"
                     : log.isSafeToOperate === false ? "no"
                     : "unknown",
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
          groups: getDeviceGroups(log),
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
    var logTypeFilter = els.fleetLogType.value;
    var safeFilter = els.fleetSafe.value;

    if (searchTerm) {
      rows = rows.filter(function (r) {
        return r.vehicle.toLowerCase().indexOf(searchTerm) >= 0 ||
               r.driver.toLowerCase().indexOf(searchTerm) >= 0 ||
               r.groups.toLowerCase().indexOf(searchTerm) >= 0;
      });
    }

    if (logTypeFilter !== "all") {
      rows = rows.filter(function (r) { return r.logType === logTypeFilter; });
    }

    if (safeFilter !== "all") {
      rows = rows.filter(function (r) { return r.safeToOperate === safeFilter; });
    }

    sortRows(rows, sortState.fleet);

    // Keep full filtered set for CSV export, limit DOM rendering
    var totalFiltered = rows.length;
    var displayRows = rows.length > FLEET_ROW_LIMIT ? rows.slice(0, FLEET_ROW_LIMIT) : rows;

    renderTableBody(els.fleetBody, displayRows, function (r) {
      var safeClass, safeText;
      if (r.safeToOperate === "yes") { safeClass = "dvir-badge-safe"; safeText = "Yes"; }
      else if (r.safeToOperate === "no") { safeClass = "dvir-badge-unsafe"; safeText = "No"; }
      else { safeClass = "dvir-badge-unknown"; safeText = "Unknown"; }
      var outstandingClass = r.outstandingDefects > 0 ? ' class="dvir-outstanding-count"' : '';

      return '<td>' + dvirLink(r.id, r.deviceId, r.vehicle) + '</td>' +
        '<td>' + escapeHtml(r.driver) + '</td>' +
        '<td>' + escapeHtml(r.groups) + '</td>' +
        '<td>' + formatDateTime(r.date) + '</td>' +
        '<td>' + escapeHtml(r.logType) + '</td>' +
        '<td><span class="' + safeClass + '">' + safeText + '</span></td>' +
        '<td>' + r.totalDefects + '</td>' +
        '<td' + outstandingClass + '>' + r.outstandingDefects + '</td>' +
        '<td>' + r.notNecessary + '</td>' +
        '<td>' + r.repaired + '</td>';
    });

    // Show row limit indicator
    var limitMsg = els.fleetBody.parentElement.parentElement.querySelector(".dvir-row-limit-msg");
    if (!limitMsg) {
      limitMsg = document.createElement("div");
      limitMsg.className = "dvir-row-limit-msg";
      els.fleetBody.parentElement.parentElement.appendChild(limitMsg);
    }
    if (totalFiltered > FLEET_ROW_LIMIT) {
      limitMsg.textContent = "Showing " + FLEET_ROW_LIMIT + " of " + totalFiltered + " DVIRs";
      limitMsg.style.display = "";
    } else {
      limitMsg.style.display = "none";
    }

    if (rows.length === 0) {
      els.fleetBody.innerHTML = '<tr><td colspan="10" style="text-align:center;color:#888;padding:20px;">No DVIRs found.</td></tr>';
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
               r.groups.toLowerCase().indexOf(searchTerm) >= 0 ||
               r.part.toLowerCase().indexOf(searchTerm) >= 0 ||
               r.defect.toLowerCase().indexOf(searchTerm) >= 0 ||
               r.remarks.toLowerCase().indexOf(searchTerm) >= 0;
      });
    }

    sortRows(rows, sortState.defects);
    renderTableBody(els.defectBody, rows, function (r) {
      var badgeClass = "dvir-badge ";
      if (r.repairStatusKey === "outstanding") badgeClass += "dvir-badge-outstanding";
      else if (r.repairStatusKey === "notNecessary") badgeClass += "dvir-badge-not-necessary";
      else if (r.repairStatusKey === "repaired") badgeClass += "dvir-badge-repaired";

      return '<td>' + dvirLink(r.dvirLogId, r.deviceId, r.vehicle) + '</td>' +
        '<td>' + escapeHtml(r.driver) + '</td>' +
        '<td>' + escapeHtml(r.groups) + '</td>' +
        '<td>' + formatDateTime(r.date) + '</td>' +
        '<td>' + escapeHtml(r.part) + '</td>' +
        '<td>' + escapeHtml(r.defect) + '</td>' +
        '<td>' + escapeHtml(r.severity) + '</td>' +
        '<td><span class="' + badgeClass + '">' + escapeHtml(r.repairStatus) + '</span></td>' +
        '<td>' + escapeHtml(r.repairedBy) + '</td>' +
        '<td>' + formatDate(r.repairDate) + '</td>' +
        '<td>' + escapeHtml(r.remarks) + '</td>';
    });

    if (rows.length === 0) {
      els.defectBody.innerHTML = '<tr><td colspan="11" style="text-align:center;color:#888;padding:20px;">No defects found.</td></tr>';
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

    showLoading(true, "Fetching DVIR inspections...");
    showEmpty(false);
    showWarning(null);
    setProgress(0);

    var dateRange = getDateRange();

    els.progress.textContent = "Loading...";

    // Single-phase fetch via GetFeed — returns full DVIRLog objects with dVIRDefects populated
    fetchDVIRLogsViaFeed(dateRange, function (count) {
      els.loadingText.textContent = "Fetching inspections... " + count + " so far";
    }).then(function (logs) {
      if (isAborted()) return;

      console.log("DVIR Dashboard: GetFeed complete —", logs.length, "DVIRLogs");

      var logsWithDefects = logs.filter(function (l) { return getDefects(l).length > 0; });
      console.log("DVIR Dashboard:", logsWithDefects.length, "DVIRs have defects");

      // Build rows and render (driver names already resolved from init)
      dvirData.logs = logs;
      dvirData.fleetRows = buildFleetRows(logs);
      dvirData.defectRows = buildDefectRows(logs);
      renderKpis();
      renderActiveTab();
      showLoading(false);

      if (logs.length === 0) {
        showEmpty(true);
        els.empty.textContent = "No DVIRs found for the selected filters.";
        els.progress.textContent = "";
        return;
      }

      var totalDefects = dvirData.defectRows.length;
      els.progress.textContent = dvirData.fleetRows.length + " DVIRs, " + totalDefects + " defects";
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

    var isCustom = btn.dataset.preset === "custom";
    els.customDates.style.display = isCustom ? "" : "none";

    if (isCustom && !els.fromDate.value) {
      var now = new Date();
      var from = new Date(now);
      from.setDate(from.getDate() - 30);
      els.fromDate.value = from.toISOString().slice(0, 10);
      els.toDate.value = now.toISOString().slice(0, 10);
    }
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
      els.customDates = $("dvir-custom-dates");
      els.fromDate = $("dvir-from");
      els.toDate = $("dvir-to");
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
      els.fleetLogType = $("dvir-fleet-logtype");
      els.fleetSafe = $("dvir-fleet-safe");
      els.fleetSearch = $("dvir-fleet-search");
      els.fleetBody = $("dvir-fleet-body");
      els.defectFilter = $("dvir-defect-filter");
      els.defectSearch = $("dvir-defect-search");
      els.defectBody = $("dvir-defect-body");

      // Init multi-select widgets
      groupPicker = initMultiSelect({ id: "dvir-group", placeholder: "All Groups" });
      vehiclePicker = initMultiSelect({ id: "dvir-vehicle", placeholder: "All Vehicles" });

      // Close dropdowns on outside click
      document.addEventListener("click", closeAllDropdowns);

      // Event listeners
      els.apply.addEventListener("click", loadData);
      document.querySelector(".dvir-presets").addEventListener("click", onPresetClick);
      $("dvir-tabs").addEventListener("click", onTabClick);

      // DVIR link click handler (delegated on content area)
      $("dvir-content").addEventListener("click", function (e) {
        var link = e.target.closest(".dvir-log-link");
        if (link) {
          e.preventDefault();
          goToDvir(link.dataset.dvirId, link.dataset.deviceId);
        }
      });

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
      els.fleetLogType.addEventListener("change", renderFleetTable);
      els.fleetSafe.addEventListener("change", renderFleetTable);
      els.fleetSearch.addEventListener("input", renderFleetTable);
      els.defectFilter.addEventListener("change", renderDefectsTable);
      els.defectSearch.addEventListener("input", renderDefectsTable);

      // CSV export listeners
      $("dvir-fleet-export").addEventListener("click", function () {
        var headers = ["vehicle", "driver", "groups", "date", "logType", "safeToOperate", "totalDefects", "outstandingDefects", "notNecessary", "repaired"];
        exportCsv("dvir_fleet_summary.csv", headers, dvirData.fleetRows);
      });
      $("dvir-defect-export").addEventListener("click", function () {
        var headers = ["vehicle", "driver", "groups", "date", "part", "defect", "severity", "repairStatus", "repairedBy", "repairDate", "remarks"];
        exportCsv("dvir_defect_detail.csv", headers, dvirData.defectRows);
      });

      // Load foundation data: Devices + Groups + Users (3 API calls total)
      apiMultiCall([
        ["Get", { typeName: "Device", resultsLimit: 5000 }],
        ["Get", { typeName: "Group", resultsLimit: 5000 }],
        ["Get", { typeName: "User", resultsLimit: 50000 }]
      ]).then(function (results) {
        var now = new Date();
        allDevices = (results[0] || []).filter(function (d) {
          if (!d.activeTo) return true;
          return new Date(d.activeTo) > now;
        });

        var groups = results[1] || [];
        var users = results[2] || [];

        // Build device map
        allDevices.forEach(function (d) {
          deviceMap[d.id] = d;
        });

        // Build group map
        groups.forEach(function (g) {
          allGroups[g.id] = g;
        });

        // Build driver/user map
        users.forEach(function (u) {
          driverMap[u.id] = u;
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
