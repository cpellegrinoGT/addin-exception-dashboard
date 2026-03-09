geotab.addin.exceptionDashboard = function () {
  "use strict";

  var api;
  var abortController = null;
  var chartInstance = null;

  // ─── Settings Manager (localStorage) ───

  var settings = (function () {
    var STORAGE_PREFIX = "exd_settings_v1_";
    var storageKey = "";
    var data = { defaultRules: [], ruleColors: {}, views: [] };

    function persist() {
      try { localStorage.setItem(storageKey, JSON.stringify(data)); } catch (e) { /* quota */ }
    }

    return {
      init: function (database) {
        storageKey = STORAGE_PREFIX + (database || location.hostname);
        try {
          var raw = localStorage.getItem(storageKey);
          if (raw) {
            var parsed = JSON.parse(raw);
            data.defaultRules = parsed.defaultRules || [];
            data.ruleColors = parsed.ruleColors || {};
            data.views = parsed.views || [];
          }
        } catch (e) { /* corrupt data, use defaults */ }
      },
      get: function () { return data; },
      setDefaultRules: function (ids) {
        data.defaultRules = ids.slice();
        persist();
      },
      setRuleColor: function (ruleId, color) {
        data.ruleColors[ruleId] = color;
        persist();
      },
      clearRuleColor: function (ruleId) {
        delete data.ruleColors[ruleId];
        persist();
      },
      saveView: function (view) {
        data.views.push(view);
        persist();
      },
      deleteView: function (viewId) {
        data.views = data.views.filter(function (v) { return v.id !== viewId; });
        persist();
      }
    };
  })();

  // Reference data
  var allRules = [];
  var allGroups = [];
  var allDevices = [];
  var deviceGroupMap = {}; // deviceId -> [groupId, ...]
  var groupChildrenMap = {}; // groupId -> [direct child groupIds]

  // Current results
  var currentRows = [];
  var currentHeaders = [];
  var sortState = { col: null, dir: "desc" };

  // Multi-select instances
  var rulePicker, groupPicker;

  // DOM cache
  var els = {};

  // ─── Helpers ───

  function $(id) { return document.getElementById(id); }

  function delay(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
  }

  function formatDate(d) {
    return d.getFullYear() + "-" +
      String(d.getMonth() + 1).padStart(2, "0") + "-" +
      String(d.getDate()).padStart(2, "0");
  }

  function parseLocalDate(str) {
    var parts = str.split("-");
    return new Date(+parts[0], +parts[1] - 1, +parts[2]);
  }

  function toISODate(d) {
    return d.toISOString();
  }

  function isAborted() {
    return abortController && abortController.signal && abortController.signal.aborted;
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

  // ─── Group Hierarchy ───

  function buildGroupChildrenMap(groups) {
    groupChildrenMap = {};
    groups.forEach(function (g) {
      if (g.children && g.children.length) {
        groupChildrenMap[g.id] = g.children.map(function (c) { return c.id; });
      }
    });
  }

  function getDescendantIds(groupId) {
    // Returns a Set containing groupId and all its descendants
    var result = new Set();
    var stack = [groupId];
    while (stack.length) {
      var gid = stack.pop();
      if (result.has(gid)) continue;
      result.add(gid);
      var children = groupChildrenMap[gid];
      if (children) {
        children.forEach(function (cid) { stack.push(cid); });
      }
    }
    return result;
  }

  // ─── Multi-Select Widget ───

  function initMultiSelect(cfg) {
    var container = $(cfg.id);
    var toggle = container.querySelector(".exd-ms-toggle");
    var dropdown = container.querySelector(".exd-ms-dropdown");
    var searchInput = container.querySelector(".exd-ms-search");
    var selectAllCb = container.querySelector(".exd-ms-select-all input");
    var clearBtn = container.querySelector(".exd-ms-clear");
    var listEl = container.querySelector(".exd-ms-list");

    var items = [];
    var selected = new Set();
    var getColorFn = cfg.getColor || null; // optional callback(value) -> color|null

    function render(filter) {
      var filt = (filter || "").toLowerCase();
      listEl.innerHTML = "";
      var visibleCount = 0;
      var checkedCount = 0;

      // Sort: selected items first, then alphabetical
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
        label.className = "exd-ms-item";
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = checked;
        cb.addEventListener("change", function () {
          if (cb.checked) selected.add(item.value);
          else selected.delete(item.value);
          updateToggleText();
          updateSelectAll();
          render(searchInput.value); // re-render to re-sort
        });
        label.appendChild(cb);
        // Color dot (if getColor callback returns a color for this item)
        if (getColorFn) {
          var dotColor = getColorFn(item.value);
          if (dotColor) {
            var dot = document.createElement("span");
            dot.className = "exd-ms-color-dot";
            dot.style.backgroundColor = dotColor;
            label.appendChild(dot);
          }
        }
        var span = document.createElement("span");
        span.textContent = item.label;
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
      setSelected: function (ids) {
        selected.clear();
        var validValues = new Set(items.map(function (it) { return it.value; }));
        ids.forEach(function (id) {
          if (validValues.has(id)) selected.add(id);
        });
        updateToggleText();
        render(searchInput.value);
      },
      refresh: function () {
        render(searchInput.value);
      },
      getItems: function () {
        return items.slice();
      },
      container: container,
      dropdown: dropdown
    };
  }

  function closeAllDropdowns() {
    document.querySelectorAll("#exd-root .exd-ms-dropdown.open").forEach(function (d) {
      d.classList.remove("open");
    });
  }

  // ─── Loading / UI helpers ───

  function showLoading(show, text) {
    els.loading.style.display = show ? "flex" : "none";
    if (text) els.loadingText.textContent = text;
  }

  function setProgress(pct) {
    els.progressFill.style.width = Math.min(100, Math.round(pct)) + "%";
  }

  function showEmpty(show) {
    els.empty.style.display = show ? "flex" : "none";
  }

  function showWarning(msg) {
    if (msg) {
      els.warning.textContent = msg;
      els.warning.style.display = "block";
    } else {
      els.warning.style.display = "none";
    }
  }

  function setStatus(text) {
    els.status.textContent = text || "";
  }

  // ─── Reference Data Loading ───

  function loadReferenceData() {
    showLoading(true, "Loading rules, groups, and devices...");
    setProgress(0);
    return apiMultiCall([
      ["Get", { typeName: "Rule", resultsLimit: 10000 }],
      ["Get", { typeName: "Group", resultsLimit: 10000 }],
      ["Get", { typeName: "Device", resultsLimit: 50000 }]
    ]).then(function (results) {
      allRules = results[0] || [];
      allGroups = results[1] || [];
      allDevices = results[2] || [];

      // Build group hierarchy map
      buildGroupChildrenMap(allGroups);

      // Build device-to-groups map
      deviceGroupMap = {};
      allDevices.forEach(function (d) {
        if (d.groups && d.groups.length) {
          deviceGroupMap[d.id] = d.groups.map(function (g) { return g.id; });
        }
      });

      // Populate rule picker
      var ruleItems = allRules
        .filter(function (r) { return r.name && r.name !== ""; })
        .map(function (r) {
          return { value: r.id, label: r.name };
        });
      rulePicker.setItems(ruleItems);

      // Apply default rules from settings
      var defaultRuleIds = settings.get().defaultRules;
      if (defaultRuleIds.length > 0) {
        rulePicker.setSelected(defaultRuleIds);
      }

      // Populate group picker — filter to non-system groups
      var systemGroupIds = new Set([
        "GroupCompanyId", "GroupRootId", "GroupNothingId",
        "GroupSecurityId", "GroupEverythingId", "GroupPrivateUserId"
      ]);
      var groupItems = allGroups
        .filter(function (g) {
          return g.name && !systemGroupIds.has(g.id) && g.id.indexOf("Group") !== 0;
        })
        .map(function (g) {
          return { value: g.id, label: g.name };
        });
      groupPicker.setItems(groupItems);

      showLoading(false);
      setStatus(allRules.length + " rules loaded");
    }).catch(function (err) {
      showLoading(false);
      setStatus("Error loading reference data: " + (err.message || err));
      console.error("loadReferenceData error:", err);
    });
  }

  // ─── Date Range Splitting ───

  function splitDateRange(from, to, maxDays) {
    var chunks = [];
    var cursor = from.getTime();
    var end = to.getTime();
    var step = maxDays * 86400000;
    while (cursor < end) {
      var chunkEnd = Math.min(cursor + step, end);
      chunks.push({ from: new Date(cursor), to: new Date(chunkEnd) });
      cursor = chunkEnd;
    }
    return chunks;
  }

  // ─── Fetch Exception Events ───

  function fetchExceptionEvents(ruleIds, fromDate, toDate, onProgress) {
    var MAX_DAYS = 14;
    var BATCH_SIZE = 20;
    var RESULT_LIMIT = 50000;
    var chunks = splitDateRange(fromDate, toDate, MAX_DAYS);

    // Build calls: one Get per rule per chunk
    var allCalls = [];
    ruleIds.forEach(function (ruleId) {
      chunks.forEach(function (chunk) {
        allCalls.push(["Get", {
          typeName: "ExceptionEvent",
          search: {
            fromDate: toISODate(chunk.from),
            toDate: toISODate(chunk.to),
            ruleSearch: { id: ruleId }
          },
          resultsLimit: RESULT_LIMIT
        }]);
      });
    });

    // Batch into groups of BATCH_SIZE for multiCall
    var batches = [];
    for (var i = 0; i < allCalls.length; i += BATCH_SIZE) {
      batches.push(allCalls.slice(i, i + BATCH_SIZE));
    }

    var allEvents = [];
    var completedBatches = 0;
    var totalBatches = batches.length;
    var hitLimit = false;

    return batches.reduce(function (chain, batch, idx) {
      return chain.then(function () {
        if (isAborted()) return;
        var pause = idx > 0 ? delay(200) : Promise.resolve();
        return pause.then(function () {
          if (isAborted()) return;
          return apiMultiCall(batch).then(function (results) {
            results.forEach(function (arr) {
              if (arr && arr.length) {
                if (arr.length >= RESULT_LIMIT) hitLimit = true;
                allEvents = allEvents.concat(arr);
              }
            });
            completedBatches++;
            if (onProgress) onProgress(completedBatches / totalBatches * 100);
          });
        });
      });
    }, Promise.resolve()).then(function () {
      return { events: allEvents, hitLimit: hitLimit };
    });
  }

  // ─── Fetch Trips (for mileage) ───

  function fetchTrips(fromDate, toDate, onProgress) {
    var MAX_DAYS = 14;
    var BATCH_SIZE = 10;
    var RESULT_LIMIT = 50000;
    var chunks = splitDateRange(fromDate, toDate, MAX_DAYS);

    var allCalls = chunks.map(function (chunk) {
      return ["Get", {
        typeName: "Trip",
        search: {
          fromDate: toISODate(chunk.from),
          toDate: toISODate(chunk.to)
        },
        resultsLimit: RESULT_LIMIT
      }];
    });

    var batches = [];
    for (var i = 0; i < allCalls.length; i += BATCH_SIZE) {
      batches.push(allCalls.slice(i, i + BATCH_SIZE));
    }

    var allTrips = [];
    var completedBatches = 0;
    var totalBatches = batches.length;

    return batches.reduce(function (chain, batch, idx) {
      return chain.then(function () {
        if (isAborted()) return;
        var pause = idx > 0 ? delay(200) : Promise.resolve();
        return pause.then(function () {
          if (isAborted()) return;
          return apiMultiCall(batch).then(function (results) {
            results.forEach(function (arr) {
              if (arr && arr.length) {
                allTrips = allTrips.concat(arr);
              }
            });
            completedBatches++;
            if (onProgress) onProgress(completedBatches / totalBatches * 100);
          });
        });
      });
    }, Promise.resolve()).then(function () {
      return allTrips;
    });
  }

  function aggregateMileageByPeriod(trips, mode) {
    // Returns { bucketKey -> totalKm } or null if no mileage data
    var mileage = {};
    var hasData = false;
    trips.forEach(function (trip) {
      if (!trip.start) return;
      var dist = trip.distance;
      if (typeof dist !== "number" || dist <= 0) return;
      var key = getBucketKey(trip.start, mode);
      if (!mileage[key]) mileage[key] = 0;
      mileage[key] += dist; // distance is in km
      hasData = true;
    });
    return hasData ? mileage : null;
  }

  // ─── Aggregation ───

  function getBucketKey(dateStr, mode) {
    var d = new Date(dateStr);
    if (mode === "day") {
      return formatDate(d);
    } else if (mode === "week") {
      // ISO week start (Monday)
      var day = d.getDay();
      var diff = d.getDate() - day + (day === 0 ? -6 : 1);
      var monday = new Date(d);
      monday.setDate(diff);
      return "W" + formatDate(monday);
    } else {
      // month
      return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
    }
  }

  function aggregateEvents(events, ruleIds, groupIds, mode, viewMode) {
    var ruleIdSet = new Set(ruleIds);
    var bucketsMap = {}; // key -> { seriesId -> count }
    var uniqueDevices = new Set();

    // For groups mode: pre-compute expanded descendant sets per selected group
    // Maps every descendant groupId back to its selected ancestor(s)
    var descendantToSelected = null;
    if (viewMode === "groups" && groupIds && groupIds.length) {
      descendantToSelected = {};
      groupIds.forEach(function (selectedGid) {
        var descendants = getDescendantIds(selectedGid);
        descendants.forEach(function (descId) {
          if (!descendantToSelected[descId]) descendantToSelected[descId] = [];
          descendantToSelected[descId].push(selectedGid);
        });
      });
    }

    events.forEach(function (evt) {
      if (!evt.activeFrom) return;
      var ruleId = evt.rule ? evt.rule.id : null;
      if (!ruleId || !ruleIdSet.has(ruleId)) return;

      var deviceId = evt.device ? evt.device.id : null;
      var key = getBucketKey(evt.activeFrom, mode);
      if (!bucketsMap[key]) bucketsMap[key] = {};

      if (viewMode === "groups" && descendantToSelected) {
        // Groups mode: match device groups against expanded descendant sets,
        // attribute count to the selected parent group
        var devGroups = deviceId ? (deviceGroupMap[deviceId] || []) : [];
        var attributed = new Set(); // avoid double-counting to same parent
        devGroups.forEach(function (dgid) {
          var parents = descendantToSelected[dgid];
          if (!parents) return;
          parents.forEach(function (selectedGid) {
            if (attributed.has(selectedGid)) return;
            attributed.add(selectedGid);
            var compoundKey = selectedGid + "::" + ruleId;
            if (!bucketsMap[key][compoundKey]) bucketsMap[key][compoundKey] = 0;
            bucketsMap[key][compoundKey]++;
          });
        });
        if (attributed.size === 0) return;
      } else {
        // Company mode: aggregate by rule
        if (!bucketsMap[key][ruleId]) bucketsMap[key][ruleId] = 0;
        bucketsMap[key][ruleId]++;
      }

      if (deviceId) uniqueDevices.add(deviceId);
    });

    var orderedKeys = Object.keys(bucketsMap).sort();

    return {
      orderedKeys: orderedKeys,
      buckets: bucketsMap,
      uniqueDeviceCount: uniqueDevices.size,
      totalEvents: events.length
    };
  }

  // ─── Chart Rendering ───

  var CHART_COLORS = [
    "#4a90d9", "#e74c3c", "#27ae60", "#f39c12", "#9b59b6",
    "#1abc9c", "#e67e22", "#3498db", "#e91e63", "#00bcd4",
    "#8bc34a", "#ff5722", "#607d8b", "#795548", "#cddc39"
  ];

  function resolveRuleColor(ruleId, fallbackIndex) {
    var custom = settings.get().ruleColors[ruleId];
    if (custom) return custom;
    return CHART_COLORS[fallbackIndex % CHART_COLORS.length];
  }

  function buildChartData(orderedKeys, buckets, mode, seriesMeta, mileageByPeriod) {
    // seriesMeta: array of { id, label, stack, colorIndex, ruleId }
    var labels = orderedKeys.slice();
    var datasets = [];

    // One bar dataset per series
    seriesMeta.forEach(function (s) {
      var data = orderedKeys.map(function (key) {
        return (buckets[key] && buckets[key][s.id]) ? buckets[key][s.id] : 0;
      });
      var color = resolveRuleColor(s.ruleId || s.id, s.colorIndex);
      var ds = {
        type: "bar",
        label: s.label,
        _ruleName: s.ruleLabel || s.label,
        data: data,
        backgroundColor: color + "CC",
        borderColor: color,
        borderWidth: 1,
        order: 4
      };
      if (s.stack) ds.stack = s.stack;
      datasets.push(ds);
    });

    // Total per period (sum all series)
    var totalData = orderedKeys.map(function (key) {
      var sum = 0;
      if (buckets[key]) {
        Object.keys(buckets[key]).forEach(function (sid) { sum += buckets[key][sid]; });
      }
      return sum;
    });

    // Total trend line
    datasets.push({
      type: "line",
      label: "Total",
      data: totalData,
      borderColor: "#333",
      backgroundColor: "transparent",
      borderWidth: 2,
      pointRadius: 3,
      pointBackgroundColor: "#333",
      tension: 0.3,
      yAxisID: "y",
      order: 2
    });

    // % Change line (period-over-period on total)
    var pctChangeData = totalData.map(function (val, i) {
      if (i === 0) return null;
      var prev = totalData[i - 1];
      if (prev === 0) return val === 0 ? 0 : null;
      return Math.round(((val - prev) / prev) * 1000) / 10; // one decimal
    });

    datasets.push({
      type: "line",
      label: "% Change",
      data: pctChangeData,
      borderColor: "#e74c3c",
      backgroundColor: "transparent",
      borderWidth: 2,
      borderDash: [5, 3],
      pointRadius: 3,
      pointBackgroundColor: "#e74c3c",
      tension: 0.3,
      yAxisID: "y1",
      order: 1
    });

    // Events per 1k miles trend line
    if (mileageByPeriod) {
      var KM_PER_MILE = 1.60934;
      var per1kData = orderedKeys.map(function (key, i) {
        var events = totalData[i];
        var km = mileageByPeriod[key] || 0;
        if (km === 0) return null;
        var miles = km / KM_PER_MILE;
        return Math.round((events / miles) * 10000) / 10; // per 1k miles, one decimal
      });

      datasets.push({
        type: "line",
        label: "Events / 1k mi",
        data: per1kData,
        borderColor: "#9b59b6",
        backgroundColor: "transparent",
        borderWidth: 2,
        borderDash: [8, 4],
        pointRadius: 3,
        pointBackgroundColor: "#9b59b6",
        tension: 0.3,
        yAxisID: "y2",
        order: 0
      });
    }

    // Collect unique stack (group) names in order for the plugin
    var stackNames = [];
    var stackSeen = {};
    seriesMeta.forEach(function (s) {
      if (s.stack && !stackSeen[s.stack]) {
        stackSeen[s.stack] = true;
        stackNames.push(s.stack);
      }
    });

    return { labels: labels, datasets: datasets, _stackNames: stackNames, _hasY2: !!mileageByPeriod };
  }

  // Chart.js plugin to draw group name labels above each stack cluster
  var groupLabelPlugin = {
    id: "exdGroupLabels",
    afterDatasetsDraw: function (chart) {
      var stackNames = chart.data._stackNames;
      if (!stackNames || stackNames.length <= 1) return;

      var ctx = chart.ctx;
      var xScale = chart.scales.x;
      var yScale = chart.scales.y;
      var meta = {};

      // Collect bar positions per stack per x-index
      chart.data.datasets.forEach(function (ds, dsIdx) {
        if (ds.type !== "bar" || !ds.stack) return;
        var dsMeta = chart.getDatasetMeta(dsIdx);
        if (!dsMeta.visible) return;
        dsMeta.data.forEach(function (bar, i) {
          var key = i + "::" + ds.stack;
          if (!meta[key]) meta[key] = { stack: ds.stack, xIdx: i, minX: Infinity, maxX: -Infinity, topY: Infinity };
          var barX = bar.x;
          var halfW = bar.width / 2;
          meta[key].minX = Math.min(meta[key].minX, barX - halfW);
          meta[key].maxX = Math.max(meta[key].maxX, barX + halfW);
          meta[key].topY = Math.min(meta[key].topY, bar.y);
        });
      });

      ctx.save();
      ctx.font = "bold 10px -apple-system, BlinkMacSystemFont, sans-serif";
      ctx.textAlign = "center";
      ctx.fillStyle = "#666";

      Object.keys(meta).forEach(function (key) {
        var m = meta[key];
        if (m.topY === Infinity) return;
        var centerX = (m.minX + m.maxX) / 2;
        var labelY = m.topY - 6;
        if (labelY < yScale.top + 10) labelY = m.topY + 12;
        ctx.fillText(m.stack, centerX, labelY);
      });

      ctx.restore();
    }
  };

  function renderChart(chartData) {
    // Destroy and recreate to handle axis changes cleanly
    if (chartInstance) { chartInstance.destroy(); chartInstance = null; }
    var ctx = els.chart.getContext("2d");
    chartInstance = new Chart(ctx, {
      data: chartData,
      plugins: [groupLabelPlugin],
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: {
          mode: "index",
          intersect: false
        },
        plugins: {
          legend: {
            position: "bottom",
            labels: {
              boxWidth: 12, padding: 12, font: { size: 12 },
              generateLabels: function (chart) {
                var stackNames = chart.data._stackNames;
                var isGrouped = stackNames && stackNames.length > 1;
                var seen = {};
                var labels = [];

                chart.data.datasets.forEach(function (ds, idx) {
                  if (ds.type === "bar" && isGrouped) {
                    // Dedup by color — show one legend entry per rule
                    var color = ds.borderColor;
                    if (seen[color]) return;
                    seen[color] = true;
                    // Check if ALL datasets with this color are hidden
                    var allHidden = chart.data.datasets.every(function (d, i) {
                      if (d.type !== "bar" || d.borderColor !== color) return true;
                      return !chart.isDatasetVisible(i);
                    });
                    labels.push({
                      text: ds._ruleName || ds.label,
                      fillStyle: allHidden ? "#ddd" : ds.backgroundColor,
                      strokeStyle: allHidden ? "#ccc" : ds.borderColor,
                      lineWidth: 1,
                      hidden: allHidden,
                      _color: color,
                      datasetIndex: idx
                    });
                  } else {
                    // Non-grouped bars, lines (Total, % Change)
                    labels.push({
                      text: ds.label,
                      fillStyle: ds.type === "line" ? "transparent" : ds.backgroundColor,
                      strokeStyle: ds.borderColor,
                      lineWidth: ds.borderWidth || 1,
                      lineDash: ds.borderDash || [],
                      hidden: !chart.isDatasetVisible(idx),
                      datasetIndex: idx
                    });
                  }
                });
                return labels;
              }
            },
            onClick: function (e, legendItem, legend) {
              var chart = legend.chart;
              var stackNames = chart.data._stackNames;
              var isGrouped = stackNames && stackNames.length > 1;

              if (isGrouped && legendItem._color) {
                // Toggle ALL datasets with this rule's color
                var color = legendItem._color;
                var anyVisible = false;
                chart.data.datasets.forEach(function (ds, i) {
                  if (ds.type === "bar" && ds.borderColor === color) {
                    if (chart.isDatasetVisible(i)) anyVisible = true;
                  }
                });
                chart.data.datasets.forEach(function (ds, i) {
                  if (ds.type === "bar" && ds.borderColor === color) {
                    chart.setDatasetVisibility(i, !anyVisible);
                  }
                });
              } else {
                // Default toggle for single dataset
                var idx = legendItem.datasetIndex;
                chart.setDatasetVisibility(idx, !chart.isDatasetVisible(idx));
              }
              chart.update();
            }
          },
          tooltip: {
            callbacks: {
              title: function (items) {
                return items[0].label;
              },
              label: function (ctx) {
                var label = ctx.dataset.label || "";
                var val = ctx.parsed.y;
                if (ctx.dataset.yAxisID === "y1" && val != null) {
                  return label + ": " + val + "%";
                }
                if (ctx.dataset.yAxisID === "y2" && val != null) {
                  return label + ": " + val;
                }
                return label + ": " + (val != null ? val.toLocaleString() : "—");
              }
            }
          }
        },
        scales: {
          x: {
            stacked: true,
            title: { display: true, text: "Period", font: { size: 12 } }
          },
          y: {
            stacked: true,
            beginAtZero: true,
            position: "left",
            title: { display: true, text: "Events", font: { size: 12 } },
            ticks: { precision: 0 }
          },
          y1: {
            position: "right",
            title: { display: true, text: "% Change", font: { size: 12 }, color: "#e74c3c" },
            ticks: {
              color: "#e74c3c",
              callback: function (v) { return v + "%"; }
            },
            grid: { drawOnChartArea: false }
          },
          y2: {
            display: chartData._hasY2 || false,
            position: "right",
            title: { display: true, text: "Events / 1k mi", font: { size: 12 }, color: "#9b59b6" },
            ticks: {
              color: "#9b59b6",
              precision: 1
            },
            grid: { drawOnChartArea: false }
          }
        }
      }
    });
  }

  // ─── Table Rendering ───

  function buildTableRows(orderedKeys, buckets, mode, seriesMeta, viewMode) {
    // seriesMeta: array of { id, label, groupLabel, ruleLabel }
    var rows = [];
    // Track previous period count per series for % change
    var prevCount = {}; // seriesId -> previous count

    orderedKeys.forEach(function (key) {
      seriesMeta.forEach(function (s) {
        var count = (buckets[key] && buckets[key][s.id]) ? buckets[key][s.id] : 0;
        var change = "—";
        if (prevCount.hasOwnProperty(s.id)) {
          var prev = prevCount[s.id];
          if (prev === 0) {
            change = count === 0 ? "0.0%" : "—";
          } else {
            var pct = ((count - prev) / prev) * 100;
            change = (pct >= 0 ? "+" : "") + pct.toFixed(1) + "%";
          }
        }
        prevCount[s.id] = count;

        var row = { period: key, count: count, change: change };
        if (viewMode === "groups") {
          row.group = s.groupLabel || "";
          row.rule = s.ruleLabel || s.label;
        } else {
          row.rule = s.label;
        }
        rows.push(row);
      });
    });

    var headers = viewMode === "groups"
      ? ["period", "group", "rule", "count", "change"]
      : ["period", "rule", "count", "change"];

    return { headers: headers, rows: rows };
  }

  function renderTableHeaders() {
    var thead = els.tableHead;
    thead.innerHTML = "";
    var tr = document.createElement("tr");
    currentHeaders.forEach(function (h) {
      var th = document.createElement("th");
      th.className = "exd-sortable";
      th.dataset.col = h;
      var headerLabels = { period: "Period", group: "Group", rule: "Rule", count: "Count", change: "% Change" };
      th.textContent = headerLabels[h] || h.charAt(0).toUpperCase() + h.slice(1);
      var arrow = document.createElement("span");
      arrow.className = "exd-sort-arrow";
      th.appendChild(arrow);
      if (sortState.col === h) {
        th.classList.add("exd-sort-" + sortState.dir);
      }
      th.addEventListener("click", function () { handleSort(h); });
      tr.appendChild(th);
    });
    thead.appendChild(tr);
  }

  function renderTableBody() {
    var tbody = els.tableBody;
    tbody.innerHTML = "";
    var searchTerm = (els.tableSearch.value || "").toLowerCase();

    var rows = currentRows.slice();

    // Filter
    if (searchTerm) {
      rows = rows.filter(function (r) {
        return currentHeaders.some(function (h) {
          return String(r[h]).toLowerCase().indexOf(searchTerm) >= 0;
        });
      });
    }

    // Sort
    if (sortState.col) {
      var dir = sortState.dir === "asc" ? 1 : -1;
      var col = sortState.col;
      rows.sort(function (a, b) {
        var va = a[col], vb = b[col];
        if (typeof va === "number" && typeof vb === "number") return (va - vb) * dir;
        return String(va).localeCompare(String(vb)) * dir;
      });
    }

    rows.forEach(function (r) {
      var tr = document.createElement("tr");
      currentHeaders.forEach(function (h) {
        var td = document.createElement("td");
        var val = r[h] != null ? r[h] : "";
        td.textContent = val;
        if (h === "count") td.className = "exd-num";
        if (h === "change") {
          td.className = "exd-num";
          if (typeof val === "string" && val.indexOf("+") === 0) td.style.color = "#e74c3c";
          else if (typeof val === "string" && val.indexOf("-") === 0) td.style.color = "#27ae60";
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
  }

  function handleSort(col) {
    if (sortState.col === col) {
      sortState.dir = sortState.dir === "asc" ? "desc" : "asc";
    } else {
      sortState.col = col;
      sortState.dir = col === "count" ? "desc" : "asc";
    }
    renderTableHeaders();
    renderTableBody();
  }

  // ─── KPI Rendering ───

  function renderKpis(data, ruleCount, fromDate, toDate) {
    els.kpiTotal.textContent = data.totalEvents.toLocaleString();
    els.kpiDevices.textContent = data.uniqueDeviceCount.toLocaleString();
    els.kpiPeriod.textContent = formatDate(fromDate) + " to " + formatDate(toDate);
    els.kpiRules.textContent = ruleCount;
  }

  // ─── CSV Export ───

  function exportCsv() {
    if (!currentRows.length) return;
    var lines = [currentHeaders.join(",")];
    currentRows.forEach(function (r) {
      var vals = currentHeaders.map(function (h) {
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
    a.download = "exception_events_" + formatDate(new Date()) + ".csv";
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ─── Generate (main flow) ───

  function generate() {
    // Validate inputs
    var selectedRules = rulePicker.getSelected();
    if (selectedRules.length === 0) {
      setStatus("Please select at least one rule.");
      return;
    }

    var fromStr = els.fromDate.value;
    var toStr = els.toDate.value;
    if (!fromStr || !toStr) {
      setStatus("Please set both From and To dates.");
      return;
    }

    var fromDate = parseLocalDate(fromStr);
    var toDate = parseLocalDate(toStr);
    if (fromDate >= toDate) {
      setStatus("From date must be before To date.");
      return;
    }

    var granularity = document.querySelector("#exd-root .exd-gran-btn.active").dataset.gran;
    var viewMode = document.querySelector("#exd-root .exd-view-btn.active").dataset.view;

    var selectedGroups = [];
    if (viewMode === "groups") {
      selectedGroups = groupPicker.getSelected();
      if (selectedGroups.length === 0) {
        setStatus("Please select at least one group for Groups view.");
        return;
      }
    }

    // Abort any in-flight request
    if (abortController) abortController.abort();
    abortController = new AbortController();

    // Reset UI
    showEmpty(false);
    showWarning(null);
    showLoading(true, "Fetching exception events...");
    setProgress(0);
    setStatus("");
    els.generateBtn.disabled = true;

    // Fetch events and trips in parallel (trips are non-fatal)
    var eventsProm = fetchExceptionEvents(selectedRules, fromDate, toDate, function (pct) {
      setProgress(pct * 0.8); // events = 80% of progress
      els.loadingText.textContent = "Fetching events... " + Math.round(pct * 0.8) + "%";
    });

    els.loadingText.textContent = "Fetching events and trips...";
    var tripsProm = fetchTrips(fromDate, toDate, function (pct) {
      setProgress(80 + pct * 0.2); // trips = last 20%
    }).catch(function (err) {
      console.warn("Trip fetch failed (non-fatal):", err);
      return []; // degrade gracefully — no mileage line
    });

    Promise.all([eventsProm, tripsProm]).then(function (results) {
      var result = results[0];
      var trips = results[1] || [];
      if (isAborted()) return;
      showLoading(false);
      els.generateBtn.disabled = false;

      if (result.hitLimit) {
        showWarning("Warning: Some queries hit the 50,000 result limit. Results may be incomplete. Try a shorter date range.");
      }

      if (result.events.length === 0) {
        showEmpty(true);
        els.kpiTotal.textContent = "0";
        els.kpiDevices.textContent = "0";
        els.kpiPeriod.textContent = formatDate(fromDate) + " to " + formatDate(toDate);
        els.kpiRules.textContent = selectedRules.length;
        if (chartInstance) { chartInstance.destroy(); chartInstance = null; }
        els.tableHead.innerHTML = "";
        els.tableBody.innerHTML = "";
        currentRows = [];
        return;
      }

      // Build label lookups
      var ruleLabelMap = {};
      allRules.forEach(function (r) { ruleLabelMap[r.id] = r.name; });
      var groupLabelMap = {};
      allGroups.forEach(function (g) { groupLabelMap[g.id] = g.name; });

      // Aggregate events
      var agg = aggregateEvents(result.events, selectedRules, selectedGroups, granularity, viewMode);

      // Aggregate trip mileage by period
      var mileageByPeriod = aggregateMileageByPeriod(trips, granularity);

      // Build series metadata
      var seriesMeta = [];
      if (viewMode === "groups") {
        // Compound series: one per group × rule, stacked by group
        selectedGroups.forEach(function (gid) {
          var groupName = groupLabelMap[gid] || gid;
          selectedRules.forEach(function (rid, rIdx) {
            var ruleName = ruleLabelMap[rid] || rid;
            seriesMeta.push({
              id: gid + "::" + rid,
              ruleId: rid,
              label: groupName + " — " + ruleName,
              groupLabel: groupName,
              ruleLabel: ruleName,
              stack: groupName,
              colorIndex: rIdx  // same rule = same color across groups
            });
          });
        });
      } else {
        // Company mode: one series per rule
        selectedRules.forEach(function (rid, idx) {
          seriesMeta.push({
            id: rid,
            ruleId: rid,
            label: ruleLabelMap[rid] || rid,
            ruleLabel: ruleLabelMap[rid] || rid,
            stack: "company",
            colorIndex: idx
          });
        });
      }

      // Chart (pass mileage for events/1k miles line)
      var chartData = buildChartData(agg.orderedKeys, agg.buckets, granularity, seriesMeta, mileageByPeriod);
      renderChart(chartData);

      // Table
      var tableData = buildTableRows(agg.orderedKeys, agg.buckets, granularity, seriesMeta, viewMode);
      currentHeaders = tableData.headers;
      currentRows = tableData.rows;
      sortState = { col: "period", dir: "asc" };
      renderTableHeaders();
      renderTableBody();

      // KPIs
      renderKpis(agg, selectedRules.length, fromDate, toDate);

      var statusParts = [result.events.length.toLocaleString() + " events"];
      if (trips.length > 0) statusParts.push(trips.length.toLocaleString() + " trips");
      if (mileageByPeriod) statusParts.push("mileage \u2713");
      setStatus("Loaded " + statusParts.join(", "));
    }).catch(function (err) {
      if (isAborted()) return;
      showLoading(false);
      els.generateBtn.disabled = false;
      setStatus("Error: " + (err.message || err));
      console.error("generate error:", err);
    });
  }

  // ─── Settings Panel ───

  function openSettingsPanel() {
    populateSettingsPanel();
    els.settingsPanel.classList.add("open");
    els.settingsBackdrop.classList.add("open");
  }

  function closeSettingsPanel() {
    els.settingsPanel.classList.remove("open");
    els.settingsBackdrop.classList.remove("open");
  }

  function populateSettingsPanel() {
    populateDefaultRulesSection();
    populateRuleColorsSection();
    renderViewsList();
  }

  function populateDefaultRulesSection() {
    var listEl = els.settingsRuleList;
    var searchEl = els.settingsRuleSearch;
    var defaults = new Set(settings.get().defaultRules);
    var ruleItems = rulePicker.getItems();

    function renderList() {
      var filt = (searchEl.value || "").toLowerCase();
      listEl.innerHTML = "";
      ruleItems.forEach(function (item) {
        if (filt && item.label.toLowerCase().indexOf(filt) < 0) return;
        var row = document.createElement("div");
        row.className = "exd-settings-rule-row";
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = defaults.has(item.value);
        cb.addEventListener("change", function () {
          if (cb.checked) defaults.add(item.value);
          else defaults.delete(item.value);
        });
        var nameSpan = document.createElement("span");
        nameSpan.className = "exd-settings-rule-name";
        nameSpan.textContent = item.label;
        row.appendChild(cb);
        row.appendChild(nameSpan);
        listEl.appendChild(row);
      });
    }

    renderList();
    searchEl.oninput = renderList;

    // Save button handler
    els.settingsSaveDefaults.onclick = function () {
      settings.setDefaultRules(Array.from(defaults));
      els.settingsSaveDefaults.textContent = "Saved!";
      setTimeout(function () { els.settingsSaveDefaults.textContent = "Save as Defaults"; }, 1200);
    };
  }

  function populateRuleColorsSection() {
    var listEl = els.settingsColorList;
    var searchEl = els.settingsColorSearch;
    var ruleItems = rulePicker.getItems();
    var ruleColors = settings.get().ruleColors;

    function renderList() {
      var filt = (searchEl.value || "").toLowerCase();
      listEl.innerHTML = "";
      ruleItems.forEach(function (item, idx) {
        if (filt && item.label.toLowerCase().indexOf(filt) < 0) return;
        var row = document.createElement("div");
        row.className = "exd-settings-rule-row";
        var colorInput = document.createElement("input");
        colorInput.type = "color";
        colorInput.value = ruleColors[item.value] || CHART_COLORS[idx % CHART_COLORS.length];
        colorInput.addEventListener("change", function () {
          settings.setRuleColor(item.value, colorInput.value);
          ruleColors = settings.get().ruleColors;
          rulePicker.refresh();
        });
        var nameSpan = document.createElement("span");
        nameSpan.className = "exd-settings-rule-name";
        nameSpan.textContent = item.label;
        var resetBtn = document.createElement("button");
        resetBtn.type = "button";
        resetBtn.className = "exd-settings-reset-color";
        resetBtn.textContent = "Reset";
        resetBtn.addEventListener("click", function () {
          settings.clearRuleColor(item.value);
          ruleColors = settings.get().ruleColors;
          colorInput.value = CHART_COLORS[idx % CHART_COLORS.length];
          rulePicker.refresh();
        });
        row.appendChild(colorInput);
        row.appendChild(nameSpan);
        row.appendChild(resetBtn);
        listEl.appendChild(row);
      });
    }

    renderList();
    searchEl.oninput = renderList;
  }

  function captureView(name) {
    var granularity = document.querySelector("#exd-root .exd-gran-btn.active").dataset.gran;
    var viewMode = document.querySelector("#exd-root .exd-view-btn.active").dataset.view;
    return {
      id: "v" + Date.now(),
      name: name,
      rules: rulePicker.getSelected(),
      fromDate: els.fromDate.value,
      toDate: els.toDate.value,
      granularity: granularity,
      viewMode: viewMode,
      groups: groupPicker.getSelected()
    };
  }

  function applyView(view) {
    // Rules
    if (view.rules) rulePicker.setSelected(view.rules);
    // Dates
    if (view.fromDate) els.fromDate.value = view.fromDate;
    if (view.toDate) els.toDate.value = view.toDate;
    // Granularity
    if (view.granularity) {
      document.querySelectorAll("#exd-root .exd-gran-btn").forEach(function (b) {
        b.classList.toggle("active", b.dataset.gran === view.granularity);
      });
    }
    // View mode
    if (view.viewMode) {
      document.querySelectorAll("#exd-root .exd-view-btn").forEach(function (b) {
        b.classList.toggle("active", b.dataset.view === view.viewMode);
      });
      var groupWrap = document.querySelector("#exd-root .exd-group-picker-wrap");
      groupWrap.style.display = view.viewMode === "groups" ? "" : "none";
    }
    // Groups
    if (view.groups) groupPicker.setSelected(view.groups);
    closeSettingsPanel();
  }

  function renderViewsList() {
    var listEl = els.settingsViewsList;
    var views = settings.get().views;
    listEl.innerHTML = "";

    if (views.length === 0) {
      var empty = document.createElement("div");
      empty.style.cssText = "font-size:12px;color:#999;padding:8px 0;";
      empty.textContent = "No saved views yet.";
      listEl.appendChild(empty);
      return;
    }

    views.forEach(function (view) {
      var item = document.createElement("div");
      item.className = "exd-settings-view-item";

      var nameSpan = document.createElement("span");
      nameSpan.className = "exd-settings-view-name";
      nameSpan.textContent = view.name;
      nameSpan.title = view.name;

      var loadBtn = document.createElement("button");
      loadBtn.type = "button";
      loadBtn.className = "exd-settings-view-load";
      loadBtn.textContent = "Load";
      loadBtn.addEventListener("click", function () { applyView(view); });

      var deleteBtn = document.createElement("button");
      deleteBtn.type = "button";
      deleteBtn.className = "exd-settings-view-delete";
      deleteBtn.textContent = "Delete";
      deleteBtn.addEventListener("click", function () {
        settings.deleteView(view.id);
        renderViewsList();
      });

      item.appendChild(nameSpan);
      item.appendChild(loadBtn);
      item.appendChild(deleteBtn);
      listEl.appendChild(item);
    });
  }

  // ─── Event Binding ───

  function bindEvents() {
    // Granularity buttons
    document.querySelectorAll("#exd-root .exd-gran-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        document.querySelectorAll("#exd-root .exd-gran-btn").forEach(function (b) {
          b.classList.remove("active");
        });
        btn.classList.add("active");
      });
    });

    // View toggle
    document.querySelectorAll("#exd-root .exd-view-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        document.querySelectorAll("#exd-root .exd-view-btn").forEach(function (b) {
          b.classList.remove("active");
        });
        btn.classList.add("active");
        var groupWrap = document.querySelector("#exd-root .exd-group-picker-wrap");
        groupWrap.style.display = btn.dataset.view === "groups" ? "" : "none";
      });
    });

    // Generate button
    els.generateBtn.addEventListener("click", generate);

    // Search filter
    els.tableSearch.addEventListener("input", function () {
      renderTableBody();
    });

    // Export CSV
    els.exportCsv.addEventListener("click", exportCsv);

    // Close dropdowns on outside click
    document.addEventListener("click", function () {
      closeAllDropdowns();
    });

    // Settings panel
    els.settingsBtn.addEventListener("click", openSettingsPanel);
    els.settingsClose.addEventListener("click", closeSettingsPanel);
    els.settingsBackdrop.addEventListener("click", closeSettingsPanel);

    // Save view
    els.settingsSaveView.addEventListener("click", function () {
      var name = els.settingsViewName.value.trim();
      if (!name) return;
      var view = captureView(name);
      settings.saveView(view);
      els.settingsViewName.value = "";
      renderViewsList();
    });
  }

  // ─── Set Default Dates ───

  function setDefaultDates() {
    var now = new Date();
    var from = new Date(now);
    from.setDate(from.getDate() - 30);
    els.toDate.value = formatDate(now);
    els.fromDate.value = formatDate(from);
  }

  // ─── Add-in Lifecycle ───

  return {
    initialize: function (freshApi, state, callback) {
      api = freshApi;

      // Init settings (use database name from state, or hostname)
      var dbName = (state && state.database) ? state.database : location.hostname;
      settings.init(dbName);

      // Cache DOM elements
      els.loading = $("exd-loading");
      els.loadingText = els.loading.querySelector(".exd-loading-text");
      els.progressFill = els.loading.querySelector(".exd-progress-fill");
      els.empty = $("exd-empty");
      els.warning = $("exd-warning");
      els.status = $("exd-status");
      els.fromDate = $("exd-from");
      els.toDate = $("exd-to");
      els.generateBtn = $("exd-generate");
      els.chart = $("exd-chart");
      els.tableHead = $("exd-table-head");
      els.tableBody = $("exd-table-body");
      els.tableSearch = $("exd-table-search");
      els.exportCsv = $("exd-export-csv");
      els.kpiTotal = $("exd-kpi-total");
      els.kpiDevices = $("exd-kpi-devices");
      els.kpiPeriod = $("exd-kpi-period");
      els.kpiRules = $("exd-kpi-rules");

      // Settings panel elements
      els.settingsBtn = $("exd-settings-btn");
      els.settingsPanel = $("exd-settings-panel");
      els.settingsBackdrop = $("exd-settings-backdrop");
      els.settingsClose = $("exd-settings-close");
      els.settingsRuleSearch = $("exd-settings-rule-search");
      els.settingsRuleList = $("exd-settings-rule-list");
      els.settingsSaveDefaults = $("exd-settings-save-defaults");
      els.settingsColorSearch = $("exd-settings-color-search");
      els.settingsColorList = $("exd-settings-color-list");
      els.settingsViewName = $("exd-settings-view-name");
      els.settingsSaveView = $("exd-settings-save-view");
      els.settingsViewsList = $("exd-settings-views-list");

      // Init multi-select widgets (with color callback for rule picker)
      rulePicker = initMultiSelect({
        id: "exd-rule-picker",
        placeholder: "Select rules...",
        getColor: function (ruleId) {
          var colors = settings.get().ruleColors;
          return colors[ruleId] || null;
        }
      });
      groupPicker = initMultiSelect({ id: "exd-group-picker", placeholder: "Select groups..." });

      // Bind events
      bindEvents();

      // Set default dates
      setDefaultDates();

      // Load reference data
      loadReferenceData().then(function () {
        callback();
      }).catch(function () {
        callback();
      });
    },

    focus: function (freshApi, state) {
      api = freshApi;
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
