/* =====================================================================
   Malad East Bal Mandal — Sabha Information Management System
   Vanilla JS. No backend. Data comes from /data/*.json (converted
   from the master Excel workbook).
   ===================================================================== */

(function () {
  "use strict";

  /* ------------------------------------------------------------ */
  /* STATE                                                          */
  /* ------------------------------------------------------------ */
  const DATA = {
    sabha: [],
    karyakar: [],
    balak: [],
    mapping: [],
    master: {}
  };

  // lookup helpers, built after data loads
  const IDX = {
    sabhaById: new Map(),
    karyakarById: new Map(),
    karyakarIdsBySabha: new Map(),   // sabhaId -> [karyakarId,...]
    sabhaIdsByKaryakar: new Map(),   // karyakarId -> [sabhaId,...]
    balaksBySabha: new Map()         // sabhaId -> [balak,...]
  };

  const STATE = {
    section: "home",
    balakPage: 1,
    balakPageSize: 24,
    theme: "light"
  };

  const WA_COUNTRY = "91"; // India

  /* ------------------------------------------------------------ */
  /* UTIL                                                           */
  /* ------------------------------------------------------------ */
  function $(sel, ctx) { return (ctx || document).querySelector(sel); }
  function $all(sel, ctx) { return Array.from((ctx || document).querySelectorAll(sel)); }
  function esc(str) {
    return String(str == null ? "" : str).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function digits10(v) { return (v || "").toString().replace(/\D/g, "").slice(-10); }
  function telHref(v) {
    const d = digits10(v);
    return d ? "tel:+91" + d : "";
  }
  function waHref(v) {
    const d = digits10(v);
    return d ? "https://wa.me/" + WA_COUNTRY + d : "";
  }
  function fmtPhone(v) {
    const d = digits10(v);
    if (!d) return "";
    return d.slice(0, 5) + " " + d.slice(5);
  }
  function highlight(text, term) {
    if (!term) return esc(text);
    const t = esc(text);
    const re = new RegExp("(" + term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")", "ig");
    return t.replace(re, "<mark>$1</mark>");
  }
  function debounce(fn, ms) {
    let t;
    return function () {
      clearTimeout(t);
      const args = arguments, ctx = this;
      t = setTimeout(function () { fn.apply(ctx, args); }, ms);
    };
  }
  function showToast(msg, icon) {
    const stack = $("#toastStack");
    const el = document.createElement("div");
    el.className = "toast";
    el.innerHTML = '<span class="material-symbols-outlined">' + (icon || "check_circle") + '</span><span>' + esc(msg) + "</span>";
    stack.appendChild(el);
    setTimeout(function () {
      el.style.opacity = "0";
      el.style.transition = "opacity .3s ease";
      setTimeout(function () { el.remove(); }, 300);
    }, 2200);
  }
  function copyToClipboard(text, label) {
    if (navigator.clipboard) {
      navigator.clipboard.writeText(text).then(function () {
        showToast((label || "Value") + " copied: " + text);
      }).catch(function () { fallbackCopy(text, label); });
    } else {
      fallbackCopy(text, label);
    }
  }
  function fallbackCopy(text, label) {
    const ta = document.createElement("textarea");
    ta.value = text; document.body.appendChild(ta); ta.select();
    try { document.execCommand("copy"); showToast((label || "Value") + " copied: " + text); } catch (e) {}
    ta.remove();
  }

  /* ------------------------------------------------------------ */
  /* DATA LOADING                                                   */
  /* Data is embedded statically in data.js (window.APP_DATA) so    */
  /* the site works by simply double-clicking index.html — no       */
  /* server, no fetch, no CORS issues.                               */
  /* ------------------------------------------------------------ */
  function loadAllData() {
    return new Promise(function (resolve, reject) {
      if (!window.APP_DATA) {
        reject(new Error("APP_DATA not found — make sure data.js is loaded before script.js"));
        return;
      }
      DATA.sabha = window.APP_DATA.sabha || [];
      DATA.karyakar = window.APP_DATA.karyakar || [];
      DATA.balak = window.APP_DATA.balak || [];
      DATA.mapping = window.APP_DATA.mapping || [];
      DATA.master = window.APP_DATA.master || {};
      buildIndexes();
      resolve();
    });
  }

  // IDs come from the source data with inconsistent types (some sabha/karyakar
  // ids are strings, some balak.sabha_id / mapping values are numbers). Every
  // index below is keyed by String(id) so lookups never fail due to a
  // "1" !== 1 type mismatch. Always look these maps up via idKey(...).
  function idKey(v) { return v == null ? "" : String(v); }

  function buildIndexes() {
    IDX.sabhaById.clear();
    IDX.karyakarById.clear();
    IDX.karyakarIdsBySabha.clear();
    IDX.sabhaIdsByKaryakar.clear();
    IDX.balaksBySabha.clear();
    DATA.sabha.forEach(function (s) { IDX.sabhaById.set(idKey(s.id), s); });
    DATA.karyakar.forEach(function (k) { IDX.karyakarById.set(idKey(k.id), k); });
    DATA.mapping.forEach(function (m) {
      const sId = idKey(m.sabha_id), kId = idKey(m.karyakar_id);
      if (!IDX.karyakarIdsBySabha.has(sId)) IDX.karyakarIdsBySabha.set(sId, []);
      IDX.karyakarIdsBySabha.get(sId).push(kId);
      if (!IDX.sabhaIdsByKaryakar.has(kId)) IDX.sabhaIdsByKaryakar.set(kId, []);
      IDX.sabhaIdsByKaryakar.get(kId).push(sId);
    });
    DATA.balak.forEach(function (b) {
      if (b.sabha_id == null || b.sabha_id === "") return;
      const sId = idKey(b.sabha_id);
      if (!IDX.balaksBySabha.has(sId)) IDX.balaksBySabha.set(sId, []);
      IDX.balaksBySabha.get(sId).push(b);
    });
  }

  /* ------------------------------------------------------------ */
  /* LOCAL EDITS (Add / Edit Balak)                                  */
  /* This is a static, backend-free site — there is no server or      */
  /* database to write to. New / edited Balak records are therefore   */
  /* saved in this browser's localStorage only, layered on top of     */
  /* the original data embedded in data.js. They will not appear on   */
  /* other devices or browsers unless exported/shared manually.        */
  /* ------------------------------------------------------------ */
  const LOCAL_EDITS_KEY = "balmandal_balak_local_edits_v1";

  function loadLocalEdits() {
    try {
      const raw = localStorage.getItem(LOCAL_EDITS_KEY);
      if (!raw) return { edited: {}, added: [] };
      const parsed = JSON.parse(raw);
      return { edited: parsed.edited || {}, added: parsed.added || [] };
    } catch (e) { return { edited: {}, added: [] }; }
  }
  function saveLocalEdits(obj) {
    try { localStorage.setItem(LOCAL_EDITS_KEY, JSON.stringify(obj)); } catch (e) { showToast("Could not save locally — storage unavailable", "error"); }
  }
  function applyLocalEditsToBalak() {
    const local = loadLocalEdits();
    const base = (window.APP_DATA.balak || []).map(function (b) {
      return local.edited[b.id] ? Object.assign({}, b, local.edited[b.id]) : b;
    });
    DATA.balak = base.concat(local.added);
  }
  function nextLocalBalakId() {
    const allIds = DATA.balak.map(function (b) { return b.id; });
    return (allIds.length ? Math.max.apply(null, allIds) : 0) + 1;
  }
  function upsertLocalBalak(record, isNew) {
    const local = loadLocalEdits();
    if (isNew) {
      local.added.push(record);
    } else {
      const wasAdded = local.added.some(function (b) { return b.id === record.id; });
      if (wasAdded) {
        local.added = local.added.map(function (b) { return b.id === record.id ? record : b; });
      } else {
        local.edited[record.id] = record;
      }
    }
    saveLocalEdits(local);
    applyLocalEditsToBalak();
    buildIndexes();
  }
  function deleteLocalBalak(id) {
    const local = loadLocalEdits();
    local.added = local.added.filter(function (b) { return b.id !== id; });
    // if it was an originally-embedded record, mark it removed via a tombstone flag
    local.edited[id] = Object.assign({}, local.edited[id], { _deleted: true });
    saveLocalEdits(local);
    const local2 = loadLocalEdits();
    const base = (window.APP_DATA.balak || []).filter(function (b) { return !(local2.edited[b.id] && local2.edited[b.id]._deleted); });
    DATA.balak = base.map(function (b) {
      return local2.edited[b.id] ? Object.assign({}, b, local2.edited[b.id]) : b;
    }).concat(local2.added);
    buildIndexes();
  }

  /* ------------------------------------------------------------ */
  /* CONTACT CHIP / BUTTON BUILDERS  (phone + WhatsApp "prompt open") */
  /* ------------------------------------------------------------ */
  function contactChip(number, label) {
    const d = digits10(number);
    if (!d) return "";
    return (
      '<span class="contact-chip">' +
        '<a class="icon-btn call" href="' + telHref(number) + '" title="Call ' + esc(label || "") + '"><span class="material-symbols-outlined" style="font-size:13px;">call</span></a>' +
        '<a class="icon-btn wa" href="' + waHref(number) + '" target="_blank" rel="noopener" title="WhatsApp ' + esc(label || "") + '"><span class="material-symbols-outlined" style="font-size:13px;">chat</span></a>' +
        '<a href="#" class="num-text" data-action="copy" data-value="' + esc(fmtPhone(number)) + '">' + esc(fmtPhone(number)) + "</a>" +
      "</span>"
    );
  }

  function ecButtonsCallWa(number, label) {
    const d = digits10(number);
    if (!d) return "";
    return (
      '<a class="ec-btn call" href="' + telHref(number) + '"><span class="material-symbols-outlined">call</span>' + esc(label || "Call") + "</a>" +
      '<a class="ec-btn wa" href="' + waHref(number) + '" target="_blank" rel="noopener"><span class="material-symbols-outlined">chat</span>WhatsApp</a>'
    );
  }

  /* ------------------------------------------------------------ */
  /* ROUTING / SIDEBAR                                              */
  /* ------------------------------------------------------------ */
  const SECTION_META = {
    home: { icon: "home", title: "Home" },
    summary: { icon: "bar_chart", title: "Summary" },
    sabha: { icon: "temple_hindu", title: "Sabha Directory" },
    karyakar: { icon: "groups", title: "Karyakar Directory" },
    balak: { icon: "child_care", title: "Balak Directory" },
    daywise: { icon: "calendar_month", title: "Day Wise" },
    contact: { icon: "call", title: "Contact" }
  };

  function goto(section) {
    STATE.section = section;
    $all(".page-section").forEach(function (el) { el.classList.toggle("active", el.id === "sec-" + section); });
    $all(".sidebar-link").forEach(function (el) { el.classList.toggle("active", el.getAttribute("data-goto") === section); });
    closeSidebarMobile();
    window.scrollTo({ top: 0, behavior: "smooth" });
    if (section === "summary") renderCharts();
  }

  function closeSidebarMobile() {
    document.body.classList.remove("sidebar-open");
  }

  /* ------------------------------------------------------------ */
  /* HOME                                                           */
  /* ------------------------------------------------------------ */
 function renderHome() {
    $("#homeTotalSabha").textContent = DATA.sabha.filter(function (s) { return s.is_active; }).length;
    $("#homeTotalKaryakar").textContent = DATA.karyakar.filter(function (s) { return s.is_active; }).length;
    $("#homeTotalBalak").textContent = DATA.balak.length;
    $("#homeActiveSabha").textContent = DATA.sabha.filter(function (s) { return s.is_active; }).length;
  }

  /* ------------------------------------------------------------ */
  /* SUMMARY                                                        */
  /* ------------------------------------------------------------ */
  let charts = {};
  function renderSummary() {
    const totalSabha = DATA.sabha.filter(function (s) { return s.is_active; }).length;
    const activeSabha = DATA.sabha.filter(function (s) { return s.is_active; }).length;
    const inactiveSabha = totalSabha - activeSabha;
    const totalKaryakar = DATA.karyakar.filter(function (s) { return s.is_active; }).length;
    const activeKaryakar = DATA.karyakar.filter(function (k) { return k.is_active; }).length;
    const totalBalak = DATA.balak.length;
    const registeredBalak = DATA.balak.filter(function (b) { return b.registered; }).length;

    $("#sumTotalSabha").textContent = totalSabha;
    $("#sumActiveSabha").textContent = activeSabha;
    $("#sumInactiveSabha").textContent = inactiveSabha;
    $("#sumTotalKaryakar").textContent = totalKaryakar;
    $("#sumActiveKaryakar").textContent = activeKaryakar;
    $("#sumTotalBalak").textContent = totalBalak;
    $("#sumRegisteredBalak").textContent = registeredBalak;
  }

  function countBy(list, keyFn) {
    const map = {};
    list.forEach(function (item) {
      const k = keyFn(item) || "Unspecified";
      map[k] = (map[k] || 0) + 1;
    });
    return map;
  }

  const CHART_COLORS = ["#6a2c91", "#e0762f", "#2f8f6e", "#9b3fb0", "#c0392b", "#2f7bc0", "#c9a227"];

  function renderCharts() {
    if (typeof Chart === "undefined") return;
    const dayOrder = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

    const typeCounts = countBy(DATA.sabha, function (s) { return s.sabha_type; });
    const kshetraCounts = countBy(DATA.sabha, function (s) { return s.kshetra; });
    const dayCounts = countBy(DATA.sabha, function (s) { return s.day; });
    const dayLabels = dayOrder.filter(function (d) { return dayCounts[d]; });
    const dayValues = dayLabels.map(function (d) { return dayCounts[d]; });

    buildChart("chartSabhaType", "pie", Object.keys(typeCounts), Object.values(typeCounts), "Sabha Type");
    buildChart("chartKshetra", "bar", Object.keys(kshetraCounts), Object.values(kshetraCounts), "Kshetra Wise Sabha");
    buildChart("chartDayWise", "bar", dayLabels, dayValues, "Day Wise Sabha Count");
  }

  function buildChart(canvasId, type, labels, data, label) {
    const el = document.getElementById(canvasId);
    if (!el) return;
    if (charts[canvasId]) charts[canvasId].destroy();
    charts[canvasId] = new Chart(el.getContext("2d"), {
      type: type,
      data: {
        labels: labels,
        datasets: [{
          label: label,
          data: data,
          backgroundColor: CHART_COLORS,
          borderRadius: type === "bar" ? 6 : 0,
          borderWidth: type === "pie" ? 2 : 0,
          borderColor: "#fff"
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: type === "pie" ? "bottom" : "none", labels: { boxWidth: 12, font: { size: 11 } } } },
        scales: type === "bar" ? { y: { beginAtZero: true, ticks: { precision: 0 } } } : {}
      }
    });
  }

  /* ------------------------------------------------------------ */
  /* SABHA DIRECTORY                                                */
  /* ------------------------------------------------------------ */
  function populateSabhaFilters() {
    fillSelect("#sabhaFilterDay", DATA.master.Day || uniq(DATA.sabha, "day"));
    fillSelect("#sabhaFilterArea", uniq(DATA.sabha, "area"));
    fillSelect("#sabhaFilterKshetra", DATA.master.Kshetra || uniq(DATA.sabha, "kshetra"));
    fillSelect("#sabhaFilterType", DATA.master.SabhaType || uniq(DATA.sabha, "sabha_type"));
  }
  function uniq(list, key) {
    return Array.from(new Set(list.map(function (i) { return i[key]; }).filter(Boolean))).sort();
  }
  function fillSelect(sel, values) {
    const el = $(sel);
    if (!el) return;
    const current = el.value;
    el.innerHTML = '<option value="">All</option>' + values.map(function (v) {
      return '<option value="' + esc(v) + '">' + esc(v) + "</option>";
    }).join("");
    el.value = current || "";
  }

  function getSabhaFiltered() {
    const term = ($("#sabhaSearch").value || "").trim().toLowerCase();
    const day = $("#sabhaFilterDay").value;
    const area = $("#sabhaFilterArea").value;
    const kshetra = $("#sabhaFilterKshetra").value;
    const type = $("#sabhaFilterType").value;
    const sortBy = $("#sabhaSort").value;

    let list = DATA.sabha.filter(function (s) {
      if (!s.is_active) return false;
      if (day && s.day !== day) return false;
      if (area && s.area !== area) return false;
      if (kshetra && s.kshetra !== kshetra) return false;
      if (type && s.sabha_type !== type) return false;
      if (term) {
        const hay = [s.name, s.area, s.day, s.sabha_type, s.kshetra, s.reg_no, s.address].join(" ").toLowerCase();
        if (hay.indexOf(term) === -1) return false;
      }
      return true;
    });

    if (sortBy === "name") list.sort(function (a, b) { return a.name.localeCompare(b.name); });
    else if (sortBy === "day") {
      const order = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
      list.sort(function (a, b) { return order.indexOf(a.day) - order.indexOf(b.day); });
    } else if (sortBy === "area") list.sort(function (a, b) { return (a.area || "").localeCompare(b.area || ""); });

    return { list: list, term: term };
  }

  function renderSabhaList() {
    const grid = $("#sabhaGrid");
    const res = getSabhaFiltered();
    $("#sabhaResultCount").textContent = res.list.length + " sabha" + (res.list.length === 1 ? "" : "s");

    if (!res.list.length) {
      grid.innerHTML = emptyState("No Sabha found", "Try adjusting your search or filters.");
      return;
    }

    grid.innerHTML = res.list.map(function (s) {
      const karyakarCount = (IDX.karyakarIdsBySabha.get(idKey(s.id)) || []).length;
      const balakCount = (IDX.balaksBySabha.get(idKey(s.id)) || []).length;
      return (
        '<div class="entity-card" id="sabha-card-' + s.id + '">' +
          '<div class="ec-head">' +
            '<div><div class="ec-name">' + highlight(s.name, res.term) + '</div>' +
            '<div class="ec-sub">' + esc(s.reg_no) + "</div></div>" +
            '<div class="ec-badges">' +
              '<span class="ec-badge purple">' + esc(s.reg_no) + "</span>" +
              '<span class="ec-badge orange">' + esc(s.sabha_type) + "</span>" +
            "</div>" +
          "</div>" +
          '<div class="ec-body">' +
            row("event", "Day / Time", esc(s.day) + " &middot; " + esc(s.time)) +
            row("place", "Area / Kshetra", esc(s.area) + " (" + esc(s.kshetra) + ")") +
            row("home_pin", "Address", esc(s.address)) +
            row("person", "Owner", esc(s.owner_name)) +
            row("groups", "Karyakar / Balak", karyakarCount + " Karyakar &middot; " + balakCount + " Balak") +
          "</div>" +
          '<div class="ec-actions">' +
            (s.google_link ? '<a class="ec-btn map" href="' + esc(s.google_link) + '" target="_blank" rel="noopener"><span class="material-symbols-outlined">location_on</span>Map</a>' : "") +
            ecButtonsCallWa(s.contact1, "Call") +
            '<button class="ec-btn outline" data-action="view-sabha" data-id="' + s.id + '"><span class="material-symbols-outlined">visibility</span>Details</button>' +
          "</div>" +
        "</div>"
      );
    }).join("");
  }

  function row(icon, label, valueHtml) {
    if (!valueHtml || valueHtml === "undefined") return "";
    return (
      '<div class="ec-row"><span class="material-symbols-outlined">' + icon + '</span>' +
      '<div><span class="ec-label">' + esc(label) + '</span><span class="ec-value">' + valueHtml + "</span></div></div>"
    );
  }

  function emptyState(title, sub) {
    return (
      '<div class="empty-state" style="grid-column:1/-1;">' +
        '<span class="material-symbols-outlined">search_off</span>' +
        "<h4>" + esc(title) + "</h4><p>" + esc(sub) + "</p>" +
      "</div>"
    );
  }

  /* ------------------------------------------------------------ */
  /* SABHA DETAIL MODAL                                             */
  /* ------------------------------------------------------------ */
  function openSabhaModal(id) {
    const s = IDX.sabhaById.get(idKey(id));
    if (!s) return;
    const karyakarIds = IDX.karyakarIdsBySabha.get(idKey(s.id)) || [];
    const karyakarList = karyakarIds.map(function (kid) { return IDX.karyakarById.get(idKey(kid)); }).filter(Boolean);
    const balakList = IDX.balaksBySabha.get(idKey(s.id)) || [];

    $("#dmTitle").textContent = s.name;
    $("#dmSub").textContent = s.reg_no + " \u00b7 " + s.area;

    let html = "";
    html += '<h4>Sabha Information</h4><div class="dm-grid">';
    html += dmItem("Mandal Type", s.mandal_type);
    html += dmItem("Kshetra", s.kshetra);
    html += dmItem("Area", s.area);
    html += dmItem("Sabha Type", s.sabha_type);
    html += dmItem("Day", s.day);
    html += dmItem("Time", s.from_to_time || s.time);
    html += dmItem("Status", s.is_active ? "Active" : "Inactive");
    html += dmItem("Santo Visit", s.when_santo_visited);
    html += dmItemFull("Address", s.address);
    html += dmItemFull("Owner", s.owner_name);
    html += "</div>";

    html += '<h4>Contact</h4><div class="dm-grid">';
    if (digits10(s.contact1)) html += dmItemFull("Contact 1", contactChip(s.contact1, s.owner_name));
    if (digits10(s.contact2)) html += dmItemFull("Contact 2", contactChip(s.contact2, s.owner_name));
    if (s.google_link) html += dmItemFull("Google Maps", '<a href="' + esc(s.google_link) + '" target="_blank" rel="noopener">Open location &rarr;</a>');
    html += "</div>";

    html += '<h4>Assigned Karyakar (' + karyakarList.length + ')</h4>';
    html += karyakarList.length
      ? karyakarList.map(function (k) {
          return '<div class="mini-person"><div><div class="mp-name">' + esc(k.name) + '</div><div class="mp-sub">' + esc(k.dob || "") + '</div></div>' + contactChip(k.contact1, k.name) + "</div>";
        }).join("")
      : '<p class="mp-sub">No karyakar assigned yet.</p>';

    html += '<h4>Balak List (' + balakList.length + ')</h4>';
    html += balakList.length
      ? balakList.map(function (b) {
          return '<div class="mini-person"><div><div class="mp-name">' + esc(b.full_name || (b.name + " " + b.surname)) + '</div><div class="mp-sub">Std ' + esc(b.current_std) + "</div></div>" + contactChip(b.contact_home, b.full_name) + "</div>";
        }).join("")
      : '<p class="mp-sub">No balak registered yet.</p>';

    $("#dmBody").innerHTML = html;

    $("#dmActions").innerHTML =
      (s.google_link ? '<a class="ec-btn map" href="' + esc(s.google_link) + '" target="_blank" rel="noopener"><span class="material-symbols-outlined">location_on</span>Navigate</a>' : "") +
      '<button class="ec-btn outline" data-action="print-sabha"><span class="material-symbols-outlined">print</span>Print</button>' +
      '<button class="ec-btn outline" data-action="copy-map" data-value="' + esc(s.google_link || "") + '"><span class="material-symbols-outlined">content_copy</span>Copy Map Link</button>';

    showModal();
    document.body.dataset.currentSabha = s.id;
  }
  function dmItem(label, val) {
    if (val == null || val === "" || val === false) return "";
    return '<div class="dm-item"><span class="lbl">' + esc(label) + '</span><span class="val">' + esc(val) + "</span></div>";
  }
  function dmItemFull(label, valHtml) {
    if (!valHtml) return "";
    return '<div class="dm-item full"><span class="lbl">' + esc(label) + '</span><span class="val">' + valHtml + "</span></div>";
  }
  function showModal() { $("#detailModal").classList.add("show"); document.body.style.overflow = "hidden"; }
  function hideModal() { $("#detailModal").classList.remove("show"); document.body.style.overflow = ""; }

  /* ------------------------------------------------------------ */
  /* KARYAKAR DIRECTORY                                             */
  /* ------------------------------------------------------------ */
  function getKaryakarFiltered() {
    const term = ($("#karyakarSearch").value || "").trim().toLowerCase();
    const dob = $("#karyakarFilterdob").value;
       
    let list = DATA.karyakar.filter(function (k) {
      if (!k.is_active) return false;
      // if (dob && k.dob !== dob) return false;
      
      // DOB Month Filter
   
       if (dob) {
         const month = k.dob ? k.dob.trim().split(/\s+/)[1] : "";
         if (month !== dob) return false;
       }

     if (term) {
        const hay = [k.name, k.contact1, k.area, k.email, k.dob, k.company].join(" ").toLowerCase();
        if (hay.indexOf(term) === -1) return false;
      }
 


      return true;
    });
    list.sort(function (a, b) { return a.name.localeCompare(b.name); });
    return { list: list, term: term };
  }

 // function populateKaryakarFilters() {
 //   fillSelect("#karyakarFilterdob", uniq(DATA.karyakar, "dob"));
 // }

  function populateKaryakarFilters() {
  const monthOrder = [
    "Jan","Feb","Mar","Apr","May","Jun",
    "Jul","Aug","Sep","Oct","Nov","Dec"
  ];
   const months = [...new Set(
    DATA.karyakar
      .map(k => k.dob?.trim().split(/\s+/)[1])
      .filter(Boolean)
  )].sort((a, b) => monthOrder.indexOf(a) - monthOrder.indexOf(b));

  fillSelect("#karyakarFilterdob", months);
}

  function renderKaryakarList() {
    const grid = $("#karyakarGrid");
    const res = getKaryakarFiltered();
    $("#karyakarResultCount").textContent = res.list.length + " karyakar" + (res.list.length === 1 ? "" : "s");

    if (!res.list.length) {
      grid.innerHTML = emptyState("No Karyakar found", "Try adjusting your search or filters.");
      return;
    }

    grid.innerHTML = res.list.map(function (k) {
      const sabhaIds = IDX.sabhaIdsByKaryakar.get(idKey(k.id)) || [];
      const sabhaNames = sabhaIds.map(function (sid) { const s = IDX.sabhaById.get(idKey(sid)); return s ? s.name : null; }).filter(Boolean);
      return (
        '<div class="entity-card">' +
          '<div class="ec-head">' +
            '<div><div class="ec-name">' + highlight(k.name, res.term) + '</div>' +
            '<div class="ec-sub">Reg: ' + esc(k.reg_no || "-") + "</div></div>" +
            '<div class="ec-badges"><span class="ec-badge purple">' + esc(k.reg_no || "-") + "</span></div>" +
          "</div>" +
          '<div class="ec-body">' +
            row("place", "Area", esc(k.area)) +
            (k.education ? row("school", "Education", esc(k.education)) : "") +
            (k.company ? row("apartment", "Company", esc(k.company)) : "") +
            (k.dob ? row("calendar_month", "Dob", esc(k.dob)) : "") +
            (k.work_profile ? row("work", "Work_profile", esc(k.work_profile)) : "") + 
            (k.address ? row("home_pin", "Address", esc(k.address)) : "") +
            (sabhaNames.length ? row("temple_hindu", "Assigned Sabha", sabhaNames.map(function (n, i) {
              return '<a href="#" data-action="view-sabha" data-id="' + sabhaIds[i] + '">' + esc(n) + "</a>";
            }).join(", ")) : "") +
          "</div>" +
          '<div class="ec-actions">' +
            ecButtonsCallWa(k.contact1, "Call") +
            (k.email ? '<a class="ec-btn outline" href="mailto:' + esc(k.email) + '"><span class="material-symbols-outlined">mail</span>Email</a>' : "") +
          "</div>" +
        "</div>"
      );
    }).join("");
  }

  /* ------------------------------------------------------------ */
  /* BALAK DIRECTORY                                                */
  /* ------------------------------------------------------------ */
  function getBalakFiltered() {
    const term = ($("#balakSearch").value || "").trim().toLowerCase();
    const sabhaId = $("#balakFilterSabha").value;
    const std = $("#balakFilterStd").value;
    const registered = $("#balakFilterRegistered").value;
    const dob = $("#balakFilterdob").value;
  
      
    let list = DATA.balak.filter(function (b) {
      if (sabhaId && String(b.sabha_id) !== sabhaId) return false;
      if (std && b.current_std !== std) return false;
      if (registered === "yes" && !b.registered) return false;
      if (registered === "no" && b.registered) return false;

       // DOB Month Filter
      if (dob) {
         const month = b.dob ? b.dob.trim().split(/\s+/)[1] : "";
         if (month !== dob) return false;
       }

      if (term) {
        const hay = [b.full_name, b.name, b.father_name, b.dob, b.school].join(" ").toLowerCase();
        if (hay.indexOf(term) === -1) return false;
      }
      return true;
    });
    list.sort(function (a, b) { return (a.full_name || a.name || "").localeCompare(b.full_name || b.name || ""); });
    return { list: list, term: term };
  }

  function populateBalakFilters() {
    const sel = $("#balakFilterSabha");
    const current = sel.value;
    const monthOrder = [
      "Jan","Feb","Mar","Apr","May","Jun",
      "Jul","Aug","Sep","Oct","Nov","Dec"
    ];
    const months = [...new Set(
      DATA.balak
        .map(b=> b.dob?.trim().split(/\s+/)[1])
        .filter(Boolean)
    )].sort((a, b) => monthOrder.indexOf(a) - monthOrder.indexOf(b));


    sel.innerHTML = '<option value="">All Sabha</option>' + DATA.sabha.slice().sort(function (a, b) { return a.name.localeCompare(b.name); }).map(function (s) {
      return '<option value="' + s.id + '">' + esc(s.name) + "</option>";
    }).join("");
    sel.value = current || "";
    fillSelect("#balakFilterStd", uniq(DATA.balak, "current_std"));
         
    fillSelect("#balakFilterdob", months);
      
  }

  

  function renderBalakList() {
    const grid = $("#balakGrid");
    const res = getBalakFiltered();
    $("#balakResultCount").textContent = res.list.length + " balak";

    if (!res.list.length) {
      grid.innerHTML = emptyState("No Balak found", "Try adjusting your search or filters.");
      $("#balakPaginationBar").innerHTML = "";
      return;
    }

    const totalPages = Math.max(1, Math.ceil(res.list.length / STATE.balakPageSize));
    if (STATE.balakPage > totalPages) STATE.balakPage = totalPages;
    const startIdx = (STATE.balakPage - 1) * STATE.balakPageSize;
    const pageItems = res.list.slice(startIdx, startIdx + STATE.balakPageSize);

    grid.innerHTML = pageItems.map(function (b) {
      const sabha = b.sabha_id != null ? IDX.sabhaById.get(idKey(b.sabha_id)) : null;
      return (
        '<div class="entity-card">' +
          '<div class="ec-head">' +
            '<div><div class="ec-name">' + highlight(b.full_name || b.name, res.term) + '</div>' +
            '<div class="ec-sub">S/o D/o ' + esc(b.father_name) + "</div></div>" +
            '<div class="ec-badges"><span class="ec-badge purple">Std ' + esc(b.current_std || "-") + "</span>" +
            (b.registered ? '<span class="ec-badge green">Registered</span>' : "") + "</div>" +
          "</div>" +
          '<div class="ec-body">' +
            (b.dob ? row("calendar_month", "dob", esc(b.dob)) : "") +
            (b.puja_kare_che ? row("book", "Puja Kare Che?", esc(b.puja_kare_che)) : "") +
            (b.svp ? row("book", "SVP", esc(b.svp)) : "") +
            (b.satsangi_category ? row("book", "Satsangi Category", esc(b.satsangi_category)) : "") +
            (b.bal_prakash ? row("book", "Bal Prakash", esc(b.bal_prakash)) : "") +
                      
            (b.school ? row("school", "School", esc(b.school)) : "") +
            (sabha ? row("temple_hindu", "Sabha", '<a href="#" data-action="view-sabha" data-id="' + sabha.id + '">' + esc(sabha.name) + "</a>") : "") +
            (b.address ? row("home_pin", "Address", esc(b.address)) : "") +
            (b.remarks ? row("sticky_note_2", "Remarks", esc(b.remarks)) : "") +
          "</div>" +
          '<div class="ec-actions">' +
            (digits10(b.contact_home) ? ecButtonsCallWa(b.contact_home, "Home") : "") +
            (digits10(b.contact_father) ? ecButtonsCallWa(b.contact_father, "Father") : "") +
            (sabha ? '<button class="ec-btn outline" data-action="view-sabha" data-id="' + sabha.id + '"><span class="material-symbols-outlined">visibility</span>Sabha</button>' : "") +
          "</div>" +
        "</div>"
      );
    }).join("");

    $("#balakPaginationBar").innerHTML =
      '<button data-action="balak-prev" ' + (STATE.balakPage <= 1 ? "disabled" : "") + '><span class="material-symbols-outlined" style="vertical-align:middle;font-size:16px;">chevron_left</span>Prev</button>' +
      '<span>Page ' + STATE.balakPage + " of " + totalPages + "</span>" +
      '<button data-action="balak-next" ' + (STATE.balakPage >= totalPages ? "disabled" : "") + '>Next<span class="material-symbols-outlined" style="vertical-align:middle;font-size:16px;">chevron_right</span></button>';
  }

  /* ------------------------------------------------------------ */
  /* DAY WISE                                                       */
  /* ------------------------------------------------------------ */
  function renderDayWise() {
    const order = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    const wrap = $("#dayWiseWrap");
    wrap.innerHTML = order.map(function (day, i) {
      const items = DATA.sabha.filter(function (s) { return s.day === day && s.is_active; });
      if (!items.length) return "";
      return (
        '<div class="day-accordion-item' + (i === 0 ? " open" : "") + '">' +
          '<div class="day-accordion-head" data-action="toggle-day">' +
            '<div class="dh-left"><span class="material-symbols-outlined">event</span><h3>' + day + '</h3></div>' +
            '<div class="dh-left"><span class="dh-count">' + items.length + " sabha</span><span class=\"material-symbols-outlined chev\">expand_more</span></div>" +
          "</div>" +
          '<div class="day-accordion-body"><div class="cards-grid">' +
            items.map(function (s) {
              const karyakarCount = (IDX.karyakarIdsBySabha.get(idKey(s.id)) || []).length;
              const balakCount = (IDX.balaksBySabha.get(idKey(s.id)) || []).length;
              return (
                '<div class="entity-card">' +
                  '<div class="ec-head"><div><div class="ec-name">' + esc(s.name) + '</div><div class="ec-sub">' + esc(s.area) + "</div></div>" +
                  '<div class="ec-badges"><span class="ec-badge purple">' + esc(s.reg_no) + '</span><span class="ec-badge orange">' + esc(s.sabha_type) + "</span></div></div>" +
                  '<div class="ec-body">' +
                    row("schedule", "Time", esc(s.from_to_time || s.time)) +
                    row("home_pin", "Address", esc(s.address)) +
                    row("groups", "Karyakar / Balak", karyakarCount + " Karyakar &middot; " + balakCount + " Balak") +
                  "</div>" +
                  '<div class="ec-actions"><button class="ec-btn outline" data-action="view-sabha" data-id="' + s.id + '"><span class="material-symbols-outlined">visibility</span>Details</button></div>' +
                "</div>"
              );
            }).join("") +
          "</div></div>" +
        "</div>"
      );
    }).join("");
  }

  /* ------------------------------------------------------------ */
  /* GLOBAL SEARCH                                                  */
  /* ------------------------------------------------------------ */
  function globalSearch(term) {
    const box = $("#globalSearchResults");
    if (!term || term.length < 2) { box.classList.remove("show"); box.innerHTML = ""; return; }
    const t = term.toLowerCase();

    const sabhaHits = DATA.sabha.filter(function (s) { return s.name.toLowerCase().indexOf(t) !== -1 || (s.area || "").toLowerCase().indexOf(t) !== -1; }).slice(0, 5);
    const karyakarHits = DATA.karyakar.filter(function (k) { return k.name.toLowerCase().indexOf(t) !== -1; }).slice(0, 5);
    const balakHits = DATA.balak.filter(function (b) { return (b.full_name || "").toLowerCase().indexOf(t) !== -1 || (b.father_name || "").toLowerCase().indexOf(t) !== -1; }).slice(0, 5);

    let html = "";
    if (sabhaHits.length) {
      html += '<div class="gsr-group-title">Sabha</div>' + sabhaHits.map(function (s) {
        return '<div class="gsr-item" data-action="goto-sabha" data-id="' + s.id + '"><span class="material-symbols-outlined">temple_hindu</span><div>' + highlight(s.name, t) + "<small>" + esc(s.area) + " &middot; " + esc(s.day) + "</small></div></div>";
      }).join("");
    }
    if (karyakarHits.length) {
      html += '<div class="gsr-group-title">Karyakar</div>' + karyakarHits.map(function (k) {
        return '<div class="gsr-item" data-action="goto-karyakar" data-id="' + k.id + '"><span class="material-symbols-outlined">person</span><div>' + highlight(k.name, t) + "<small>" + esc(k.area) + "</small></div></div>";
      }).join("");
    }
    if (balakHits.length) {
      html += '<div class="gsr-group-title">Balak</div>' + balakHits.map(function (b) {
        return '<div class="gsr-item" data-action="goto-balak" data-id="' + b.id + '"><span class="material-symbols-outlined">child_care</span><div>' + highlight(b.full_name || b.name, t) + "<small>Std " + esc(b.current_std) + "</small></div></div>";
      }).join("");
    }
    if (!html) html = '<div class="gsr-empty">No results for "' + esc(term) + '"</div>';
    box.innerHTML = html;
    box.classList.add("show");
  }

  /* ------------------------------------------------------------ */
  /* EVENT WIRING                                                   */
  /* ------------------------------------------------------------ */
  function wireEvents() {
    // sidebar nav
    $all("[data-goto]").forEach(function (el) {
      el.addEventListener("click", function (e) { e.preventDefault(); goto(el.getAttribute("data-goto")); });
    });

    $("#burgerBtn").addEventListener("click", function () {
      if (window.innerWidth < 992) document.body.classList.toggle("sidebar-open");
      else document.body.classList.toggle("sidebar-collapsed");
    });
    $("#sidebarBackdrop").addEventListener("click", closeSidebarMobile);

    // theme toggle
    $("#themeToggleBtn").addEventListener("click", function () {
      STATE.theme = STATE.theme === "light" ? "dark" : "light";
      document.documentElement.setAttribute("data-theme", STATE.theme);
      $("#themeToggleBtn .material-symbols-outlined").textContent = STATE.theme === "light" ? "dark_mode" : "light_mode";
    });

    // global search
    const gsInput = $("#globalSearchInput");
    gsInput.addEventListener("input", debounce(function () { globalSearch(gsInput.value.trim()); }, 150));
    gsInput.addEventListener("focus", function () { if (gsInput.value.trim().length >= 2) globalSearch(gsInput.value.trim()); });
    document.addEventListener("click", function (e) {
      if (!e.target.closest(".global-search-wrap")) $("#globalSearchResults").classList.remove("show");
    });

    // hero quick search -> jumps to sabha search
    $("#heroSearchInput").addEventListener("input", debounce(function () {
      const v = this.value.trim();
      if (v.length >= 2) {
        goto("sabha");
        $("#sabhaSearch").value = v;
        renderSabhaList();
      }
    }, 300));

    // sabha filters
    ["#sabhaSearch", "#sabhaFilterDay", "#sabhaFilterArea", "#sabhaFilterKshetra", "#sabhaFilterType", "#sabhaSort"].forEach(function (sel) {
      $(sel).addEventListener("input", debounce(renderSabhaList, 150));
      $(sel).addEventListener("change", renderSabhaList);
    });
    $("#sabhaClearBtn").addEventListener("click", function () {
      $("#sabhaSearch").value = ""; $("#sabhaFilterDay").value = ""; $("#sabhaFilterArea").value = "";
      $("#sabhaFilterKshetra").value = ""; $("#sabhaFilterType").value = ""; $("#sabhaSort").value = "";
      renderSabhaList();
    });

    // karyakar filters
    ["#karyakarSearch", "#karyakarFilterdob"].forEach(function (sel) {
      $(sel).addEventListener("input", debounce(renderKaryakarList, 150));
      $(sel).addEventListener("change", renderKaryakarList);
    });
    $("#karyakarClearBtn").addEventListener("click", function () {
      $("#karyakarSearch").value = ""; $("#karyakarFilterdob").value = "";
      renderKaryakarList();
    });

    // balak filters
    ["#balakSearch", "#balakFilterSabha", "#balakFilterStd", "#balakFilterRegistered","#balakFilterdob"].forEach(function (sel) {
      $(sel).addEventListener("input", debounce(function () { STATE.balakPage = 1; renderBalakList(); }, 150));
      $(sel).addEventListener("change", function () { STATE.balakPage = 1; renderBalakList(); });
    });
    $("#balakClearBtn").addEventListener("click", function () {
      $("#balakSearch").value = ""; $("#balakFilterSabha").value = ""; $("#balakFilterStd").value = ""; $("#balakFilterRegistered").value = ""; $("#balakFilterdob").value = "";
      STATE.balakPage = 1;
      renderBalakList();
    });

    // modal close
    $("#dmCloseBtn").addEventListener("click", hideModal);
    $("#detailModal").addEventListener("click", function (e) { if (e.target === $("#detailModal")) hideModal(); });

    // back to top + floating buttons
    window.addEventListener("scroll", debounce(function () {
      $("#fabTop").classList.toggle("show", window.scrollY > 400);
    }, 50));
    $("#fabTop").addEventListener("click", function () { window.scrollTo({ top: 0, behavior: "smooth" }); });

    // delegated clicks (cards, global search results, actions)
    document.addEventListener("click", function (e) {
      const t = e.target.closest("[data-action]");
      if (!t) return;
      const action = t.getAttribute("data-action");

      if (action === "view-sabha") { e.preventDefault(); openSabhaModal(t.getAttribute("data-id")); }
      else if (action === "goto-sabha") { openSabhaModal(t.getAttribute("data-id")); $("#globalSearchResults").classList.remove("show"); $("#globalSearchInput").value = ""; }
      else if (action === "goto-karyakar") {
        goto("karyakar"); $("#karyakarSearch").value = IDX.karyakarById.get(idKey(t.getAttribute("data-id"))).name;
        renderKaryakarList(); $("#globalSearchResults").classList.remove("show"); $("#globalSearchInput").value = "";
      }
      else if (action === "goto-balak") {
        goto("balak");
        const b = DATA.balak.find(function (x) { return x.id === Number(t.getAttribute("data-id")); });
        $("#balakSearch").value = b ? (b.full_name || b.name) : "";
        STATE.balakPage = 1; renderBalakList();
        $("#globalSearchResults").classList.remove("show"); $("#globalSearchInput").value = "";
      }
      else if (action === "copy") { e.preventDefault(); copyToClipboard(t.getAttribute("data-value"), "Number"); }
      else if (action === "copy-map") { copyToClipboard(t.getAttribute("data-value"), "Map link"); }
      else if (action === "print-sabha") 
         { e.preventDefault();  shareSabha(t.getAttribute("data-id")); }
        // { alert("Printing is disabled.\n\nPlease contact the System Administrator.");}
      else if (action === "toggle-day") { t.closest(".day-accordion-item").classList.toggle("open"); }
      else if (action === "balak-prev") { STATE.balakPage = Math.max(1, STATE.balakPage - 1); renderBalakList(); scrollToBalakTop(); }
      else if (action === "balak-next") { STATE.balakPage += 1; renderBalakList(); scrollToBalakTop(); }
    });

    document.addEventListener("keydown", function (e) { if (e.key === "Escape") hideModal(); });
  }

  function scrollToBalakTop() {
    const el = $("#sec-balak");
    if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
  }

    /* ------------------------------------------------------------ */
  /* Share sabha details                                                     */
  /* ------------------------------------------------------------ */

  function shareSabha(id) {
    const sabha = IDX.sabhaById.get(idKey(id));
    if (!sabha) return;

    const karyakarCount =
        IDX.karyakarIdsBySabha.get(idKey(id))?.length || 0;

    const balakCount =
        IDX.balaksBySabha.get(idKey(id))?.length || 0;

    const text =
`🙏 *Malad East Bal Mandal*

🏠 *${sabha.name}*
📋 ${sabha.reg_no}

📅 *Day:* ${sabha.day}
🕢 *Time:* ${sabha.from_to_time}

📍 *Area:* ${sabha.area} (${sabha.kshetra})

👤 *Owner:*
${sabha.owner_name}

🏡 *Address:*
${sabha.address}

👥 *Karyakar:* ${karyakarCount}
👦 *Balak:* ${balakCount}

📞 Contact:
${sabha.contact1}

📍 Google Map:
${sabha.google_link}

🙏 Jay Swaminarayan`;

    
    if (navigator.share) {
    navigator.share({
        title: sabha.name,
        text: text
    });
   } else {
    window.open(
        "https://wa.me/?text=" + encodeURIComponent(text),
        "_blank"
    );
   }

}

  /* ------------------------------------------------------------ */
  /* Password                                                     */
  /* ------------------------------------------------------------ */
  const PASSWORD = "400097"; // Change this

document.getElementById("loginForm").addEventListener("submit", function (e) {
    e.preventDefault();

    const input = document.getElementById("loginPasswordInput").value;
    const error = document.getElementById("loginError");

    if (input === PASSWORD) {
        document.getElementById("loginOverlay").style.display = "none";
    } else {
        error.style.display = "block";
        document.getElementById("loginPasswordInput").value = "";
    }
});

  /* ------------------------------------------------------------ */
  /* INIT                                                           */
  /* ------------------------------------------------------------ */
  function init() {
    loadAllData().then(function () {
      $("#loadingOverlay").style.display = "none";
      renderHome();
      renderSummary();
      populateSabhaFilters();
      populateKaryakarFilters();
      populateBalakFilters();
      renderSabhaList();
      renderKaryakarList();
      renderBalakList();
      renderDayWise();
      wireEvents();
      goto("home");
    }).catch(function (err) {
      console.error(err);
      $("#loadingOverlay").innerHTML =
        '<div class="loading-spinner-wrap"><span class="material-symbols-outlined" style="font-size:40px;color:#c0392b;">error</span>' +
        "<p>Could not load data. Please make sure <code>data.js</code> is in the same folder as " +
        "<code>index.html</code> and is loaded before <code>script.js</code>.</p></div>";
    });
  }

  document.addEventListener("DOMContentLoaded", init);


    // Start ---- Copy / selection restriction (best-effort deterrent, not foolproof) ----
  document.addEventListener('copy', function(e){ e.preventDefault(); });
  document.addEventListener('cut', function(e){ e.preventDefault(); });
  document.addEventListener('contextmenu', function(e){ e.preventDefault(); });
  document.addEventListener('selectstart', function(e){
    if (e.target.tagName !== 'A') e.preventDefault();
  });
  document.addEventListener('dragstart', function(e){ e.preventDefault(); });
  document.addEventListener('keydown', function(e){
    var k = e.key ? e.key.toLowerCase() : '';
    if ((e.ctrlKey || e.metaKey) && ['c','x','a','s','p','u'].indexOf(k) !== -1) {
      e.preventDefault();
    }
    if (e.key === 'F12') e.preventDefault();
      });
  // End ---- Copy / selection restriction (best-effort deterrent, not foolproof) ----



})();