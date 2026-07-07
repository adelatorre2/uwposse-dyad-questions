/* ============================================================
   Get to Know Your Posse — app logic (vanilla JS)
   Reads data/questions.xlsx at runtime via SheetJS. No build step:
   edit the Excel file, push, and the site updates.
   ============================================================ */

(function () {
  "use strict";

  /* --------- Config --------- */
  var DATA_URL = "data/questions.xlsx";
  var SHEET_NAME = "Questions";

  /* Brand-anchored category palette. Colors are assigned to categories in
     order of first appearance in the spreadsheet. It leads with the Badger
     reds and adds accessible complementary accents so ~11 categories stay
     visually distinct while still reading as UW-branded. Every color is dark
     enough that WHITE tag text clears WCAG AA (>= 4.5:1). If there are more
     categories than colors, it wraps around. Categories are shown as clean
     color-coded TEXT tags (no icons/emoji), so a brand-new category added in
     Excel renders perfectly with no extra mapping. */
  var CATEGORY_PALETTE = [
    "#C5050C", // Badger red
    "#14607F", // teal-blue
    "#2E7D46", // green
    "#5B3A8E", // plum
    "#A8480B", // burnt orange
    "#334A5E", // slate
    "#A0285A", // raspberry
    "#1F3A93", // navy
    "#6B4423", // coffee brown
    "#0F6E62", // pine
    "#9B0000", // dark red (brand)
    "#6D28D9"  // violet (spare / wraparound)
  ];

  /* --------- State --------- */
  var allQuestions = [];      // { category, depth, question }
  var categories = [];        // ordered unique category names
  var categoryMeta = {};      // name -> { color }
  var activeCategory = "All"; // "All" or a category name
  var activeDepth = "any";    // "any" | "warm" | "deep"
  var mode = "single";        // "single" | "browse"
  var lastQuestionText = null;

  /* --------- DOM --------- */
  var $ = function (id) { return document.getElementById(id); };
  var statusEl = $("status");
  var catChipsEl = $("category-chips");
  var singleModeEl = $("single-mode");
  var browseModeEl = $("browse-mode");

  /* ============================================================
     Loading + parsing
     ============================================================ */
  function loadQuestions() {
    // Cache-busting so edits to the xlsx appear promptly.
    var url = DATA_URL + "?v=" + Date.now();
    fetch(url, { cache: "no-store" })
      .then(function (res) {
        if (!res.ok) throw new Error("HTTP " + res.status);
        return res.arrayBuffer();
      })
      .then(function (buf) {
        var rows = parseWorkbook(buf);
        if (!rows.length) {
          showError(
            "No questions found. Check that <code>data/questions.xlsx</code> has a " +
            "sheet named <strong>Questions</strong> with <em>category</em>, " +
            "<em>depth</em>, and <em>question</em> columns."
          );
          return;
        }
        allQuestions = rows;
        buildCategories();
        clearStatus();
        renderCategoryChips();
        renderCurrentMode();
      })
      .catch(function (err) {
        showError(
          "Couldn't load the questions file. Make sure " +
          "<code>data/questions.xlsx</code> exists and try refreshing." +
          '<br><span class="err-detail">(' + escapeHtml(String(err.message || err)) + ")</span>"
        );
      });
  }

  function parseWorkbook(buf) {
    var out = [];
    var wb;
    try {
      wb = XLSX.read(new Uint8Array(buf), { type: "array" });
    } catch (e) {
      showError("The Excel file couldn't be read. It may be corrupted — re-save it and push again.");
      return out;
    }
    // Prefer the "Questions" sheet; fall back to the first sheet.
    var sheet = wb.Sheets[SHEET_NAME] || wb.Sheets[wb.SheetNames[0]];
    if (!sheet) return out;

    var raw;
    try {
      raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    } catch (e) {
      return out;
    }

    for (var i = 0; i < raw.length; i++) {
      var row = raw[i];
      var rec = normalizeRow(row);
      if (rec) out.push(rec);
    }
    return out;
  }

  // Case/space-insensitive column lookup so header typos don't break things.
  function pick(row, keys) {
    var map = {};
    for (var k in row) {
      if (Object.prototype.hasOwnProperty.call(row, k)) {
        map[String(k).trim().toLowerCase()] = row[k];
      }
    }
    for (var j = 0; j < keys.length; j++) {
      var key = keys[j];
      if (map[key] !== undefined) return map[key];
    }
    return "";
  }

  function normalizeRow(row) {
    var question = String(pick(row, ["question"]) || "").trim();
    if (!question) return null; // skip empty / malformed rows

    var category = String(pick(row, ["category"]) || "").trim() || "Uncategorized";

    var depthRaw = pick(row, ["depth"]);
    var depth = parseInt(depthRaw, 10);
    if (isNaN(depth) || depth < 1) depth = 1;   // blank / invalid defaults to 1
    if (depth > 4) depth = 4;

    return { category: category, depth: depth, question: question };
  }

  function buildCategories() {
    categories = [];
    categoryMeta = {};
    for (var i = 0; i < allQuestions.length; i++) {
      var cat = allQuestions[i].category;
      if (!categoryMeta[cat]) {
        var color = CATEGORY_PALETTE[categories.length % CATEGORY_PALETTE.length];
        categoryMeta[cat] = { color: color };
        categories.push(cat);
      }
    }
    // Reset category filter if it no longer exists.
    if (activeCategory !== "All" && categories.indexOf(activeCategory) === -1) {
      activeCategory = "All";
    }
  }

  /* ============================================================
     Filtering
     ============================================================ */
  function depthMatches(depth) {
    if (activeDepth === "warm") return depth <= 2;
    if (activeDepth === "deep") return depth >= 3;
    return true; // "any"
  }

  function filtered() {
    return allQuestions.filter(function (q) {
      if (activeCategory !== "All" && q.category !== activeCategory) return false;
      return depthMatches(q.depth);
    });
  }

  /* ============================================================
     Rendering — filters
     ============================================================ */
  function renderCategoryChips() {
    catChipsEl.innerHTML = "";
    catChipsEl.appendChild(makeCatChip("All", "All", null));
    for (var i = 0; i < categories.length; i++) {
      var cat = categories[i];
      catChipsEl.appendChild(makeCatChip(cat, cat, categoryMeta[cat]));
    }
    updateChipStates();
  }

  function makeCatChip(value, label, meta) {
    var btn = document.createElement("button");
    btn.className = "chip category-chip";
    btn.dataset.category = value;
    if (meta) {
      // Color-coded text chip: a small color dot + the category name (no icons).
      btn.style.setProperty("--chip-color", meta.color);
      btn.innerHTML = '<span class="chip-dot" aria-hidden="true"></span>' + escapeHtml(label);
    } else {
      btn.textContent = label;
    }
    btn.addEventListener("click", function () {
      activeCategory = value;
      updateChipStates();
      renderCurrentMode();
    });
    return btn;
  }

  function updateChipStates() {
    var catBtns = catChipsEl.querySelectorAll(".category-chip");
    catBtns.forEach(function (b) {
      b.classList.toggle("is-active", b.dataset.category === activeCategory);
    });
    var depthBtns = document.querySelectorAll(".depth-chip");
    depthBtns.forEach(function (b) {
      b.classList.toggle("is-active", b.dataset.depth === activeDepth);
    });
  }

  function wireDepthChips() {
    var depthBtns = document.querySelectorAll(".depth-chip");
    depthBtns.forEach(function (b) {
      b.addEventListener("click", function () {
        activeDepth = b.dataset.depth;
        updateChipStates();
        renderCurrentMode();
      });
    });
  }

  /* ============================================================
     Rendering — single mode
     ============================================================ */
  function depthDots(depth, color) {
    var html = '<span class="depth-dots" title="Depth ' + depth + ' of 4">';
    for (var i = 1; i <= 4; i++) {
      var filled = i <= depth ? " filled" : "";
      var style = i <= depth && color ? ' style="--dot-color:' + color + '"' : "";
      html += '<span class="dot' + filled + '"' + style + "></span>";
    }
    html += "</span>";
    return html;
  }

  function pickRandom(pool) {
    if (!pool.length) return null;
    if (pool.length === 1) return pool[0];
    var choice;
    var guard = 0;
    do {
      choice = pool[Math.floor(Math.random() * pool.length)];
      guard++;
    } while (choice.question === lastQuestionText && guard < 12);
    return choice;
  }

  function renderSingle() {
    var pool = filtered();
    var countEl = $("single-count");
    if (!pool.length) {
      $("single-cat-tag").textContent = "";
      $("single-cat-tag").style.display = "none";
      $("single-depth-dots").innerHTML = "";
      $("single-question").textContent = "No questions match these filters. Try “Shuffle / reset”.";
      $("qcard").style.setProperty("--card-accent", "var(--badger-red)");
      countEl.textContent = "";
      return;
    }
    var q = pickRandom(pool);
    lastQuestionText = q.question;
    var meta = categoryMeta[q.category] || { color: "#C5050C" };

    var card = $("qcard");
    card.classList.add("is-swapping");

    window.setTimeout(function () {
      var tag = $("single-cat-tag");
      tag.style.display = "";
      tag.style.background = meta.color;
      tag.textContent = q.category;

      $("single-depth-dots").innerHTML = depthDots(q.depth, meta.color);
      $("single-question").textContent = q.question;
      card.style.setProperty("--card-accent", meta.color);
      card.classList.remove("is-swapping");
    }, 120);

    countEl.textContent = pool.length + (pool.length === 1 ? " question" : " questions") + " in this mix";
  }

  /* ============================================================
     Rendering — browse mode
     ============================================================ */
  function renderBrowse() {
    browseModeEl.innerHTML = "";
    var pool = filtered();
    if (!pool.length) {
      browseModeEl.innerHTML =
        '<p class="count-hint">No questions match these filters. Try “Shuffle / reset”.</p>';
      return;
    }
    // Group by category, preserving category first-appearance order.
    for (var c = 0; c < categories.length; c++) {
      var cat = categories[c];
      if (activeCategory !== "All" && activeCategory !== cat) continue;
      var items = pool.filter(function (q) { return q.category === cat; });
      if (!items.length) continue;

      var meta = categoryMeta[cat];
      var group = document.createElement("section");
      group.className = "cat-group";
      group.style.setProperty("--cat-color", meta.color);

      var header = document.createElement("div");
      header.className = "cat-group-header";
      header.innerHTML =
        '<span class="cat-swatch" aria-hidden="true"></span>' +
        "<h2>" + escapeHtml(cat) + "</h2>" +
        '<span class="cat-count">' + items.length + "</span>";
      group.appendChild(header);

      var list = document.createElement("ul");
      list.className = "qlist";
      for (var i = 0; i < items.length; i++) {
        var li = document.createElement("li");
        li.className = "qitem";
        li.innerHTML =
          '<p class="qitem-text">' + escapeHtml(items[i].question) + "</p>" +
          depthDots(items[i].depth, meta.color);
        list.appendChild(li);
      }
      group.appendChild(list);
      browseModeEl.appendChild(group);
    }
  }

  function renderCurrentMode() {
    if (mode === "single") renderSingle();
    else renderBrowse();
  }

  /* ============================================================
     Mode toggle + actions
     ============================================================ */
  function setMode(next) {
    mode = next;
    var isSingle = mode === "single";
    $("mode-single").classList.toggle("is-active", isSingle);
    $("mode-browse").classList.toggle("is-active", !isSingle);
    $("mode-single").setAttribute("aria-selected", String(isSingle));
    $("mode-browse").setAttribute("aria-selected", String(!isSingle));
    singleModeEl.hidden = !isSingle;
    browseModeEl.hidden = isSingle;
    renderCurrentMode();
  }

  function wireControls() {
    $("mode-single").addEventListener("click", function () { setMode("single"); });
    $("mode-browse").addEventListener("click", function () { setMode("browse"); });
    $("next-btn").addEventListener("click", function () { renderSingle(); });
    $("shuffle-btn").addEventListener("click", function () {
      activeCategory = "All";
      activeDepth = "any";
      lastQuestionText = null;
      updateChipStates();
      // scroll chip rows back to start
      catChipsEl.scrollLeft = 0;
      renderCurrentMode();
    });
    wireDepthChips();
  }

  /* ============================================================
     Status helpers
     ============================================================ */
  function showError(html) {
    statusEl.innerHTML = '<div class="msg">' + html + "</div>";
    // Keep the single-mode card from saying "Loading…" forever.
    var sq = $("single-question");
    if (sq && sq.textContent === "Loading questions…") {
      sq.textContent = "Couldn’t load questions — see the note above.";
    }
  }
  function clearStatus() { statusEl.innerHTML = ""; }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  /* --------- Boot --------- */
  document.addEventListener("DOMContentLoaded", function () {
    wireControls();
    loadQuestions();
  });
})();
