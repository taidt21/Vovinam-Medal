(function ($) {
  "use strict";

  const LS_KEY = "vovinam.medals.v2.1";
  const APP_VERSION = 2;

  // ===== Focus-lock để không mất con trỏ khi gõ =====
  let isTyping = false;

  // ===== Utils chung =====
  function normalizeStr(str) {
    return (str || "")
      .toString()
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
  }

  function uid() {
    return "a" + Math.random().toString(36).slice(2, 9);
  }

  function escapeHtml(s) {
    return (s + "").replace(
      /[&<>"']/g,
      (m) =>
        ({
          "&": "&amp;",
          "<": "&lt;",
          ">": "&gt;",
          '"': "&quot;",
          "'": "&#39;",
        }[m])
    );
  }

  function escapeAttr(s) {
    return escapeHtml(s).replace(/"/g, "&quot;");
  }

  function toast(msg, isErr) {
    const el = $("<div/>")
      .text(msg)
      .css({
        position: "fixed",
        bottom: "24px",
        right: "24px",
        padding: "10px 14px",
        background: isErr ? "#3b0d0d" : "#062e2e",
        border: "1px solid #0d4d4d",
        color: "#e5f5f5",
        borderRadius: "10px",
        zIndex: 9999,
        boxShadow: "0 10px 30px rgba(0,0,0,.35)",
      });
    $("body").append(el);
    setTimeout(() => el.fadeOut(300, () => el.remove()), 1800);
  }

  // ===== Mode & trọng số =====
  function mode() {
    return $('input[name="calcMode"]:checked').val() || "score";
  }

  function applyModeUI() {
    const m = mode();
    if (m === "count") {
      $(".col-pts").hide();
    } else {
      $(".col-pts").show();
    }
    const $weightsRow = $("#wGold, #wSilver, #wBronze").closest("label");
    $weightsRow.css("opacity", m === "count" ? 0.5 : 1);
  }

  function weights() {
    return {
      g: parseInt($("#wGold").val() || 3, 10),
      s: parseInt($("#wSilver").val() || 2, 10),
      b: parseInt($("#wBronze").val() || 1, 10),
    };
  }

  // ===== Dòng VĐV =====
  function rowData($tr) {
    const g = parseInt($tr.find(".inp-medal.g").val() || 0, 10);
    const s = parseInt($tr.find(".inp-medal.s").val() || 0, 10);
    const b = parseInt($tr.find(".inp-medal.b").val() || 0, 10);
    const w = weights();
    const pts = g * w.g + s * w.s + b * w.b;
    const sum = g + s + b;
    const name = $tr.find(".ath-name").text().trim();
    const yob = ($tr.find(".inp-yob").val() || "").trim();
    const team = $tr.find(".inp-team").val().trim();
    return { id: $tr.data("id"), name, team, yob, g, s, b, sum, pts };
  }

  function recalcAthleteRow($tr) {
    const d = rowData($tr);
    $tr.find(".sum").text(d.sum);
    $tr.find(".pts").text(d.pts);
    $tr.attr("data-team", d.team);
    return d;
  }

  function makeRowHtml(id, name, team, yob, g, s, b) {
    return `
      <tr data-id="${id}" data-team="${escapeAttr(team || "")}">
        <td class="rank">–</td>
        <td contenteditable="true" class="ath-name">${escapeHtml(
          name || "Tên VĐV"
        )}</td>
        <td><input class="inp-yob" type="number" min="1900" max="2100" value="${escapeAttr(
          yob || ""
        )}" /></td>
        <td><input class="inp-team" type="text" value="${escapeAttr(
          team || ""
        )}" /></td>
        <td><input class="inp-medal g" type="number" min="0" value="${
          g ?? 0
        }" /></td>
        <td><input class="inp-medal s" type="number" min="0" value="${
          s ?? 0
        }" /></td>
        <td><input class="inp-medal b" type="number" min="0" value="${
          b ?? 0
        }" /></td>
        <td class="sum">0</td>
        <td class="pts col-pts">0</td>
        <td class="medal-ctrl">
          <button class="btn-inc" data-type="G">+V</button>
          <button class="btn-inc" data-type="S">+B</button>
          <button class="btn-inc" data-type="B">+Đ</button>
        </td>
        <td><button class="btnDelRow danger">Xoá</button></td>
      </tr>`;
  }

  // ===== BXH VĐV & Đoàn =====
  function sortAthletes() {
    const $tbody = $("#tblAthletes tbody");
    const $rows = $tbody.children("tr").get();
    const m = mode();

    $rows.sort((a, b) => {
      const da = rowData($(a));
      const db = rowData($(b));

      if (m === "score") {
        if (db.pts !== da.pts) return db.pts - da.pts;
      }
      if (db.g !== da.g) return db.g - da.g;
      if (db.s !== da.s) return db.s - da.s;
      if (db.b !== da.b) return db.b - da.b;

      return da.name.localeCompare(db.name, "vi", { sensitivity: "base" });
    });

    $rows.forEach((tr, i) => {
      tr.querySelector(".rank").textContent = i + 1;
      $tbody.append(tr);
    });
  }

  function aggregateTeams() {
    const map = new Map();
    $("#tblAthletes tbody tr").each(function () {
      const d = rowData($(this));
      if (!d.team) return;
      if (!map.has(d.team)) {
        map.set(d.team, { team: d.team, g: 0, s: 0, b: 0, sum: 0, pts: 0 });
      }
      const t = map.get(d.team);
      t.g += d.g;
      t.s += d.s;
      t.b += d.b;
      t.sum += d.sum;
      t.pts += d.pts;
    });

    const m = mode();
    return Array.from(map.values()).sort((a, b) => {
      if (m === "score") {
        if (b.pts !== a.pts) return b.pts - a.pts;
      }
      if (b.g !== a.g) return b.g - a.g;
      if (b.s !== a.s) return b.s - a.s;
      if (b.b !== a.b) return b.b - a.b;
      return a.team.localeCompare(b.team, "vi", { sensitivity: "base" });
    });
  }

  function renderTeams() {
    const teams = aggregateTeams();
    const $tb = $("#tblTeams tbody").empty();
    teams.forEach((t, idx) => {
      const tr = `
        <tr>
          <td class="rank">${idx + 1}</td>
          <td>${t.team}</td>
          <td>${t.g}</td>
          <td>${t.s}</td>
          <td>${t.b}</td>
          <td>${t.sum}</td>
          <td class="pts col-pts">${t.pts}</td>
        </tr>`;
      $tb.append(tr);
    });
  }

  function recalcAll() {
    $("#tblAthletes tbody tr").each(function () {
      recalcAthleteRow($(this));
    });
    if (!isTyping) {
      sortAthletes();
    }
    renderTeams();
    applyModeUI();
  }

  // ===== State (localStorage / JSON) =====
  function collectState() {
    return {
      version: APP_VERSION,
      mode: mode(),
      weights: weights(),
      athletes: $("#tblAthletes tbody tr")
        .map(function () {
          const d = rowData($(this));
          return {
            id: d.id,
            name: d.name,
            team: d.team,
            yob: d.yob,
            g: d.g,
            s: d.s,
            b: d.b,
          };
        })
        .get(),
    };
  }

  function applyState(data) {
    if (!data) return;

    if (data.mode) {
      $('input[name="calcMode"][value="' + data.mode + '"]').prop(
        "checked",
        true
      );
    }

    const w = data.weights || data.w;
    if (w) {
      $("#wGold").val(w.g ?? 3);
      $("#wSilver").val(w.s ?? 2);
      $("#wBronze").val(w.b ?? 1);
    }

    const list = data.athletes || data.rows || [];
    const $tb = $("#tblAthletes tbody").empty();
    list.forEach((r) => {
      $tb.append(
        makeRowHtml(
          r.id || uid(),
          r.name || "",
          r.team || "",
          r.yob || "",
          r.g || 0,
          r.s || 0,
          r.b || 0
        )
      );
    });

    recalcAll();
  }

  function saveToLocal() {
    try {
      const payload = collectState();
      localStorage.setItem(LS_KEY, JSON.stringify(payload));
    } catch (e) {
      console.error("Save local failed", e);
      toast(
        "Không lưu được vào máy. Có thể localStorage bị chặn hoặc đầy.",
        true
      );
    }
  }

  function loadFromLocal() {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) {
      recalcAll();
      return;
    }
    try {
      const data = JSON.parse(raw);
      applyState(data);
    } catch (e) {
      console.warn("Load local failed", e);
      recalcAll();
    }
  }

  function dumpJson() {
    return collectState();
  }

  // ===== Import Excel / CSV =====
  function isSameAthlete(aName, aTeam, aYob, bName, bTeam, bYob) {
    const na = normalizeStr(aName);
    const ta = normalizeStr(aTeam);
    const nb = normalizeStr(bName);
    const tb = normalizeStr(bTeam);

    if (!na || !nb || na !== nb) return false;
    if (ta !== tb) return false;

    const ya = (aYob || "").toString().trim();
    const yb = (bYob || "").toString().trim();

    // 1 trong 2 không có năm sinh -> coi là trùng
    if (!ya || !yb) return true;

    // Cả 2 đều có năm sinh: chỉ trùng nếu bằng nhau
    return ya === yb;
  }

  // Kiểm tra 1 VĐV đã tồn tại trong list hay chưa
  function existsInList(list, name, team, yob) {
    for (const it of list) {
      if (isSameAthlete(name, team, yob, it.name, it.team, it.yob)) {
        return true;
      }
    }
    return false;
  }

  // ===== Import Excel / CSV =====
  function importAthletesFrom2D(rows) {
    if (!rows || !rows.length) {
      toast("File không có dữ liệu.", true);
      return;
    }

    const $tb = $("#tblAthletes tbody");

    // Lấy danh sách VĐV đang có trên bảng (trước khi import)
    const existingList = [];
    $tb.find("tr").each(function () {
      const $tr = $(this);
      const d = rowData($tr); // đã có name, team, yob bên trong
      if (d.name) {
        existingList.push({
          name: d.name,
          team: d.team,
          yob: d.yob,
        });
      }
    });

    // List VĐV được thêm trong LẦN IMPORT NÀY (chống trùng trong chính file)
    const sessionList = [];

    let added = 0;
    let skipped = 0;

    rows.forEach((row, idx) => {
      if (!row || !row.length) return;

      const rawName = (row[0] || "").toString().trim(); // cột A: Tên VĐV
      const rawTeam = (row[1] || "").toString().trim(); // cột B: Đoàn
      const rawYob = (row[2] || "").toString().trim(); // cột C: Năm sinh (có cũng được, không có cũng ok)

      if (!rawName) return;

      // Bỏ qua header
      if (idx === 0) {
        const lowerName = rawName.toLowerCase();
        if (
          lowerName.includes("tên") ||
          lowerName.includes("họ tên") ||
          lowerName.includes("họ tên") ||
          lowerName.includes("name")
        ) {
          return;
        }
      }

      // Nếu đã tồn tại trên bảng HOẶC trong lần import này -> bỏ qua
      if (
        existsInList(existingList, rawName, rawTeam, rawYob) ||
        existsInList(sessionList, rawName, rawTeam, rawYob)
      ) {
        skipped++;
        return;
      }

      const id = uid();
      $tb.append(makeRowHtml(id, rawName, rawTeam, rawYob, 0, 0, 0));

      sessionList.push({
        name: rawName,
        team: rawTeam,
        yob: rawYob,
      });

      added++;
    });

    if (added > 0) {
      recalcAll();
      saveToLocal();
    }

    if (!added && !skipped) {
      toast("Không tìm thấy dòng VĐV hợp lệ trong file.", true);
    } else {
      const msg = skipped
        ? `Đã thêm ${added} VĐV mới. Bỏ qua ${skipped} dòng trùng theo Tên/Đoàn/Năm sinh.`
        : `Đã thêm ${added} VĐV mới từ file.`;
      toast(msg);
    }
  }

  function handleImportFile(file) {
    if (!file) return;

    const ext = (file.name.split(".").pop() || "").toLowerCase();

    // CSV
    if (ext === "csv") {
      const reader = new FileReader();
      reader.onload = function (ev) {
        const text = ev.target.result || "";
        const lines = text.split(/\r?\n/);
        const rows = lines
          .map((line) => line.split(/,|;|\t/))
          .filter((r) => r.some((cell) => (cell || "").trim() !== ""));
        importAthletesFrom2D(rows);
      };
      reader.readAsText(file, "utf-8");
      return;
    }

    // Excel
    if (typeof XLSX === "undefined") {
      toast(
        "Không tìm thấy thư viện XLSX. Kiểm tra lại thẻ script import XLSX.",
        true
      );
      return;
    }

    const reader = new FileReader();
    reader.onload = function (ev) {
      const data = new Uint8Array(ev.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        toast("File Excel không có sheet nào.", true);
        return;
      }

      const firstSheetName = workbook.SheetNames[0];
      const ws = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

      importAthletesFrom2D(rows);
    };
    reader.readAsArrayBuffer(file);
  }

  // ===== CSV Export =====
  function csvEsc(s) {
    s = (s ?? "").toString();
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  }

  function buildCsv() {
    const m = mode();
    let a =
      m === "score"
        ? "Rank,Name,BirthYear,Team,Gold,Silver,Bronze,Total,Points\n"
        : "Rank,Name,BirthYear,Team,Gold,Silver,Bronze,Total\n";

    $("#tblAthletes tbody tr").each(function () {
      const d = rowData($(this));
      const rank = $(this).find(".rank").text().trim() || "";
      a +=
        m === "score"
          ? [
              rank,
              csvEsc(d.name),
              csvEsc(d.yob),
              csvEsc(d.team),
              d.g,
              d.s,
              d.b,
              d.sum,
              d.pts,
            ].join(",") + "\n"
          : [
              rank,
              csvEsc(d.name),
              csvEsc(d.yob),
              csvEsc(d.team),
              d.g,
              d.s,
              d.b,
              d.sum,
            ].join(",") + "\n";
    });

    let t =
      m === "score"
        ? "Rank,Team,Gold,Silver,Bronze,Total,Points\n"
        : "Rank,Team,Gold,Silver,Bronze,Total\n";

    $("#tblTeams tbody tr").each(function () {
      const rank = $(this).find(".rank").text().trim();
      const cells = $(this).children("td");
      const team = cells.eq(1).text().trim();
      const g = cells.eq(2).text().trim();
      const s = cells.eq(3).text().trim();
      const b = cells.eq(4).text().trim();
      const sum = cells.eq(5).text().trim();
      const pts = cells.eq(6).text().trim();
      t +=
        m === "score"
          ? [rank, csvEsc(team), g, s, b, sum, pts].join(",") + "\n"
          : [rank, csvEsc(team), g, s, b, sum].join(",") + "\n";
    });

    return { csvAthletes: a, csvTeams: t };
  }

  function download(filename, text) {
    const BOM = "\uFEFF";
    const blob = new Blob([BOM + text], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  // ===== Events =====
  // Focus-lock
  $(document).on(
    "focusin",
    ".ath-name, .inp-team, .inp-yob, .inp-medal",
    function () {
      isTyping = true;
    }
  );

  $(document).on(
    "blur",
    ".ath-name, .inp-team, .inp-yob, .inp-medal",
    function () {
      isTyping = false;
      recalcAll();
      saveToLocal();
    }
  );

  // Thay đổi số/đội/trọng số
  $(document).on(
    "input change",
    ".inp-medal, .inp-team, .inp-yob, #wGold, #wSilver, #wBronze",
    function () {
      const $tr = $(this).closest("tr");
      if ($tr.length) {
        recalcAthleteRow($tr);
      }
      if (!isTyping) {
        sortAthletes();
      }
      renderTeams();
      saveToLocal();
    }
  );

  // Tên VĐV: recalc, sort khi blur (ở trên)
  $(document).on("input", ".ath-name", function () {
    const $tr = $(this).closest("tr");
    recalcAthleteRow($tr);
    renderTeams();
  });

  // Nút tăng nhanh
  $(document).on("click", ".btn-inc", function () {
    const type = $(this).data("type"); // G S B
    const $tr = $(this).closest("tr");
    const cls = type === "G" ? ".g" : type === "S" ? ".s" : ".b";
    const $inp = $tr.find(".inp-medal" + cls);
    const val = parseInt($inp.val() || 0, 10) + 1;
    $inp.val(val);
    isTyping = false;
    recalcAll();
    saveToLocal();
  });

  // Thêm / xoá dòng
  $("#btnAddRow").on("click", function () {
    $("#tblAthletes tbody").append(
      makeRowHtml(uid(), "VĐV mới", "", "", 0, 0, 0)
    );
    recalcAll();
    saveToLocal();
  });

  $(document).on("click", ".btnDelRow", function () {
    $(this).closest("tr").remove();
    recalcAll();
    saveToLocal();
  });

  // Tìm kiếm
  $("#search").on("input", function () {
    const q = normalizeStr($(this).val());
    $("#tblAthletes tbody tr").each(function () {
      const name = normalizeStr($(this).find(".ath-name").text());
      const team = normalizeStr($(this).find(".inp-team").val());
      const yob = normalizeStr($(this).find(".inp-yob").val());
      $(this).toggle(
        !q || name.includes(q) || team.includes(q) || yob.includes(q)
      );
    });
  });

  // Lưu / reset local
  $("#btnSave").on("click", function () {
    saveToLocal();
    toast("Đã lưu vào máy.");
  });

  $("#btnReset").on("click", function () {
    if (confirm("Xoá toàn bộ dữ liệu đang lưu trong máy?")) {
      localStorage.removeItem(LS_KEY);
      location.reload();
    }
  });

  // JSON xuất / nhập
  $("#btnDump").on("click", function () {
    $("#ioJson").val(JSON.stringify(dumpJson(), null, 2));
    toast("Đã xuất JSON bên dưới.");
  });

  $("#btnLoad").on("click", function () {
    const raw = $("#ioJson").val().trim();
    if (!raw) {
      toast("Chưa có JSON để nhập.", true);
      return;
    }
    try {
      const data = JSON.parse(raw);
      applyState(data);
      saveToLocal();
      toast("Đã nhập JSON.");
    } catch (e) {
      toast("JSON không hợp lệ.", true);
    }
  });

  // Export CSV
  $("#btnExportCsv").on("click", function () {
    const { csvAthletes, csvTeams } = buildCsv();
    download("athletes.csv", csvAthletes);
    download("teams.csv", csvTeams);
  });

  // Đổi mode tính
  $(document).on("change", 'input[name="calcMode"]', function () {
    applyModeUI();
    recalcAll();
    saveToLocal();
  });

  // Import Excel
  $("#btnImport").on("click", function () {
    $("#fileImport").trigger("click");
  });

  $("#fileImport").on("change", function (e) {
    const file = e.target.files[0];
    if (file) {
      handleImportFile(file);
    }
    $(this).val("");
  });

  // ===== init =====
  $(function () {
    loadFromLocal();
    applyModeUI();
  });
})(jQuery);
