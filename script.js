/** @format */

(function ($) {
  "use strict";

  const LS_KEY = "vovinam.scores.v2";

  const DEFAULT_MEDAL_WEIGHTS = {
    gold: 10,
    silver: 6,
    bronze: 3,
  };

  function getMedalWeights() {
    const read = (selector, fallback) => {
      const $el = $(selector);
      if ($el.length === 0) return fallback;
      const raw = $el.val();
      if (raw === undefined || raw === null || raw === "") return fallback;
      const num = parseFloat(raw.toString().replace(",", "."));
      return isNaN(num) || num <= 0 ? fallback : num;
    };

    return {
      gold: read("#coefGold", DEFAULT_MEDAL_WEIGHTS.gold),
      silver: read("#coefSilver", DEFAULT_MEDAL_WEIGHTS.silver),
      bronze: read("#coefBronze", DEFAULT_MEDAL_WEIGHTS.bronze),
    };
  }

  function setMedalWeightInputs(weights) {
    const w = weights || DEFAULT_MEDAL_WEIGHTS;
    if ($("#coefGold").length) {
      $("#coefGold").val(w.gold);
      $("#coefSilver").val(w.silver);
      $("#coefBronze").val(w.bronze);
    }
  }

  // ======================
  // Utils
  // ======================
  function uid() {
    return "r" + Math.random().toString(36).slice(2, 9);
  }

  function normalizeStr(str) {
    return (str || "")
      .toString()
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "") // bỏ dấu
      .replace(/đ/g, "d"); // chuyển đ -> d để bắt được "dong doi"
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
        boxShadow: "0 10px 30px rgba(0, 0, 0, 0.35)",
      });
    $("body").append(el);
    setTimeout(() => el.fadeOut(300, () => el.remove()), 1800);
  }

  // ======================
  // Create row HTML
  // type: "single" | "master" | "member"
  // ======================
  function makeRowHtml(row) {
    const id = row.id || uid();
    const type = row.type || "single";
    const parentId = row.parentId || "";

    const isMaster = type === "master";
    const isMember = type === "member";

    const age = row.age || "";
    const ev = row.event || "";
    const name = row.name || "";
    const yob = row.yob || "";
    const team = row.team || "";
    const g1 = row.g1 ?? "";
    const g2 = row.g2 ?? "";
    const g3 = row.g3 ?? "";
    const g4 = row.g4 ?? "";
    const g5 = row.g5 ?? "";
    const total = row.total ?? "";
    const rank = row.rank ?? "";
    const medal = row.medal || "";

    const rowClass =
      type === "master"
        ? "team-master-row"
        : type === "member"
        ? "team-member-row"
        : "";

    // Age cell
    const ageCell = isMember
      ? `<td>
           <span class="sub-age">${age}</span>
           <input type="hidden" class="inp-age" value="${age}" />
         </td>`
      : `<td><input class="inp-age" type="text" value="${age}" /></td>`;

    // Event cell
    const eventCell = isMember
      ? `<td>
           <span class="sub-event">${ev}</span>
           <input type="hidden" class="inp-event" value="${ev}" />
         </td>`
      : `<td><input class="inp-event" type="text" value="${ev}" /></td>`;

    // Name cell
    const nameCell =`<td contenteditable="true" class="ath-name ${isMember ? " is-member" : ""}">
    <div>${name || ""}</div>
    </td>`;

    // YOB cell
    const yobCell = isMember
      ? `<td>
           <span class="sub-yob">${yob}</span>
           <input type="hidden" class="inp-yob" value="${yob}" />
         </td>`
      : `<td><input class="inp-yob" type="number" value="${yob}" /></td>`;

    // Team cell
    const teamCell = isMember
      ? `<td>
           <span class="sub-team">${team}</span>
           <input type="hidden" class="inp-team" value="${team}" />
         </td>`
      : `<td><input class="inp-team" type="text" value="${team}" /></td>`;

    // Score cells
    const scoreCells = isMember
      ? `
        <td class="sub-empty"></td>
        <td class="sub-empty"></td>
        <td class="sub-empty"></td>
        <td class="sub-empty"></td>
        <td class="sub-empty"></td>
        <td class="total"></td>
        <td class="rank-num"></td>
        <td class="medal"></td>
      `
      : `
        <td><input class="inp-g g1" type="number" step="0.1" value="${g1}" /></td>
        <td><input class="inp-g g2" type="number" step="0.1" value="${g2}" /></td>
        <td><input class="inp-g g3" type="number" step="0.1" value="${g3}" /></td>
        <td><input class="inp-g g4" type="number" step="0.1" value="${g4}" /></td>
        <td><input class="inp-g g5" type="number" step="0.1" value="${g5}" /></td>
        <td class="total">${total}</td>
        <td class="rank-num">${rank}</td>
        <td class="medal">${medal}</td>
      `;

    return `
      <tr data-id="${id}" data-type="${type}" ${
      parentId ? `data-parent="${parentId}"` : ""
    } class="${rowClass}">
        <td class="rank stt">-</td>
        ${ageCell}
        ${eventCell}
        ${nameCell}
        ${yobCell}
        ${teamCell}
        ${scoreCells}
        <td><button class="btnDelRow danger">Xoá</button></td>
      </tr>
    `;
  }

  // ======================
  // Compute total score for a row
  // ======================
  function computeTotalFromRow($tr) {
    const type = $tr.data("type") || "single";
    if (type === "member") {
      $tr.find(".total").text("");
      return 0;
    }

    function readScore(cls) {
      const v = $tr.find(cls).val();
      if (v === undefined || v === null || v === "") return null;
      const num = parseFloat(v.toString().replace(",", "."));
      return isNaN(num) ? null : num;
    }

    const s1 = readScore(".g1");
    const s2 = readScore(".g2");
    const s3 = readScore(".g3");
    const s4 = readScore(".g4");
    const s5 = readScore(".g5");
    const scores = [s1, s2, s3, s4, s5].filter((v) => v !== null);

    if (!scores.length) {
      $tr.find(".total").text("");
      return 0;
    }

    const sum = scores.reduce((a, b) => a + b, 0);
    let total = sum;
    if (scores.length >= 3) {
      const min = Math.min(...scores);
      const max = Math.max(...scores);
      total = sum - min - max;
    }

    $tr.find(".total").text(total ? total.toFixed(2) : "");
    return total;
  }

  // ======================
  // Read row data object
  // ======================
  function getRowData($tr) {
    const type = $tr.data("type") || "single";
    const parentId = $tr.data("parent") || "";

    const g1 = parseFloat($tr.find(".g1").val()) || 0;
    const g2 = parseFloat($tr.find(".g2").val()) || 0;
    const g3 = parseFloat($tr.find(".g3").val()) || 0;
    const g4 = parseFloat($tr.find(".g4").val()) || 0;
    const g5 = parseFloat($tr.find(".g5").val()) || 0;

    const total = computeTotalFromRow($tr);

    return {
      id: $tr.data("id"),
      type,
      parentId,
      age: ($tr.find(".inp-age").val() || "").trim(),
      event: ($tr.find(".inp-event").val() || "").trim(),
      name: $tr.find(".ath-name").text().trim(),
      yob: ($tr.find(".inp-yob").val() || "").trim(),
      team: ($tr.find(".inp-team").val() || "").trim(),
      g1,
      g2,
      g3,
      g4,
      g5,
      total,
      rank: parseInt($tr.find(".rank-num").text().trim() || "0", 10),
      medal: $tr.find(".medal").text().trim(),
    };
  }

  // ======================
  // Append team group (MASTER + SUB-ROWS)
  // ======================
  function appendTeamGroup({ age, event, team, members }) {
    const $tb = $("#tblScores tbody");
    const masterId = uid();

    $tb.append(
      makeRowHtml({
        id: masterId,
        type: "master",
        age,
        event,
        name: team ? `Đội ${team}` : members[0]?.name || "",
        yob: "",
        team,
        g1: "",
        g2: "",
        g3: "",
        g4: "",
        g5: "",
      })
    );

    (members || []).forEach((m) => {
      $tb.append(
        makeRowHtml({
          type: "member",
          parentId: masterId,
          age,
          event,
          name: m.name,
          yob: m.yob,
          team: m.team || team,
        })
      );
    });
  }
  // ======================
  // Append single row (thi cá nhân / cặp)
  // ======================
  function appendSingleRow({ age, event, name, yob, team }) {
    const $tb = $("#tblScores tbody");
    $tb.append(
      makeRowHtml({
        type: "single",
        age,
        event,
        name,
        yob,
        team,
        g1: "",
        g2: "",
        g3: "",
        g4: "",
        g5: "",
      })
    );
  }

  // ======================
  // Recalc: total + ranking + medal
  // ======================
  function recalcAllScores() {
    const $allRows = $("#tblScores tbody tr");

    if ($allRows.length === 0) {
      $("#tblMedalAthletes tbody").empty();
      $("#tblMedalTeams tbody").empty();
      rebuildFilters();

      return;
    }

    const allowDoubleBronze = $("#chkDoubleBronze").is(":checked");

    // Xóa highlight tie cũ (nếu bạn đã dùng score-tie)
    $allRows.removeClass("score-tie");

    // Chỉ xếp hạng cho single + master
    const $rows = $allRows.filter(function () {
      const type = $(this).data("type") || "single";
      return type !== "member";
    });

    const groups = {};

    // Gom nhóm theo Lứa tuổi + Nội dung
    $rows.each(function () {
      const $tr = $(this);
      const d = getRowData($tr);
      const key = (d.age || "") + "||" + (d.event || "");
      if (!groups[key]) groups[key] = [];
      groups[key].push({ $tr, data: d });
    });

    Object.values(groups).forEach((list) => {
      // Tính tổng điểm cho từng dòng
      list.forEach((item) => {
        item.data.total = computeTotalFromRow(item.$tr);
      });

      // Map tổng điểm -> các dòng cùng tổng (để highlight tie)
      const tieMap = new Map();
      list.forEach((item) => {
        const t = item.data.total;
        if (!t || t <= 0) return;
        const k = t.toFixed(2);
        if (!tieMap.has(k)) tieMap.set(k, []);
        tieMap.get(k).push(item);
      });

      // Sắp xếp để xếp hạng (từ cao xuống)
      list.sort((a, b) => b.data.total - a.data.total);

      // Gán rank + medal (tự động) rồi áp dụng override nếu có
      list.forEach((item, idx) => {
        const d = item.data;
        const $tr = item.$tr;

        const rank = d.total > 0 ? idx + 1 : "";
        let medal = "";

        if (d.total > 0) {
          if (rank === 1) medal = "Vàng";
          else if (rank === 2) medal = "Bạc";
          else if (rank === 3) medal = "Đồng";
        }

        // Đồng huy chương Đồng (nếu bật checkbox)
        if (allowDoubleBronze && rank > 3 && d.total > 0) {
          const prev = list[idx - 1];
          if (prev && prev.data.total === d.total && rank <= 4) {
            medal = "Đồng";
          }
        }

        // Set rank
        $tr.find(".rank-num").text(rank);

        // Nếu dòng này có đặt huy chương tay -> dùng huy chương tay
        const manual = $tr.attr("data-manual-medal");
        if (manual && manual.trim().length > 0) {
          $tr.find(".medal").text(manual.trim());
        } else {
          $tr.find(".medal").text(medal);
        }
      });

      // Highlight các dòng có cùng tổng điểm (nếu bạn muốn)
      tieMap.forEach((itemsWithSameScore) => {
        if (itemsWithSameScore.length >= 2) {
          itemsWithSameScore.forEach((item) => {
            item.$tr.addClass("score-tie");
          });
        }
      });
    });

    // STT cho tất cả (kể cả member)
    let stt = 1;
    $("#tblScores tbody tr:visible").each(function () {
      $(this).find(".stt").text(stt++);
    });

    rebuildMedalTables();
    rebuildFilters();
    applyFilters();
    scheduleAutoSave();
  }

  // ======================
  // BXH huy chương: athletes + teams
  // ======================
  function rebuildMedalTables() {
    const mapAth = new Map();
    const mapTeam = new Map();

    $("#tblScores tbody tr").each(function () {
      const $tr = $(this);
      const type = $tr.data("type") || "single";
      if (type === "member") return;

      const medal = $tr.find(".medal").text().trim();
      if (!medal) return;

      const medalKey =
        medal === "Vàng"
          ? "g"
          : medal === "Bạc"
          ? "s"
          : medal === "Đồng"
          ? "b"
          : null;
      if (!medalKey) return;

      const team = ($tr.find(".inp-team").val() || "").trim();
      if (!team) return;

      if (type === "master") {
        // MASTER: chia huy chương cho tất cả MEMBER
        const id = $tr.data("id");
        $(`#tblScores tbody tr[data-parent="${id}"]`).each(function () {
          const $m = $(this);
          const rawName = $m.find(".ath-name").text();
          const name = rawName.trim();

          if (!name) return;
          const teamName = ($m.find(".inp-team").val() || team).trim();

          const athKey = name + "|" + teamName;
          if (!mapAth.has(athKey)) {
            mapAth.set(athKey, { name, team: teamName, g: 0, s: 0, b: 0 });
          }
          mapAth.get(athKey)[medalKey]++;
        });

        const teamKey = team;
        if (!mapTeam.has(teamKey)) {
          mapTeam.set(teamKey, { team, g: 0, s: 0, b: 0 });
        }
        mapTeam.get(teamKey)[medalKey]++;
      } else {
        // SINGLE: có thể là cá nhân hoặc đội gộp bằng dấu "/"
        const rawName = $tr.find(".ath-name").text().trim();
        if (!rawName) return;

        const names = rawName
          .split("/")
          .map((s) => s.trim())
          .filter(Boolean);
        const listNames = names.length ? names : [rawName];

        listNames.forEach((name) => {
          const athKey = name + "|" + team;
          if (!mapAth.has(athKey)) {
            mapAth.set(athKey, { name, team, g: 0, s: 0, b: 0 });
          }
          mapAth.get(athKey)[medalKey]++;
        });

        const teamKey = team;
        if (!mapTeam.has(teamKey)) {
          mapTeam.set(teamKey, { team, g: 0, s: 0, b: 0 });
        }
        mapTeam.get(teamKey)[medalKey]++;
      }
    });

    const weights = getMedalWeights();

    const athList = Array.from(mapAth.values()).map((a) => {
      const sum = a.g + a.s + a.b;
      const points =
        a.g * weights.gold + a.s * weights.silver + a.b * weights.bronze;
      return { ...a, sum, points };
    });

    const teamList = Array.from(mapTeam.values()).map((t) => {
      const sum = t.g + t.s + t.b;
      const points =
        t.g * weights.gold + t.s * weights.silver + t.b * weights.bronze;
      return { ...t, sum, points };
    });

    const sortAth = (a, b) => {
      if (b.points !== a.points) return b.points - a.points; // ưu tiên điểm
      if (b.g !== a.g) return b.g - a.g;
      if (b.s !== a.s) return b.s - a.s;
      if (b.b !== a.b) return b.b - a.b;
      return a.name.localeCompare(b.name, "vi", { sensitivity: "base" });
    };
    athList.sort(sortAth);

    const sortTeam = (a, b) => {
      if (b.points !== a.points) return b.points - a.points;
      if (b.g !== a.g) return b.g - a.g;
      if (b.s !== a.s) return b.s - a.s;
      if (b.b !== a.b) return b.b - a.b;
      return a.team.localeCompare(b.team, "vi", { sensitivity: "base" });
    };
    teamList.sort(sortTeam);

    const $tbAth = $("#tblMedalAthletes tbody").empty();
    athList.forEach((a, idx) => {
      $tbAth.append(`
        <tr>
          <td class="rank">${idx + 1}</td>
          <td>${a.name}</td>
          <td>${a.team}</td>
          <td class="gold">${a.g}</td>
          <td class="silver">${a.s}</td>
          <td class="bronze">${a.b}</td>
          <td class="sum">${a.sum}</td>
          <td class="points">${a.points}</td>

        </tr>
      `);
    });

    const $tbTeam = $("#tblMedalTeams tbody").empty();
    teamList.forEach((t, idx) => {
      $tbTeam.append(`
        <tr>
          <td class="rank">${idx + 1}</td>
          <td>${t.team}</td>
          <td class="gold">${t.g}</td>
          <td class="silver">${t.s}</td>
          <td class="bronze">${t.b}</td>
          <td class="sum">${t.sum}</td>
          <td class="points">${t.points}</td>
        </tr>
      `);
    });
  }

  // ======================
  // LocalStorage
  // ======================
  function collectState() {
    const rows = $("#tblScores tbody tr")
      .map(function () {
        const d = getRowData($(this));
        return {
          id: d.id,
          type: d.type,
          parentId: d.parentId,
          age: d.age,
          event: d.event,
          name: d.name,
          yob: d.yob,
          team: d.team,
          g1: d.g1,
          g2: d.g2,
          g3: d.g3,
          g4: d.g4,
          g5: d.g5,
        };
      })
      .get();

    return {
      version: 2,
      rows,
      doubleBronze: $("#chkDoubleBronze").is(":checked"),
      medalWeights: getMedalWeights(),
    };
  }

  function applyState(data) {
    if (!data) return;
    const $tb = $("#tblScores tbody").empty();
    (data.rows || []).forEach((r) => {
      $tb.append(makeRowHtml(r));
    });
    $("#chkDoubleBronze").prop("checked", !!data.doubleBronze);

    if (data.medalWeights) {
      setMedalWeightInputs(data.medalWeights);
    } else {
      setMedalWeightInputs(DEFAULT_MEDAL_WEIGHTS);
    }

    recalcAllScores();
  }

  function saveToLocal(showToast = true) {
    try {
      const payload = collectState();
      localStorage.setItem(LS_KEY, JSON.stringify(payload));
      if (showToast) {
        toast("Đã lưu dữ liệu vào máy.");
      }
    } catch (e) {
      console.error(e);
      if (showToast) {
        toast("Không lưu được vào máy (localStorage bị chặn hoặc đầy).", true);
      }
    }
  }
  let autoSaveTimer = null;

  function scheduleAutoSave() {
    // debounce: gõ xong 500ms mới lưu, tránh lưu liên tục
    if (autoSaveTimer) clearTimeout(autoSaveTimer);
    autoSaveTimer = setTimeout(() => {
      saveToLocal(false); // lưu không hiện toast
    }, 500);
  }

  function loadFromLocal() {
    try {
      const raw = localStorage.getItem(LS_KEY);
      if (!raw) return;
      const data = JSON.parse(raw);
      if (!data || !data.rows) return;
      applyState(data);
    } catch (e) {
      console.warn("Load local failed", e);
    }
  }

  // ======================
  // Filters Age / Event
  // ======================
  function rebuildFilters() {
    const ages = new Set();
    const evs = new Set();

    $("#tblScores tbody tr")
      .filter(function () {
        const type = $(this).data("type") || "single";
        return type !== "member"; // chỉ master + single để tránh trùng
      })
      .each(function () {
        const age = ($(this).find(".inp-age").val() || "").trim();
        const ev = ($(this).find(".inp-event").val() || "").trim();
        if (age) ages.add(age);
        if (ev) evs.add(ev);
      });

    const $age = $("#filterAge");
    const curAge = $age.val() || "";
    $age.empty().append('<option value="">Tất cả lứa tuổi</option>');
    Array.from(ages)
      .sort((a, b) => a.localeCompare(b, "vi", { sensitivity: "base" }))
      .forEach((age) => {
        $age.append(`<option value="${age}">${age}</option>`);
      });
    if (curAge && ages.has(curAge)) $age.val(curAge);

    const $ev = $("#filterEvent");
    const curEv = $ev.val() || "";
    $ev.empty().append('<option value="">Tất cả nội dung</option>');
    Array.from(evs)
      .sort((a, b) => a.localeCompare(b, "vi", { sensitivity: "base" }))
      .forEach((ev) => {
        $ev.append(`<option value="${ev}">${ev}</option>`);
      });
    if (curEv && evs.has(curEv)) $ev.val(curEv);
  }

  function applyFilters() {
    const fAge = $("#filterAge").val() || "";
    const fEv = $("#filterEvent").val() || "";
    const q = normalizeStr($("#txtSearch").val() || "");

    $("#tblScores tbody tr").each(function () {
      const $tr = $(this);

      const age = ($tr.find(".inp-age").val() || "").trim();
      const ev = ($tr.find(".inp-event").val() || "").trim();
      const name = $tr.find(".ath-name").text().trim();
      const yob = ($tr.find(".inp-yob").val() || "").trim();
      const team = ($tr.find(".inp-team").val() || "").trim();

      const okAge = !fAge || age === fAge;
      const okEv = !fEv || ev === fEv;

      let okSearch = true;
      if (q) {
        const haystack = normalizeStr([name, yob, age, team].join(" "));
        okSearch = haystack.includes(q);
      }

      const show = okAge && okEv && okSearch;
      $tr.toggle(show);
    });

    // Đánh lại STT cho các dòng đang hiển thị
    let idx = 1;
    $("#tblScores tbody tr:visible").each(function () {
      $(this).find(".stt").text(idx++);
    });
  }

  // ======================
  // Export Excel
  // ======================
  function exportScoresToExcel() {
    if (typeof XLSX === "undefined") {
      toast("Không tìm thấy thư viện XLSX để xuất Excel.", true);
      return;
    }

    // Tính lại điểm, xếp hạng, huy chương cho chắc
    recalcAllScores();

    const header = [
      "Lứa tuổi",
      "Nội dung",
      "Tên VĐV",
      "Năm sinh",
      "Đơn vị",
      "GĐ1",
      "GĐ2",
      "GĐ3",
      "GĐ4",
      "GĐ5",
      "Tổng",
      "Xếp hạng",
      "Thành tích",
    ];

    const aoa = [header];
    const merges = [];
    let currentRow = 1; // 0 là header

    // Regex nhận diện nội dung đồng đội / song luyện (nếu cần dùng sau này)
    const isTeamEventRegex = /(đồng đội|dong doi|song luyện|song luyen)/i;

    const $rows = $("#tblScores tbody tr");

    $rows.each(function () {
      const $tr = $(this);
      const type = $tr.data("type") || "single";

      // Row member của đội sẽ xử lý khi gặp master, nên bỏ qua ở đây
      if (type === "member") return;

      const d = getRowData($tr);

      // Dòng trống hoàn toàn thì bỏ qua
      if (!d.age && !d.event && !d.name && !d.team) return;

      if (type === "single") {
        // ===== Dòng thi cá nhân / cặp: mỗi dòng 1 VĐV/cặp, không merge =====
        aoa.push([
          d.age,
          d.event,
          d.name,
          d.yob,
          d.team,
          d.g1 || "",
          d.g2 || "",
          d.g3 || "",
          d.g4 || "",
          d.g5 || "",
          d.total || "",
          d.rank || "",
          d.medal || "",
        ]);
        currentRow++;
      } else if (type === "master") {
        // ===== Dòng master của đội (đồng đội / song luyện) =====
        const masterId = $tr.data("id");
        const $members = $(`#tblScores tbody tr[data-parent="${masterId}"]`);
        const members = [];

        $members.each(function () {
          const dm = getRowData($(this));
          // hàng member chỉ cần tên + năm sinh
          if (!dm.name) return;
          members.push({
            name: dm.name,
            yob: dm.yob,
          });
        });

        const isMergeGroup = members.length > 1;

        // Nếu không có member (hoặc chỉ 1) thì xuất như 1 dòng bình thường
        if (!isMergeGroup) {
          aoa.push([
            d.age,
            d.event,
            members[0]?.name || d.name, // ưu tiên tên VĐV nếu có
            members[0]?.yob || d.yob,
            d.team,
            d.g1 || "",
            d.g2 || "",
            d.g3 || "",
            d.g4 || "",
            d.g5 || "",
            d.total || "",
            d.rank || "",
            d.medal || "",
          ]);
          currentRow++;
        } else {
          // Có từ 2 VĐV trở lên -> xuất nhiều dòng + merge các cột chung
          const startRow = currentRow;

          members.forEach((m, idx) => {
            if (idx === 0) {
              // Dòng đầu của đội: ghi đủ thông tin + điểm
              aoa.push([
                d.age,
                d.event,
                m.name,
                m.yob,
                d.team,
                d.g1 || "",
                d.g2 || "",
                d.g3 || "",
                d.g4 || "",
                d.g5 || "",
                d.total || "",
                d.rank || "",
                d.medal || "",
              ]);
            } else {
              // Các dòng sau: chỉ ghi tên + năm sinh, còn lại để trống và sẽ merge
              aoa.push([
                "", // Lứa tuổi (merge)
                "", // Nội dung (merge)
                m.name,
                m.yob,
                "", // Đơn vị (merge)
                "", // GĐ1 (merge)
                "", // GĐ2 (merge)
                "", // GĐ3 (merge)
                "", // GĐ4 (merge)
                "", // GĐ5 (merge)
                "", // Tổng (merge)
                "", // Xếp hạng (merge)
                "", // Thành tích (merge)
              ]);
            }
            currentRow++;
          });

          const endRow = currentRow - 1;

          // Merge dọc các cột chung:
          // Lứa tuổi (0), Nội dung (1), Đơn vị (4),
          // GĐ1..GĐ5 (5..9), Tổng (10), Xếp hạng (11), Thành tích (12)
          const colsToMerge = [0, 1, 4, 5, 6, 7, 8, 9, 10, 11, 12];

          colsToMerge.forEach((c) => {
            merges.push({
              s: { r: startRow, c },
              e: { r: endRow, c },
            });
          });
        }
      }
    });

    // ===== 3. Tạo sheet từ AOA =====
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Merge
    if (merges.length) {
      ws["!merges"] = merges;
    }

    // Độ rộng cột
    ws["!cols"] = [
      { wch: 10 }, // Lứa tuổi
      { wch: 26 }, // Nội dung
      { wch: 22 }, // Tên VĐV
      { wch: 10 }, // Năm sinh
      { wch: 18 }, // Đơn vị
      { wch: 6 },  // GĐ1
      { wch: 6 },  // GĐ2
      { wch: 6 },  // GĐ3
      { wch: 6 },  // GĐ4
      { wch: 6 },  // GĐ5
      { wch: 8 },  // Tổng
      { wch: 9 },  // Xếp hạng
      { wch: 12 }, // Thành tích
    ];

    // AutoFilter
    const lastRow = aoa.length - 1;
    const lastCol = header.length - 1;
    ws["!autofilter"] = {
      ref: XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: lastRow, c: lastCol },
      }),
    };

    // Freeze header
    ws["!freeze"] = {
      xSplit: 0,
      ySplit: 1,
      topLeftCell: "A2",
      activePane: "bottomLeft",
      state: "frozen",
    };

    // Style header: đậm + căn giữa (nếu bản XLSX bạn dùng hỗ trợ style)
    const ref = ws["!ref"];
    if (ref) {
      const hdrRange = XLSX.utils.decode_range(ref);
      for (let C = hdrRange.s.c; C <= hdrRange.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ r: 0, c: C });
        if (!ws[addr]) continue;
        ws[addr].s = {
          font: { bold: true },
          alignment: { horizontal: "center", vertical: "center" },
        };
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Scores");
    XLSX.writeFile(wb, "vovinam_scores.xlsx");
  }


  // ======================
  // Import Excel dạng đăng ký 5 cột
  // (Lứa tuổi, Nội dung, Tên VĐV, Năm sinh, Đơn vị)
  // ======================
  function importRegistrationFormat(rows) {
    const $tb = $("#tblScores tbody");

    // =========================
    // 1. Đọc dữ liệu hiện có để chống trùng
    // =========================
    const existingSingles = new Set();
    const existingTeams = new Set();

    $tb.find("tr").each(function () {
      const $tr = $(this);
      const d = getRowData($tr);
      const type = $tr.data("type") || "single";

      if (type === "single") {
        const key =
          "S|" +
          normalizeStr(d.age) +
          "|" +
          normalizeStr(d.event) +
          "|" +
          normalizeStr(d.name) +
          "|" +
          (d.yob || "") +
          "|" +
          normalizeStr(d.team);
        existingSingles.add(key);
      } else if (type === "master") {
        // Master row của team (Song luyện / đồng đội)
        const masterId = $tr.data("id");
        const members = [];
        $tb.find(`tr[data-parent="${masterId}"]`).each(function () {
          const dm = getRowData($(this));

          // tên row con
          const memberName = dm.name

          members.push(normalizeStr(memberName) + "#" + (dm.yob || ""));
        });
        members.sort();
        const teamKey =
          "T|" +
          normalizeStr(d.age) +
          "|" +
          normalizeStr(d.event) +
          "|" +
          normalizeStr(d.team) +
          "|" +
          members.join("|");
        existingTeams.add(teamKey);
      }
    });

    // =========================
    // 2. Parse file Excel 5 cột -> regRows
    // =========================
    const regRows = [];
    let curAge = "";
    let curEvent = "";

    rows.forEach((rowRaw, idx) => {
      const row =
        rowRaw && rowRaw.map
          ? rowRaw.map((c) =>
              c === undefined || c === null ? "" : String(c).trim()
            )
          : [];

      if (!row.length) return;
      while (row.length < 5) row.push("");

      const colAge = row[0] || "";
      const colEvent = row[1] || "";
      const colName = row[2] || "";
      const colYob = row[3] || "";
      const colTeam = row[4] || "";

      // bỏ header
      if (
        idx === 0 &&
        /lứa tuổi/i.test(colAge || "") &&
        /nội dung/i.test(colEvent || "")
      ) {
        return;
      }

      // bỏ dòng rỗng
      if (!colAge && !colEvent && !colName && !colYob && !colTeam) return;

      if (colAge) curAge = colAge;
      if (colEvent) curEvent = colEvent;
      if (!colName) return;

      regRows.push({
        age: curAge,
        event: curEvent,
        name: colName,
        yob: colYob,
        rawTeam: colTeam, // Đơn vị đúng như trong file
      });
    });

    // =========================
    // 3. Duyệt regRows -> thêm single / team
    // =========================
    let addedSingles = 0;
    let skippedSingles = 0;
    let addedTeams = 0;
    let skippedTeams = 0;

    let i = 0;
    while (i < regRows.length) {
      const r = regRows[i];

      const evLower = (r.event || "").toLowerCase();
      const isTeamEvent =
        evLower.includes("song luyện") ||
        evLower.includes("song luyen") ||
        evLower.includes("đồng đội") ||
        evLower.includes("dong doi");

      // -------- CÁ NHÂN / KHÔNG TEAM --------
      if (!isTeamEvent || !r.rawTeam) {
        const key =
          "S|" +
          normalizeStr(r.age) +
          "|" +
          normalizeStr(r.event) +
          "|" +
          normalizeStr(r.name) +
          "|" +
          (r.yob || "") +
          "|" +
          normalizeStr(r.rawTeam || "");

        if (existingSingles.has(key)) {
          skippedSingles++;
        } else {
          existingSingles.add(key);
          appendSingleRow({
            age: r.age,
            event: r.event,
            name: r.name,
            yob: r.yob,
            team: r.rawTeam || "",
          });
          addedSingles++;
        }

        i++;
        continue;
      }

      // -------- TEAM: SONG LUYỆN / ĐỒNG ĐỘI --------
      const teamName = r.rawTeam; // dòng đầu tiên luôn có Đơn vị
      const group = {
        age: r.age,
        event: r.event,
        team: teamName,
        members: [
          {
            name: r.name,
            yob: r.yob,
            team: teamName,
          },
        ],
      };
      i++;

      // Gom tiếp các dòng cùng event, không có rawTeam (ô Đơn vị trống)
      while (i < regRows.length) {
        const r2 = regRows[i];
        const evLower2 = (r2.event || "").toLowerCase();

        // khác nội dung hoặc có rawTeam mới -> dừng block hiện tại
        if (
          evLower2 !== evLower ||
          r2.rawTeam // dòng này sẽ là đội kế tiếp
        ) {
          break;
        }

        group.members.push({
          name: r2.name,
          yob: r2.yob,
          team: teamName,
        });
        i++;
      }

      // Tạo key đội để chống trùng
      const membersKey = group.members
        .map((m) => normalizeStr(m.name) + "#" + (m.yob || ""))
        .sort()
        .join("|");

      const teamKey =
        "T|" +
        normalizeStr(group.age) +
        "|" +
        normalizeStr(group.event) +
        "|" +
        normalizeStr(group.team || "") +
        "|" +
        membersKey;

      if (existingTeams.has(teamKey)) {
        skippedTeams++;
      } else {
        existingTeams.add(teamKey);
        appendTeamGroup(group); // MASTER + SUB-ROWS
        addedTeams++;
      }
    }

    // =========================
    // 4. Kết quả
    // =========================
    if (addedSingles > 0 || addedTeams > 0) {
      recalcAllScores();
      const parts = [];
      if (addedSingles) parts.push(`${addedSingles} VĐV/cặp`);
      if (addedTeams) parts.push(`${addedTeams} đội`);
      let msg = "Đã thêm " + parts.join(" + ") + " từ file đăng ký.";
      if (skippedSingles || skippedTeams) {
        msg += ` (Bỏ qua ${skippedSingles} dòng cá nhân + ${skippedTeams} đội trùng.)`;
      }
      toast(msg);
    } else if (skippedSingles || skippedTeams) {
      toast("Tất cả dữ liệu trong file đăng ký đã có trong bảng.", true);
    } else {
      toast("Không tìm thấy dòng hợp lệ trong file đăng ký.", true);
    }
  }

  function importScore14ColsFormat(rows) {
    const $tb = $("#tblScores tbody");

    const existingKeys = new Set();
    $tb.find("tr").each(function () {
      const d = getRowData($(this));
      const key =
        "single|" +
        normalizeStr(d.age) +
        "|" +
        normalizeStr(d.event) +
        "|" +
        normalizeStr(d.name) +
        "|" +
        (d.yob || "") +
        "|" +
        normalizeStr(d.team);
      existingKeys.add(key);
    });

    let added = 0;
    let skipped = 0;

    const parseScore = (v) => {
      if (v === undefined || v === null || v === "") return "";
      const num = parseFloat(String(v).replace(",", "."));
      return isNaN(num) ? "" : num;
    };

    rows.forEach((rowRaw) => {
      const row =
        rowRaw && rowRaw.map
          ? rowRaw.map((c) =>
              c === undefined || c === null ? "" : String(c).trim()
            )
          : [];

      if (!row.length || row.every((c) => c === "")) return;

      const joined = row.join(" ").toLowerCase();
      if (
        joined.includes("stt") &&
        (joined.includes("vđv") ||
          joined.includes("ten vdv") ||
          joined.includes("tên vđv"))
      ) {
        return;
      }

      while (row.length < 14) row.push("");

      const age = row[1] || "";
      const event = row[2] || "";
      const name = row[3] || "";
      const yob = row[4] || "";
      const team = row[5] || "";
      if (!name) return;

      const g1 = parseScore(row[6]);
      const g2 = parseScore(row[7]);
      const g3 = parseScore(row[8]);
      const g4 = parseScore(row[9]);
      const g5 = parseScore(row[10]);

      const key =
        "single|" +
        normalizeStr(age) +
        "|" +
        normalizeStr(event) +
        "|" +
        normalizeStr(name) +
        "|" +
        (yob || "") +
        "|" +
        normalizeStr(team);

      if (existingKeys.has(key)) {
        skipped++;
        return;
      }
      existingKeys.add(key);
      added++;

      $tb.append(
        makeRowHtml({
          type: "single",
          age,
          event,
          name,
          yob,
          team,
          g1: g1 === "" ? "" : g1,
          g2: g2 === "" ? "" : g2,
          g3: g3 === "" ? "" : g3,
          g4: g4 === "" ? "" : g4,
          g5: g5 === "" ? "" : g5,
        })
      );
    });

    if (added > 0) {
      recalcAllScores();
      const msg = skipped
        ? `Đã thêm ${added} dòng mới từ file 14 cột, bỏ qua ${skipped} dòng trùng.`
        : `Đã thêm ${added} dòng mới từ file 14 cột.`;
      toast(msg);
    } else if (skipped > 0) {
      toast("Tất cả các dòng trong file 14 cột đã có trong bảng.", true);
    } else {
      toast("Không tìm thấy dòng hợp lệ trong file 14 cột.", true);
    }
  }

  function importScoresFrom2D(rows) {
    if (!rows || !rows.length) {
      toast("File không có dữ liệu.", true);
      return;
    }

    let first = null;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i] || [];
      if (
        r.some((c) => c !== null && c !== undefined && String(c).trim() !== "")
      ) {
        first = { idx: i, row: r };
        break;
      }
    }
    if (!first) {
      toast("File toàn dòng trống.", true);
      return;
    }

    const r0 = first.row.map((c) =>
      c === undefined || c === null ? "" : String(c).trim()
    );

    const isReg =
      /lứa tuổi/i.test(r0[0] || "") && /nội dung/i.test(r0[1] || "");

    if (isReg) {
      importRegistrationFormat(rows);
    } else {
      importScore14ColsFormat(rows);
    }
  }

  function handleImportScoresFile(file) {
    if (!file) return;

    const ext = (file.name.split(".").pop() || "").toLowerCase();

    if (ext === "csv") {
      const reader = new FileReader();
      reader.onload = function (ev) {
        const text = ev.target.result || "";
        const lines = text.split(/\r?\n/);
        const rows = lines
          .map((line) => line.split(/,|;|\t/))
          .filter((r) => r.some((cell) => (cell || "").trim() !== ""));
        importScoresFrom2D(rows);
      };
      reader.readAsText(file, "utf-8");
      return;
    }

    if (typeof XLSX === "undefined") {
      toast("Không tìm thấy thư viện XLSX (excel.min.js).", true);
      return;
    }

    const reader = new FileReader();
    reader.onload = function (ev) {
      const data = new Uint8Array(ev.target.result);
      const wb = XLSX.read(data, { type: "array" });
      if (!wb.SheetNames || !wb.SheetNames.length) {
        toast("File Excel không có sheet nào.", true);
        return;
      }
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      importScoresFrom2D(rows);
    };
    reader.readAsArrayBuffer(file);
  }

  // ======================
  // Events
  // ======================
  $(document).on(
    "input change",
    ".inp-age, .inp-event, .inp-yob, .inp-team, .inp-g",
    function () {
      recalcAllScores();
    }
  );

  $(document).on("input", ".ath-name", function () {
    recalcAllScores();
  });

  $(document).on("click", ".btnDelRow", function () {
    const $tr = $(this).closest("tr");
    const type = $tr.data("type") || "single";
    if (type === "master") {
      const id = $tr.data("id");
      $(`#tblScores tbody tr[data-parent="${id}"]`).remove();
    }
    $tr.remove();
    recalcAllScores();
  });

  $("#btnAddRow").on("click", function () {
    $("#tblScores tbody").append(
      makeRowHtml({
        type: "single",
        age: "",
        event: "",
        name: "",
        yob: "",
        team: "",
        g1: "",
        g2: "",
        g3: "",
        g4: "",
        g5: "",
      })
    );
    recalcAllScores();
  });
  // Đặt huy chương bằng tay bằng cách double-click vào ô huy chương
  $(document).on("dblclick", "#tblScores tbody .medal", function () {
    const $td = $(this);
    const $tr = $td.closest("tr");
    const type = $tr.data("type") || "single";

    // Không áp dụng cho dòng member của đội
    if (type === "member") return;

    const currentManual = $tr.attr("data-manual-medal") || "";
    const currentDisplay = $td.text().trim();
    const current = currentManual || currentDisplay;

    const input = prompt(
      "Nhập huy chương cho dòng này (Vàng/Bạc/Đồng) hoặc để trống để bỏ đặt tay:",
      current
    );

    if (input === null) return; // bấm Cancel

    let val = input.trim().toLowerCase();

    if (!val) {
      // Xoá override -> quay về auto
      $tr.removeAttr("data-manual-medal");
    } else {
      if (val === "vang" || val === "vàng" || val === "v") val = "Vàng";
      else if (val === "bac" || val === "bạc" || val === "b") val = "Bạc";
      else if (val === "dong" || val === "đồng" || val === "d") val = "Đồng";
      else {
        alert(
          "Giá trị không hợp lệ. Vui lòng nhập Vàng, Bạc hoặc Đồng (hoặc để trống)."
        );
        return;
      }
      $tr.attr("data-manual-medal", val);
    }

    // Tính lại để đảm bảo bảng, thống kê, medal table… sync với huy chương tay
    recalcAllScores();
  });

  $("#chkDoubleBronze").on("change", function () {
    recalcAllScores();
  });
  $("#coefGold, #coefSilver, #coefBronze").on("input change", function () {
    rebuildMedalTables();
    //  tự lưu luôn:
    saveToLocal(false);
  });
  $("#filterAge, #filterEvent").on("change", function () {
    applyFilters();
  });
  $("#txtSearch").on("input", function () {
    applyFilters();
  });

  $("#btnSave").on("click", function () {
    saveToLocal(true);
  });

  $("#btnReset").on("click", function () {
    if (
      confirm(
        "Xoá toàn bộ dữ liệu trên bảng và trong máy (localStorage)? Thao tác này không thể hoàn tác."
      )
    ) {
      $("#tblScores tbody").empty();
      $("#tblMedalAthletes tbody").empty();
      $("#tblMedalTeams tbody").empty();
      try {
        localStorage.removeItem(LS_KEY);
      } catch (e) {
        console.warn(e);
      }
      rebuildFilters();
    }
  });

  $("#btnExportScores").on("click", function () {
    exportScoresToExcel();
  });

  $("#btnImportScores").on("click", function () {
    $("#fileImportScores").val("");
    $("#fileImportScores").trigger("click");
  });

  $("#fileImportScores").on("change", function (e) {
    const file = e.target.files[0];
    if (file) handleImportScoresFile(file);
    $(this).val("");
  });

  // ======================
  // Init
  // ======================
  $(function () {
    loadFromLocal();
    recalcAllScores();
  });
})(jQuery);
