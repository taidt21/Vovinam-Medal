(function ($) {
  "use strict";

  const LS_KEY = "vovinam.scores.v2";

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
    .replace(/đ/g, "d");             // chuyển đ -> d để bắt được "dong doi"
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
    const nameCell = isMember
      ? `<td contenteditable="true" class="ath-name">– ${name || ""}</td>`
      : `<td contenteditable="true" class="ath-name">${name || ""}</td>`;

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
        <td><input class="inp-g g2" type="number" step="0.1" value "${g2}" /></td>
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

    // Rows được xếp hạng: single + master
    const $rows = $allRows.filter(function () {
      const type = $(this).data("type") || "single";
      return type !== "member";
    });

    const groups = {};

    $rows.each(function () {
      const $tr = $(this);
      const d = getRowData($tr);
      const key = (d.age || "") + "||" + (d.event || "");
      if (!groups[key]) groups[key] = [];
      groups[key].push({ $tr, data: d });
    });

    Object.values(groups).forEach((list) => {
      list.forEach((item) => {
        item.data.total = computeTotalFromRow(item.$tr);
      });

      list.sort((a, b) => b.data.total - a.data.total);

      list.forEach((item, idx) => {
        const rank = idx + 1;
        const d = item.data;
        const $tr = item.$tr;

        let medal = "";
        if (d.total > 0) {
          if (rank === 1) medal = "Vàng";
          else if (rank === 2) medal = "Bạc";
          else if (rank === 3) medal = "Đồng";
        }

        if (allowDoubleBronze && rank > 3 && d.total > 0) {
          const prev = list[idx - 1];
          if (prev && prev.data.total === d.total && rank <= 4) {
            medal = "Đồng";
          }
        }

        $tr.find(".rank-num").text(d.total > 0 ? rank : "");
        $tr.find(".medal").text(medal);
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
          const name = $m.find(".ath-name").text().replace(/^–\s*/, "").trim();
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

    const athList = Array.from(mapAth.values()).map((a) => ({
      ...a,
      sum: a.g + a.s + a.b,
    }));
    const teamList = Array.from(mapTeam.values()).map((t) => ({
      ...t,
      sum: t.g + t.s + t.b,
    }));

    const sortAth = (a, b) => {
      if (b.g !== a.g) return b.g - a.g;
      if (b.s !== a.s) return b.s - a.s;
      if (b.b !== a.b) return b.b - a.b;
      return a.name.localeCompare(b.name, "vi", { sensitivity: "base" });
    };
    athList.sort(sortAth);

    const sortTeam = (a, b) => {
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
          <td>${a.g}</td>
          <td>${a.s}</td>
          <td>${a.b}</td>
          <td>${a.sum}</td>
        </tr>
      `);
    });

    const $tbTeam = $("#tblMedalTeams tbody").empty();
    teamList.forEach((t, idx) => {
      $tbTeam.append(`
        <tr>
          <td class="rank">${idx + 1}</td>
          <td>${t.team}</td>
          <td>${t.g}</td>
          <td>${t.s}</td>
          <td>${t.b}</td>
          <td>${t.sum}</td>
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
    };
  }

  function applyState(data) {
    if (!data) return;
    const $tb = $("#tblScores tbody").empty();
    (data.rows || []).forEach((r) => {
      $tb.append(makeRowHtml(r));
    });
    $("#chkDoubleBronze").prop("checked", !!data.doubleBronze);
    recalcAllScores();
  }

  function saveToLocal() {
    try {
      const payload = collectState();
      localStorage.setItem(LS_KEY, JSON.stringify(payload));
      toast("Đã lưu dữ liệu vào máy.");
    } catch (e) {
      console.error(e);
      toast("Không lưu được vào máy (localStorage bị chặn hoặc đầy).", true);
    }
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

    $("#tblScores tbody tr").each(function () {
      const age = ($(this).find(".inp-age").val() || "").trim();
      const ev = ($(this).find(".inp-event").val() || "").trim();
      const okAge = !fAge || age === fAge;
      const okEv = !fEv || ev === fEv;
      $(this).toggle(okAge && okEv);
    });

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
      toast("Không tìm thấy thư viện XLSX (excel.min.js).", true);
      return;
    }

    const wb = XLSX.utils.book_new();
    const rows = [];

    rows.push([
      "STT",
      "Lứa tuổi",
      "Nội dung",
      "VĐV",
      "Năm sinh",
      "Đơn vị",
      "GĐ1",
      "GĐ2",
      "GĐ3",
      "GĐ4",
      "GĐ5",
      "Tổng điểm",
      "Xếp hạng",
      "Thành tích",
      "Loại dòng",
    ]);

    $("#tblScores tbody tr").each(function (idx) {
      const $tr = $(this);
      const d = getRowData($tr);
      const rankText = $tr.find(".rank-num").text().trim();
      const medal = $tr.find(".medal").text().trim();
      const type = $tr.data("type") || "single";

      rows.push([
        idx + 1,
        d.age,
        d.event,
        d.name,
        d.yob,
        d.team,
        d.g1,
        d.g2,
        d.g3,
        d.g4,
        d.g5,
        d.total,
        rankText,
        medal,
        type,
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, "Scores");
    XLSX.writeFile(wb, "vovinam_scores.xlsx");
  }

  // ======================
  // Import Excel dạng đăng ký 5 cột
  // (Lứa tuổi, Nội dung, Tên VĐV, Năm sinh, Đơn vị)
  // ======================
  function importRegistrationFormat(rows) {
    const $tb = $("#tblScores tbody");

    // 1) Đọc dữ liệu hiện có để tránh trùng VĐV cá nhân
    const existingSingles = new Set();
    $tb.find("tr").each(function () {
      const d = getRowData($(this));
      if ((d.type || "single") !== "single") return;
      const key =
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
    });

    // 2) Parse file Excel thành list dòng đăng ký
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

      // Bỏ dòng header
      if (
        idx === 0 &&
        /lứa tuổi/i.test(colAge || "") &&
        /nội dung/i.test(colEvent || "")
      ) {
        return;
      }
      // Bỏ dòng trống
      if (!colAge && !colEvent && !colName && !colYob && !colTeam) return;

      if (colAge) curAge = colAge;
      if (colEvent) curEvent = colEvent;
      if (!colName) return;

      regRows.push({
        age: curAge,
        event: curEvent,
        name: colName,
        yob: colYob,
        rawTeam: colTeam, // team đúng như trong file (có thể trống)
      });
    });

    // 3) Duyệt list và build bảng: singles + team (Song luyện / Đồng đội)
    let addedSingles = 0;
    let skippedSingles = 0;
    let addedTeams = 0;

    let i = 0;
    while (i < regRows.length) {
      const r = regRows[i];
      const normEvent = normalizeStr(r.event);
      const isTeamEvent =
        normEvent.includes("song luyen") || normEvent.includes("dong doi");

      // ===== Nội dung CÁ NHÂN / hoặc không có team =====
      if (!isTeamEvent || !r.rawTeam) {
        const key =
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

      // ===== Nội dung ĐỒNG ĐỘI / SONG LUYỆN =====
      // r.rawTeam chắc chắn có giá trị -> bắt đầu 1 đội mới
      const teamName = r.rawTeam;
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

      // Gom các thành viên tiếp theo cùng event, không có rawTeam (ô Đơn vị trống)
      while (i < regRows.length) {
        const r2 = regRows[i];
        const normEvent2 = normalizeStr(r2.event);

        // Khác nội dung hoặc xuất hiện rawTeam mới -> dừng block hiện tại
        if (normEvent2 !== normEvent || r2.rawTeam) break;

        group.members.push({
          name: r2.name,
          yob: r2.yob,
          team: teamName,
        });
        i++;
      }

      appendTeamGroup(group);
      addedTeams++;
    }

    // 4) Hoàn tất
    if (addedSingles > 0 || addedTeams > 0) {
      recalcAllScores();
      const parts = [];
      if (addedSingles) parts.push(`${addedSingles} VĐV/cặp`);
      if (addedTeams) parts.push(`${addedTeams} đội`);
      let msg = "Đã thêm " + parts.join(" + ") + " từ file đăng ký.";
      if (skippedSingles) msg += ` Bỏ qua ${skippedSingles} dòng cá nhân trùng.`;
      toast(msg);
    } else if (skippedSingles > 0) {
      toast("Tất cả các dòng cá nhân trong file đã có trong bảng.", true);
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

  $("#chkDoubleBronze").on("change", function () {
    recalcAllScores();
  });

  $("#filterAge, #filterEvent").on("change", function () {
    applyFilters();
  });

  $("#btnSave").on("click", function () {
    saveToLocal();
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
