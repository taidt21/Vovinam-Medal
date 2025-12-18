/** @format */

(function ($) {
  "use strict";

  const LS_KEY = "vovinam.scores.v2";

  const DEFAULT_MEDAL_WEIGHTS = {
    gold: 10,
    silver: 6,
    bronze: 3,
  };

  // Không giới hạn điểm tối đa (tuỳ thang điểm của giải)
  // Chỉ chặn số âm để giảm rủi ro nhập nhầm.
  const SCORE_MIN = 0;

  // Làm tròn 2 chữ số để so sánh tie ổn định (tránh lỗi float)
  function round2(n) {
    const x = parseFloat(n);
    if (isNaN(x)) return "";
    return (Math.round((x + Number.EPSILON) * 100) / 100).toFixed(2);
  }

  // debounce recalc để nhập nhanh hơn ở bàn tổng trọng tài
  let recalcTimer = null;
  function requestRecalc() {
    if (recalcTimer) clearTimeout(recalcTimer);
    recalcTimer = setTimeout(() => {
      recalcAllScores();
    }, 80);
  }

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

  function normalizeMedal(input) {
    const raw = (input || "").toString().trim();
    if (!raw) return "";
    const v = normalizeStr(raw);
    if (v === "vang" || v === "v") return "Vàng";
    if (v === "bac" || v === "b") return "Bạc";
    if (v === "dong" || v === "d" || v === "đ" || v === "dd") return "Đồng";
    // chấp nhận nhập đúng tiếng Việt
    if (raw === "Vàng" || raw === "Bạc" || raw === "Đồng") return raw;
    return "";
  }

  function isCombatEvent(ev) {
    const v = normalizeStr(ev);
    // bắt: "đối kháng" / "doi khang" / "doi-khang" / "sparring" / "combat"
    return (
      v.includes("doi khang") ||
      v.includes("doi-khang") ||
      v.includes("doikhang") ||
      v.includes("sparring") ||
      v.includes("combat")
    );
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
    const stt = row.stt ?? "";
    const medal = row.manualMedal || row.medal || "";

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
    const nameCell = `<td contenteditable="true" class="ath-name ${
      isMember ? " is-member" : ""
    }">
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
        <td></td>
      `
      : `
        <td><input class="inp-g g1" type="number" step="0.1" value="${g1}" /></td>
        <td><input class="inp-g g2" type="number" step="0.1" value="${g2}" /></td>
        <td><input class="inp-g g3" type="number" step="0.1" value="${g3}" /></td>
        <td><input class="inp-g g4" type="number" step="0.1" value="${g4}" /></td>
        <td><input class="inp-g g5" type="number" step="0.1" value="${g5}" /></td>
        <td class="total">${total}</td>
        <td class="rank-num"><input class="inp-rank" type="number" inputmode="numeric" value="${rank}" placeholder="#" /></td>
        <td><input class="inp-medal" type="text" value="${medal}" placeholder="V/B/Đ" /></td>
      `;

    return `
      <tr data-id="${id}" data-type="${type}" data-stt="${stt}" ${
      parentId ? `data-parent="${parentId}"` : ""
    } ${!isMember && medal ? `data-manual-medal="${medal}"` : ""} class="${rowClass}">
        <td class="rank stt">${stt || ""}</td>
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
      const $inp = $tr.find(cls);
      if ($inp.is(":disabled")) return null;
      const v = $inp.val();
      if (v === undefined || v === null || v === "") return null;
      const num = parseFloat(v.toString().replace(",", "."));
      if (isNaN(num)) return null;
      if (num < SCORE_MIN) return null;
      return num;
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
      stt: parseInt(($tr.attr("data-stt") || "").toString().trim() || "0", 10),
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
      rank: parseInt(($tr.find(".inp-rank").val() || "").toString().trim() || "0", 10),
      medal: (($tr.find(".inp-medal").val() || "") + "").trim(),
      manualMedal: (($tr.attr("data-manual-medal") || "") + "").trim(),
    };
  }

  
  // ======================
  // STT: giữ nguyên theo thứ tự tạo (không đổi khi sort/filter)
  // ======================
  function getNextStt() {
    let max = 0;
    $("#tblScores tbody tr").each(function () {
      const v = parseInt(($(this).attr("data-stt") || "").toString().trim() || "0", 10);
      if (v > max) max = v;
    });
    return max + 1;
  }

// ======================
  // Append team group (MASTER + SUB-ROWS)
  // ======================
  function appendTeamGroup({ age, event, team, members }) {
    const $tb = $("#tblScores tbody");
    const masterId = uid();

    let stt = getNextStt();

    $tb.append(
      makeRowHtml({
        id: masterId,
        type: "master",
        stt: stt++,
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
          stt: stt++,
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
        stt: getNextStt(),
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

    // Reset highlight cũ
    $allRows.removeClass("score-tie is-combat");

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
      const groupEvent = (list[0] && list[0].data ? list[0].data.event : "") || "";
      const isCombatGroup = isCombatEvent(groupEvent);

      // Đối kháng: khoá điểm (disabled), chỉ nhập huy chương
      // Biểu diễn: điểm editable bình thường
      list.forEach((item) => {
        const $tr = item.$tr;
        const $scores = $tr.find(".inp-g");

        if (isCombatGroup) {
          $tr.addClass("is-combat");
          $scores.prop("disabled", true).attr("placeholder", "—");
          $tr.find(".total").text("");
        } else {
          $scores.prop("disabled", false).removeAttr("placeholder");
        }
      });

      // ====== ĐỐI KHÁNG: khoá điểm (disabled), chỉ nhập huy chương ======
      if (isCombatGroup) {
        list.forEach((item) => {
          const $tr = item.$tr;
          const $inp = $tr.find(".inp-medal");
          const raw = (($inp.val() || "") + "").trim();
          const norm = normalizeMedal(raw);

          // Lưu manual medal cho lần save/load
          if (!raw) {
            $tr.removeAttr("data-manual-medal");
          } else if (norm) {
            $tr.attr("data-manual-medal", norm);
            if (!$inp.is(":focus")) $inp.val(norm);
          }

          $tr.attr("data-auto-medal", "");
          $tr.attr("data-final-medal", norm || "");
          // Xếp hạng: KHÔNG tự đổi ở đây (chỉ đổi khi bấm Sort)
        });

        return; // không highlight tie ở đối kháng
      }

      // ====== BIỂU DIỄN: xếp hạng theo tổng điểm ======
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

      list.forEach((item, idx) => {
        const d = item.data;
        const $tr = item.$tr;

        const rank = d.total > 0 ? idx + 1 : "";
        let autoMedal = "";

        if (d.total > 0) {
          if (rank === 1) autoMedal = "Vàng";
          else if (rank === 2) autoMedal = "Bạc";
          else if (rank === 3) autoMedal = "Đồng";
        }

        // Đồng huy chương Đồng (nếu bật checkbox): luôn trao thêm 1 huy chương Đồng cho hạng 4
        if (allowDoubleBronze && rank === 4 && d.total > 0) {
          autoMedal = "Đồng";
        }
        $tr.attr("data-auto-medal", autoMedal);

        const manual = normalizeMedal(($tr.attr("data-manual-medal") || "").trim());
        const finalMedal = manual || autoMedal;
        $tr.attr("data-final-medal", finalMedal);

        const $inp = $tr.find(".inp-medal");
        if ($inp.length && !$inp.is(":focus")) {
          $inp.val(finalMedal || "");
        }
      });

      // Highlight các dòng có cùng tổng điểm
      tieMap.forEach((itemsWithSameScore) => {
        if (itemsWithSameScore.length >= 2) {
          itemsWithSameScore.forEach((item) => {
            item.$tr.addClass("score-tie");
          });
        }
      });
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

      const medal = normalizeMedal(
        ((
          $tr.attr("data-final-medal") ||
          $tr.find(".inp-medal").val() ||
          ""
        ) + "").trim()
      );
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
          stt: d.stt || 0,
          rank: d.rank || "",
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
          manualMedal: d.manualMedal,
        };
      })
      .get();

    return {
      version: 3,
      rows,
      doubleBronze: $("#chkDoubleBronze").is(":checked"),
      medalWeights: getMedalWeights(),
    };
  }

  function applyState(data) {
    if (!data) return;
    const $tb = $("#tblScores tbody").empty();

    // Backward compatibility: nếu data cũ chưa có STT, tự gán theo thứ tự load
    let fallbackStt = getNextStt();

    (data.rows || []).forEach((r) => {
      if (!r.stt || parseInt(r.stt, 10) <= 0) {
        r.stt = fallbackStt++;
      }
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

    const $rows = $("#tblScores tbody tr:visible");

    $rows.each(function () {
      const $tr = $(this);
      const type = $tr.data("type") || "single";

      // Row member của đội sẽ xử lý khi gặp master, nên bỏ qua ở đây
      if (type === "member") return;

      const d = getRowData($tr);

      // Dòng trống hoàn toàn thì bỏ qua
      if (!d.age && !d.event && !d.name && !d.team) return;

      if (type === "single") {
        const finalMedal =
          normalizeMedal(($tr.attr("data-final-medal") || "").trim()) ||
          normalizeMedal(d.medal);
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
          finalMedal || "",
        ]);
        currentRow++;
      } else if (type === "master") {
        const finalMedal =
          normalizeMedal(($tr.attr("data-final-medal") || "").trim()) ||
          normalizeMedal(d.medal);
        // ===== Dòng master của đội (đồng đội / song luyện) =====
        const masterId = $tr.data("id");
        const $members = $(`#tblScores tbody tr[data-parent="${masterId}"]:visible`);
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
            finalMedal || "",
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
                finalMedal || "",
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
      { wch: 6 }, // GĐ1
      { wch: 6 }, // GĐ2
      { wch: 6 }, // GĐ3
      { wch: 6 }, // GĐ4
      { wch: 6 }, // GĐ5
      { wch: 8 }, // Tổng
      { wch: 9 }, // Xếp hạng
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
  // Export Excel: BXH huy chương (VĐV / Đoàn)
  // ======================
  function exportHtmlTableToExcel(tableSelector, fileName, sheetName) {
    if (typeof XLSX === "undefined") {
      toast("Không tìm thấy thư viện XLSX để xuất Excel.", true);
      return;
    }

    const $table = $(tableSelector);
    if ($table.length === 0) {
      toast("Không tìm thấy bảng để xuất Excel.", true);
      return;
    }

    const header = [];
    $table.find("thead th").each(function () {
      header.push(($(this).text() || "").trim());
    });

    const aoa = [header];
    $table.find("tbody tr:visible").each(function () {
      const row = [];
      $(this)
        .find("td")
        .each(function () {
          row.push(($(this).text() || "").trim());
        });
      // chỉ push khi có dữ liệu
      if (row.some((c) => c !== "")) aoa.push(row);
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = header.map((_, i) => {
      // ước lượng độ rộng theo max length từng cột
      let maxLen = header[i] ? header[i].length : 10;
      for (let r = 1; r < aoa.length; r++) {
        const v = (aoa[r][i] || "").toString();
        if (v.length > maxLen) maxLen = v.length;
      }
      return { wch: Math.min(Math.max(maxLen + 2, 8), 40) };
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName || "Sheet1");
    XLSX.writeFile(wb, fileName || "export.xlsx");
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
          const memberName = dm.name;

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

      const sttFromFile = parseInt((row[0] || "").toString().trim() || "0", 10);
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
      const rankRaw = row[12] || "";
      const medalRaw = row[13] || "";
      const manualMedal = normalizeMedal(medalRaw);

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
          stt: sttFromFile > 0 ? sttFromFile : getNextStt(),
          rank: parseInt((rankRaw || "").toString().trim() || "0", 10) || "",
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
          manualMedal,
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
  // Manual sort bảng chấm điểm (chỉ sort khi bấm nút)
  // - Ưu tiên theo: Lứa tuổi -> Nội dung -> Xếp hạng -> Tổng điểm -> Tên
  // - Giữ nguyên block đồng đội: MASTER đi cùng MEMBER
  // ======================
  function sortScoreTableManual() {
    // đảm bảo total/medal đang đúng trước khi sort
    recalcAllScores();

    const allowDoubleBronze = $("#chkDoubleBronze").is(":checked");

    const $tb = $("#tblScores tbody");
    const $rows = $tb.find("tr");
    if ($rows.length <= 1) return;

    // build units (single/master + members)
    const units = [];
    const taken = new Set();

    $rows.each(function () {
      const $tr = $(this);
      const id = $tr.data("id");
      if (taken.has(id)) return;

      const type = $tr.data("type") || "single";
      if (type === "member") return;

      const block = [$tr.get(0)];

      if (type === "master") {
        const masterId = $tr.data("id");
        const $members = $tb.find(`tr[data-parent="${masterId}"]`);
        $members.each(function () {
          block.push(this);
          taken.add($(this).data("id"));
        });
      }

      const d = getRowData($tr);
      const age = (d.age || "").trim();
      const ev = (d.event || "").trim();
      const isCombat = isCombatEvent(ev);

      const finalMedal = normalizeMedal(
        ((($tr.attr("data-final-medal") || "") + "").trim() || (d.medal || ""))
      );

      const medalOrder =
        finalMedal === "Vàng" ? 1 : finalMedal === "Bạc" ? 2 : finalMedal === "Đồng" ? 3 : 99;

      const total = parseFloat(($tr.find(".total").text() || "0").toString().replace(",", ".")) || 0;
      const name = (d.name || "").trim();

      // xếp hạng chỉnh tay (ưu tiên làm tie-break khi cùng huy chương)
      const mrRaw = ($tr.find(".inp-rank").val() ?? "").toString().trim();
      const manualRank = mrRaw === "" ? null : (parseInt(mrRaw, 10) || null);

      units.push({
        block,
        $head: $tr,
        age,
        ev,
        isCombat,
        medalOrder,
        total,
        name,
        finalMedal,
        manualRank,
      });
    });

    const viCmp = (a, b) => a.localeCompare(b, "vi", { sensitivity: "base" });

    units.sort((a, b) => {
      // group: lứa tuổi -> nội dung
      const c1 = viCmp(a.age, b.age);
      if (c1 !== 0) return c1;
      const c2 = viCmp(a.ev, b.ev);
      if (c2 !== 0) return c2;

      // trong group:
      if (a.isCombat || b.isCombat) {
        // đối kháng: ưu tiên thành tích
        if (a.medalOrder !== b.medalOrder) return a.medalOrder - b.medalOrder;

        // cùng huy chương: cho phép điều chỉnh thứ tự bằng cách nhập Xếp hạng tay
        // ví dụ 2 VĐV cùng "Đồng": nhập 3/4 để swap, sau đó bấm Sort
        const ar = a.manualRank;
        const br = b.manualRank;
        if (ar != null || br != null) {
          if (ar == null) return 1;
          if (br == null) return -1;
          if (ar !== br) return ar - br;
        }

        return viCmp(a.name, b.name);
      }

      // biểu diễn/quyền: ưu tiên huy chương (nếu có), rồi tổng điểm, rồi Xếp hạng nhập tay khi bị hoà
      if (a.medalOrder !== b.medalOrder) return a.medalOrder - b.medalOrder;

      const at = parseFloat(round2(a.total) || "0");
      const bt = parseFloat(round2(b.total) || "0");
      if (bt !== at) return bt - at;

      // cùng huy chương + cùng điểm: cho phép đảo thứ tự bằng cách nhập Xếp hạng tay (3/4...)
      const ar2 = a.manualRank;
      const br2 = b.manualRank;
      if (ar2 != null || br2 != null) {
        if (ar2 == null) return 1;
        if (br2 == null) return -1;
        if (ar2 !== br2) return ar2 - br2;
      }

      return viCmp(a.name, b.name);
    });

    // Re-append theo thứ tự mới
    const frag = document.createDocumentFragment();
    units.forEach((u) => {
      u.block.forEach((el) => frag.appendChild(el));
    });
    $tb.get(0).appendChild(frag);

    // ====== Gán XẾP HẠNG sau khi sort (STT giữ nguyên) ======
    const groupMap = new Map(); // key -> array units in display order
    units.forEach((u) => {
      const key = `${u.age}||${u.ev}`;
      if (!groupMap.has(key)) groupMap.set(key, []);
      groupMap.get(key).push(u);
    });

    groupMap.forEach((list) => {
      const isCombat = list[0]?.isCombat;

      if (isCombat) {
        // Đối kháng: gán theo huy chương. Nếu bật double bronze: Đồng #1 = hạng 3, Đồng #2 = hạng 4
        let bronzeCount = 0;

        list.forEach((u) => {
          const $rank = u.$head.find(".inp-rank");
          if (!$rank.length) return;

          const medal = u.finalMedal;
          if (medal === "Vàng") $rank.val("1");
          else if (medal === "Bạc") $rank.val("2");
          else if (medal === "Đồng") {
            bronzeCount++;
            if (bronzeCount === 1) $rank.val("3");
            else if (bronzeCount === 2 && allowDoubleBronze) $rank.val("4");
            else $rank.val("");
          } else {
            $rank.val("");
          }
        });

        return;
      }

      // Biểu diễn: gán theo tổng điểm (chỉ dòng có total > 0)
      let r = 0;
      list.forEach((u) => {
        const $rank = u.$head.find(".inp-rank");
        if (!$rank.length) return;
        if (u.total > 0) {
          r++;
          $rank.val(String(r));
        } else {
          $rank.val("");
        }
      });
    });

    // áp lại filter (không đổi STT)
    applyFilters();

    toast("Đã sắp xếp bảng chấm điểm và cập nhật Xếp hạng.");
  }

// ======================
  // Events
  // ======================
  // Điểm: debounce để nhập nhanh
  $(document).on("input", ".inp-g", function () {
    requestRecalc();
  });

  // Xếp hạng: cho phép chỉnh tay, chỉ auto đổi khi bấm Sort
  $(document).on("input change", ".inp-rank", function () {
    // không recalc để tránh nhảy focus; chỉ lưu
    scheduleAutoSave();
  });

  // Thông tin: đổi là recalc
  $(document).on(
    "change",
    ".inp-age, .inp-event, .inp-yob, .inp-team",
    function () {
      requestRecalc();
    }
  );

  // Tên VĐV: contenteditable
  $(document).on("input", ".ath-name", function () {
    requestRecalc();
  });

  // Huy chương: là input để sửa dễ, nhưng chỉ commit khi blur/change (tránh overwrite khi đang gõ)
  // Update realtime để "Thành tích" đổi thì "Xếp hạng" (đặc biệt nhóm Đối kháng) đổi ngay
  $(document).on("input", "#tblScores tbody .inp-medal", function () {
    const $tr = $(this).closest("tr");
    const type = $tr.data("type") || "single";
    if (type === "member") return;
    requestRecalc();
  });

  $(document).on("keydown", "#tblScores tbody .inp-medal", function (e) {
    if (e.key === "Enter") {
      e.preventDefault();
      $(this).blur();
    }
  });

  $(document).on("blur change", "#tblScores tbody .inp-medal", function () {
    const $inp = $(this);
    const $tr = $inp.closest("tr");
    const type = $tr.data("type") || "single";
    if (type === "member") return;

    const raw = (($inp.val() || "") + "").trim();
    if (!raw) {
      $tr.removeAttr("data-manual-medal");
      requestRecalc();
      return;
    }

    const norm = normalizeMedal(raw);
    if (!norm) {
      toast('Thành tích chỉ nhận: V/B/Đ hoặc Vàng/Bạc/Đồng', true);
      $inp.val("");
      $tr.removeAttr("data-manual-medal");
      requestRecalc();
      return;
    }

    const auto = normalizeMedal(($tr.attr("data-auto-medal") || "").trim());
    if (auto && norm === auto) {
      // trùng với auto => hiểu là không override
      $tr.removeAttr("data-manual-medal");
    } else {
      $tr.attr("data-manual-medal", norm);
    }
    $inp.val(norm);
    requestRecalc();
  });

  // Điều hướng ô nhập: Enter + phím mũi tên để nhảy sang ô tiếp theo
  function getNavElements($tr) {
    const els = [];
    const push = (selector) => {
      const el = $tr.find(selector).get(0);
      if (el && $(el).is(":visible") && !$(el).is(":disabled")) els.push(el);
    };

    push(".inp-age");
    push(".inp-event");
    const nameEl = $tr.find('.ath-name[contenteditable="true"]').get(0);
    if (nameEl && $(nameEl).is(":visible")) els.push(nameEl);
    push(".inp-yob");
    push(".inp-team");
    push(".g1:not(:disabled)");
    push(".g2:not(:disabled)");
    push(".g3:not(:disabled)");
    push(".g4:not(:disabled)");
    push(".g5:not(:disabled)");
    push(".inp-rank");
    push(".inp-medal");

    return els;
  }

  function focusEl(el) {
    if (!el) return;
    const $el = $(el);
    if ($el.is("[contenteditable=true]")) {
      $el.focus();
      // đưa caret về cuối
      const range = document.createRange();
      range.selectNodeContents(el);
      range.collapse(false);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      return;
    }
    $el.focus();
    if ($el.is("input")) {
      try {
        el.select();
      } catch (_) {}
    }
  }

  $(document).on(
    "keydown",
    "#tblScores tbody input, #tblScores tbody td[contenteditable=true]",
    function (e) {
      const key = e.key;
      const isNavKey =
        key === "Enter" ||
        key === "ArrowRight" ||
        key === "ArrowLeft" ||
        key === "ArrowUp" ||
        key === "ArrowDown";
      if (!isNavKey) return;

      const $tr = $(this).closest("tr");
      const nav = getNavElements($tr);
      const idx = nav.indexOf(this);
      if (idx < 0) return;

      e.preventDefault();

      if (key === "Enter" || key === "ArrowRight") {
        focusEl(nav[Math.min(idx + 1, nav.length - 1)]);
        return;
      }

      if (key === "ArrowLeft") {
        focusEl(nav[Math.max(idx - 1, 0)]);
        return;
      }

      // Up/Down: cùng cột (theo index) sang hàng trước/sau
      const $rows = $("#tblScores tbody tr:visible");
      const rowIndex = $rows.index($tr);
      const targetRow =
        key === "ArrowUp" ? $rows.eq(rowIndex - 1) : $rows.eq(rowIndex + 1);
      if (!targetRow.length) return;

      const targetNav = getNavElements(targetRow);
      focusEl(targetNav[Math.min(idx, targetNav.length - 1)]);
    }
  );

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
        stt: getNextStt(),
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

  

  // Sort thủ công: chỉ khi bấm nút
  $("#btnSortRows").on("click", function () {
    sortScoreTableManual();
  });

  // Export Excel BXH huy chương
  $("#btnExportMedalAthletes").on("click", function () {
    exportHtmlTableToExcel("#tblMedalAthletes", "bxh_huy_chuong_vdv.xlsx", "BXH_VDV");
  });
  $("#btnExportMedalTeams").on("click", function () {
    exportHtmlTableToExcel("#tblMedalTeams", "bxh_huy_chuong_doan.xlsx", "BXH_DOAN");
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
  // Fullscreen cho BXH VĐV & Đoàn
  // ======================
  function toggleFullscreenFor(targetSelector, btn) {
    const el = document.querySelector(targetSelector);
    if (!el) return;

    if (!document.fullscreenElement) {
      if (el.requestFullscreen) {
        el.requestFullscreen();
        el.classList.add("fullscreen-active");
        btn.classList.add("is-fullscreen");
        btn.textContent = "⤡ Thoát full màn";
      }
    } else {
      document.exitFullscreen();
    }
  }
  // Nút phóng to / thu nhỏ từng bảng
  $(document).on("click", ".btn-fullscreen", function () {
    const target = this.getAttribute("data-target");
    if (!target) return;
    toggleFullscreenFor(target, this);
  });

  // Khi thoát fullscreen bằng ESC hoặc nút trình duyệt
  document.addEventListener("fullscreenchange", function () {
    if (!document.fullscreenElement) {
      // remove class trên panel
      $(".table-panel.fullscreen-active").removeClass("fullscreen-active");
      // reset nút
      $(".btn-fullscreen.is-fullscreen")
        .removeClass("is-fullscreen")
        .each(function () {
          this.textContent = "⤢ Toàn màn hình";
        });
    }
  });

  // ======================
  // Init
  // ======================
  $(function () {
    loadFromLocal();
    recalcAllScores();
  });
})(jQuery);
