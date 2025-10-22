(function($){
  const LS_KEY = "vovinam.medals.v2.1";

  // ===== Focus-lock để không mất con trỏ khi gõ =====
  let isTyping = false;

  function mode(){ return $('input[name="calcMode"]:checked').val() || "score"; }
  function applyModeUI(){
    const m = mode();
    if(m === "count"){ $(".col-pts").hide(); } else { $(".col-pts").show(); }
    const $weightsRow = $("#wGold, #wSilver, #wBronze").closest("label");
    $weightsRow.css("opacity", m==="count"?.5:1);
  }
  function weights(){
    return {
      g: parseInt($("#wGold").val()||3,10),
      s: parseInt($("#wSilver").val()||2,10),
      b: parseInt($("#wBronze").val()||1,10)
    };
  }
  function rowData($tr){
    const g = parseInt($tr.find(".inp-medal.g").val()||0,10);
    const s = parseInt($tr.find(".inp-medal.s").val()||0,10);
    const b = parseInt($tr.find(".inp-medal.b").val()||0,10);
    const w = weights();
    const pts = g*w.g + s*w.s + b*w.b;
    const sum = g+s+b;
    const name = $tr.find(".ath-name").text().trim();
    const team = $tr.find(".inp-team").val().trim();
    return {id:$tr.data("id"), name, team, g,s,b,sum,pts};
  }
  function recalcAthleteRow($tr){
    const d = rowData($tr);
    $tr.find(".sum").text(d.sum);
    $tr.find(".pts").text(d.pts);
    $tr.attr("data-team", d.team);
    return d;
  }
  function sortAthletes(){
    const $tbody = $("#tblAthletes tbody");
    const $rows = $tbody.children("tr").get();
    const m = mode();
    $rows.sort((a,b)=>{
      const da = rowData($(a)), db = rowData($(b));
      if(m === "score"){
        if(db.pts!==da.pts) return db.pts-da.pts;
      }
      if(db.g!==da.g) return db.g-da.g;
      if(db.s!==da.s) return db.s-da.s;
      if(db.b!==da.b) return db.b-da.b;
      return da.name.localeCompare(db.name,'vi',{sensitivity:'base'});
    });
    $rows.forEach((tr,i)=>{
      tr.querySelector(".rank").textContent = (i+1);
      $tbody.append(tr);
    });
  }
  function aggregateTeams(){
    const map = new Map();
    $("#tblAthletes tbody tr").each(function(){
      const d = rowData($(this));
      if(!d.team) return;
      if(!map.has(d.team)) map.set(d.team, {team:d.team,g:0,s:0,b:0,sum:0,pts:0});
      const t = map.get(d.team);
      t.g+=d.g; t.s+=d.s; t.b+=d.b; t.sum+=d.sum; t.pts+=d.pts;
    });
    const m = mode();
    return Array.from(map.values()).sort((a,b)=>{
      if(m === "score"){ if(b.pts!==a.pts) return b.pts-a.pts; }
      if(b.g!==a.g) return b.g-a.g;
      if(b.s!==a.s) return b.s-a.s;
      if(b.b!==a.b) return b.b-a.b;
      return a.team.localeCompare(b.team,'vi',{sensitivity:'base'});
    });
  }
  function renderTeams(){
    const teams = aggregateTeams();
    const $tb = $("#tblTeams tbody").empty();
    teams.forEach((t,idx)=>{
      const tr = `
        <tr>
          <td class="rank">${idx+1}</td>
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
  function recalcAll(){
    $("#tblAthletes tbody tr").each(function(){ recalcAthleteRow($(this)); });
    if(!isTyping){ sortAthletes(); }
    renderTeams();
    applyModeUI();
  }

  function saveToLocal(){
    const payload = {
      mode: mode(),
      w: weights(),
      rows: $("#tblAthletes tbody tr").map(function(){
        const $tr = $(this);
        return {
          id: $tr.data("id"),
          name: $tr.find(".ath-name").text().trim(),
          team: $tr.find(".inp-team").val().trim(),
          g: parseInt($tr.find(".inp-medal.g").val()||0,10),
          s: parseInt($tr.find(".inp-medal.s").val()||0,10),
          b: parseInt($tr.find(".inp-medal.b").val()||0,10)
        };
      }).get()
    };
    localStorage.setItem(LS_KEY, JSON.stringify(payload));
  }
  function loadFromLocal(){
    const raw = localStorage.getItem(LS_KEY);
    if(!raw){ recalcAll(); return; }
    try{
      const data = JSON.parse(raw);
      if(data.mode){ $('input[name="calcMode"][value="'+data.mode+'"]').prop('checked', true); }
      if(data.w){
        $("#wGold").val(data.w.g??3);
        $("#wSilver").val(data.w.s??2);
        $("#wBronze").val(data.w.b??1);
      }
      const $tb = $("#tblAthletes tbody").empty();
      (data.rows||[]).forEach(r=>{
        $tb.append(makeRowHtml(r.id||uid(), r.name||"", r.team||"", r.g||0, r.s||0, r.b||0));
      });
    }catch(e){ console.warn("Load local failed", e); }
    recalcAll();
  }

  function uid(){ return "a"+Math.random().toString(36).slice(2,9); }
  function makeRowHtml(id, name, team, g,s,b){
    return `
      <tr data-id="${id}" data-team="${escapeAttr(team||"")}">
        <td class="rank">–</td>
        <td contenteditable="true" class="ath-name">${escapeHtml(name||"Tên VĐV")}</td>
        <td><input class="inp-team" type="text" value="${escapeAttr(team||"")}" /></td>
        <td><input class="inp-medal g" type="number" min="0" value="${g??0}" /></td>
        <td><input class="inp-medal s" type="number" min="0" value="${s??0}" /></td>
        <td><input class="inp-medal b" type="number" min="0" value="${b??0}" /></td>
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
  function escapeHtml(s){ return (s+"").replace(/[&<>"']/g, m=>({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[m])); }
  function escapeAttr(s){ return escapeHtml(s).replace(/"/g,"&quot;"); }

  // ===== CSV / JSON =====
  function dumpJson(){
    return {
      mode: mode(),
      weights: weights(),
      athletes: $("#tblAthletes tbody tr").map(function(){
        const d = rowData($(this));
        return {id:d.id,name:d.name,team:d.team,g:d.g,s:d.s,b:d.b};
      }).get()
    };
  }
  function buildCsv(){
    const m = mode();
    let a = m==="score"
      ? "Rank,Name,Team,Gold,Silver,Bronze,Total,Points\n"
      : "Rank,Name,Team,Gold,Silver,Bronze,Total\n";
    $("#tblAthletes tbody tr").each(function(){
      const d = rowData($(this));
      const rank = $(this).find(".rank").text().trim()||"";
      a += m==="score"
        ? [rank, csvEsc(d.name), csvEsc(d.team), d.g, d.s, d.b, d.sum, d.pts].join(",")+"\n"
        : [rank, csvEsc(d.name), csvEsc(d.team), d.g, d.s, d.b, d.sum].join(",")+"\n";
    });
    let t = m==="score"
      ? "Rank,Team,Gold,Silver,Bronze,Total,Points\n"
      : "Rank,Team,Gold,Silver,Bronze,Total\n";
    $("#tblTeams tbody tr").each(function(){
      const rank = $(this).find(".rank").text().trim();
      const cells = $(this).children("td");
      const team = cells.eq(1).text().trim();
      const g = cells.eq(2).text().trim();
      const s = cells.eq(3).text().trim();
      const b = cells.eq(4).text().trim();
      const sum = cells.eq(5).text().trim();
      const pts = cells.eq(6).text().trim();
      t += m==="score"
        ? [rank, csvEsc(team), g,s,b,sum,pts].join(",")+"\n"
        : [rank, csvEsc(team), g,s,b,sum].join(",")+"\n";
    });
    return {csvAthletes:a, csvTeams:t};
  }
  function csvEsc(s){ s = (s??"").toString(); return /[",\n]/.test(s) ? `"${s.replace(/"/g,'""')}"` : s; }
  function download(filename, text){
    const BOM = "\uFEFF"; // UTF-8 BOM cho Excel Windows
    const blob = new Blob([BOM + text], {type:"text/csv;charset=utf-8"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);
  }

  // ===== Events =====
  // Focus-lock: đặt cờ khi gõ
  $(document).on("focusin", ".ath-name, .inp-team, .inp-medal", function(){ isTyping = true; });
  $(document).on("blur", ".ath-name, .inp-team, .inp-medal", function(){
    isTyping = false;
    recalcAll();
    saveToLocal();
  });

  // Thay đổi số/đội/trọng số: chỉ recalc, không sort nếu đang gõ
  $(document).on("input change", ".inp-medal, .inp-team, #wGold, #wSilver, #wBronze", function(){
    const $tr = $(this).closest("tr");
    if($tr.length){ recalcAthleteRow($tr); }
    if(!isTyping){ sortAthletes(); }
    renderTeams();
    saveToLocal();
  });

  // Tên VĐV (contenteditable): recalc tức thì, sort khi blur
  $(document).on("input", ".ath-name", function(){
    const $tr = $(this).closest("tr");
    recalcAthleteRow($tr);
    renderTeams();
  });

  // Nút tăng nhanh: luôn sort ngay (vì không gõ tay)
  $(document).on("click", ".btn-inc", function(){
    const type = $(this).data("type"); // G S B
    const $tr = $(this).closest("tr");
    const cls = type==="G" ? ".g" : type==="S" ? ".s" : ".b";
    const $inp = $tr.find(".inp-medal"+cls);
    const val = parseInt($inp.val()||0,10)+1;
    $inp.val(val);
    isTyping = false;
    recalcAll(); saveToLocal();
  });

  $("#btnAddRow").on("click", function(){
    $("#tblAthletes tbody").append(makeRowHtml(uid(), "VĐV mới", "", 0,0,0));
    recalcAll(); saveToLocal();
  });

  $(document).on("click", ".btnDelRow", function(){
    $(this).closest("tr").remove();
    recalcAll(); saveToLocal();
  });

  $("#search").on("input", function(){
    const q = $(this).val().toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu,"");
    $("#tblAthletes tbody tr").each(function(){
      const name = $(this).find(".ath-name").text().toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu,"");
      const team = ($(this).find(".inp-team").val()+"").toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu,"");
      $(this).toggle(!q || name.includes(q) || team.includes(q));
    });
  });

  $("#btnSave").on("click", function(){ saveToLocal(); toast("Đã lưu vào máy."); });
  $("#btnReset").on("click", function(){
    if(confirm("Xoá toàn bộ dữ liệu đang lưu trong máy?")){ localStorage.removeItem(LS_KEY); location.reload(); }
  });

  $("#btnDump").on("click", function(){
    $("#ioJson").val(JSON.stringify(dumpJson(), null, 2));
    toast("Đã xuất JSON bên dưới.");
  });

  $("#btnLoad").on("click", function(){
    const raw = $("#ioJson").val().trim();
    if(!raw){ toast("Chưa có JSON để nhập.", true); return; }
    try{
      const data = JSON.parse(raw);
      if(data.mode){ $('input[name="calcMode"][value="'+data.mode+'"]').prop('checked', true); }
      if(data.weights){
        $("#wGold").val(data.weights.g??3);
        $("#wSilver").val(data.weights.s??2);
        $("#wBronze").val(data.weights.b??1);
      }
      const $tb = $("#tblAthletes tbody").empty();
      (data.athletes||[]).forEach(r=>{
        $tb.append(makeRowHtml(r.id||uid(), r.name||"", r.team||"", r.g||0, r.s||0, r.b||0));
      });
      recalcAll(); saveToLocal();
      toast("Đã nhập JSON.");
    }catch(e){ toast("JSON không hợp lệ.", true); }
  });

  $("#btnExportCsv").on("click", function(){
    const {csvAthletes, csvTeams} = buildCsv();
    download("athletes.csv", csvAthletes);
    download("teams.csv", csvTeams);
  });

  $(document).on("change", 'input[name="calcMode"]', function(){
    applyModeUI();
    recalcAll(); saveToLocal();
  });

  function toast(msg, isErr){
    const el = $("<div/>").text(msg).css({
      position:"fixed",bottom:"24px",right:"24px",padding:"10px 14px",
      background:isErr?"#3b0d0d":"#062e2e",border:"1px solid #0d4d4d",color:"#e5f5f5",
      borderRadius:"10px",zIndex:9999,boxShadow:"0 10px 30px rgba(0,0,0,.35)"
    });
    $("body").append(el);
    setTimeout(()=>el.fadeOut(300,()=>el.remove()),1800);
  }

  // init
  $(function(){
    loadFromLocal();
    applyModeUI();
  });
})(jQuery);