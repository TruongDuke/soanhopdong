/* Soạn hợp đồng Dương Nguyễn 79 — chạy hoàn toàn trên trình duyệt, không server, không AI */
(function(){
  "use strict";
  var $ = function(id){ return document.getElementById(id); };
  var DT = window.docxtemplater || window.Docxtemplater;
  if (DT && DT.default) DT = DT.default;

  /* ---------- tiện ích ---------- */
  function fmtNum(n){ n=Math.round(Number(n)||0); return n.toLocaleString("vi-VN").replace(/,/g,"."); }
  function onlyDigits(s){ return (s||"").replace(/[^\d]/g,""); }
  function pad2(x){ x=String(x); return x.length<2?"0"+x:x; }
  function ddmmyyyy(d){ return pad2(d.getDate())+"/"+pad2(d.getMonth()+1)+"/"+d.getFullYear(); }
  function fromQrDate(s){ s=onlyDigits(s); return s.length===8 ? s.slice(0,2)+"/"+s.slice(2,4)+"/"+s.slice(4) : s; }
  function addMonths(d,m){ var y=d.getFullYear(),mo=d.getMonth()+m,day=d.getDate();
    var last=new Date(y,mo+1,0).getDate(); return new Date(y,mo,Math.min(day,last)); }
  function parseDateInput(v){ if(!v) return null; var p=v.split("-"); return new Date(+p[0],+p[1]-1,+p[2]); }

  /* ---------- điền từ chuỗi QR CCCD ---------- */
  function fillFromQr(str){
    var p = String(str).trim().split("|");
    if(p.length < 6){ alert("Chuỗi QR không đúng định dạng CCCD."); return false; }
    // [0]=CCCD [1]=CMND [2]=Họ tên [3]=NgàySinh [4]=Giới tính [5]=Địa chỉ [6]=Ngày cấp
    setVal("ho_ten", p[2]);
    setVal("cccd", p[0]);
    if(p[6]) setVal("cap_ngay", fromQrDate(p[6]));
    setVal("dia_chi", p[5]);
    setVal("cho_o", p[5]);
    ["ho_ten","cccd","cap_ngay","dia_chi","cho_o"].forEach(function(id){ $(id).classList.add("filled"); });
    return true;
  }
  function setVal(id,v){ if($(id)) $(id).value = v||""; }

  /* ---------- tự tính + hiển thị ---------- */
  function recompute(){
    var vay = Number(onlyDigits($("tien_vay").value));
    var box = $("autoBox");
    if(!vay){ box.textContent = "Nhập tiền vay & ngày vay để xem phần tự tính…"; return null; }
    var phi = Math.round(vay*0.06);
    var dv = parseDateInput($("ngay_vay").value);
    var info = {
      phi_thue: fmtNum(phi), phi_thue_chu: window.docSoThanhChu(phi),
      tien_vay: fmtNum(vay), tien_vay_chu: window.docSoThanhChu(vay)
    };
    if(dv){
      var den = addMonths(dv,3);
      info.tu_ngay = ddmmyyyy(dv); info.den_ngay = ddmmyyyy(den);
      info.ngay_tt = pad2(dv.getDate());
      info.ngay_lap = pad2(dv.getDate()); info.thang_lap = pad2(dv.getMonth()+1); info.nam_lap = String(dv.getFullYear());
    }
    box.innerHTML =
      "Tiền vay: <b>"+info.tien_vay+" đ</b> ("+info.tien_vay_chu+")<br>"+
      "Phí thuê/tháng (6%): <b>"+info.phi_thue+" đ</b> ("+info.phi_thue_chu+")<br>"+
      (dv ? "Đáo hạn: <b>"+info.den_ngay+"</b> · thu ngày <b>"+info.ngay_tt+"</b> hàng tháng" : "<i>Chọn ngày vay để tính đáo hạn</i>");
    return info;
  }

  /* ---------- gom dữ liệu để render ---------- */
  function collect(){
    var vay = Number(onlyDigits($("tien_vay").value));
    var gia = Number(onlyDigits($("gia").value));
    var dv = parseDateInput($("ngay_vay").value) || new Date();
    var den = addMonths(dv,3);
    var phi = Math.round(vay*0.06);
    return {
      so_hd: $("so_hd").value || "……",
      ngay_lap: pad2(dv.getDate()), thang_lap: pad2(dv.getMonth()+1), nam_lap: String(dv.getFullYear()),
      ho_ten: $("ho_ten").value, cccd: $("cccd").value, cap_ngay: $("cap_ngay").value,
      noi_cap: $("noi_cap").value, dia_chi: $("dia_chi").value, cho_o: $("cho_o").value,
      dien_thoai: $("dien_thoai").value,
      may: $("may").value, mau: $("mau").value, imei: $("imei").value,
      cau_hinh: $("cau_hinh").value, tinh_trang: $("tinh_trang").value,
      tien_vay: vay?fmtNum(vay):"……", tien_vay_chu: vay?window.docSoThanhChu(vay):"……",
      tu_ngay: ddmmyyyy(dv), den_ngay: ddmmyyyy(den), ngay_tt: pad2(dv.getDate()),
      phi_thue: vay?fmtNum(phi):"……", phi_thue_chu: vay?window.docSoThanhChu(phi):"……",
      gia: gia?fmtNum(gia):"……", gia_chu: gia?window.docSoThanhChu(gia):"……",
      hinh_thuc_tt: $("hinh_thuc_tt").value
    };
  }

  /* ---------- render 1 template ---------- */
  function loadBin(url){
    return new Promise(function(res,rej){
      PizZipUtils.getBinaryContent(url, function(err,data){ err?rej(err):res(data); });
    });
  }
  var TPL = { camdo:"templates/camdo.docx", thuemay:"templates/thuemay.docx", muaban:"templates/muaban.docx" };
  var FNAME = { camdo:"HopDong_CamDo", thuemay:"HopDong_ThueMay", muaban:"GiayMuaBan" };

  function genOne(key, data){
    return loadBin(TPL[key]).then(function(content){
      var zip = new PizZip(content);
      var doc = new DT(zip, {paragraphLoop:true, linebreaks:true});
      doc.render(data);
      var blob = doc.getZip().generate({type:"blob",
        mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
      var ten = (data.ho_ten||"khach").replace(/\s+/g,"_");
      saveAs(blob, FNAME[key]+"_"+ten+".docx");
    });
  }

  function msg(t){ $("genMsg").textContent = t; }

  /* ---------- quét QR bằng camera ---------- */
  var scanner=null;
  function startScan(){
    if(!window.Html5Qrcode){ alert("Chưa tải được bộ quét."); return; }
    scanner = new Html5Qrcode("reader");
    $("btnScan").style.display="none"; $("btnStop").style.display="block";
    scanner.start({facingMode:"environment"}, {fps:10, qrbox:{width:240,height:240}},
      function(txt){ if(fillFromQr(txt)){ recompute(); stopScan(); msg("Đã đọc CCCD ✓"); } },
      function(){/* bỏ qua khung không đọc được */})
    .catch(function(e){ alert("Không mở được camera: "+e); resetScanBtn(); });
  }
  function stopScan(){ if(scanner){ scanner.stop().then(function(){ scanner.clear(); scanner=null; resetScanBtn(); }).catch(resetScanBtn); } else resetScanBtn(); }
  function resetScanBtn(){ $("btnScan").style.display="block"; $("btnStop").style.display="none"; }

  /* ---------- đọc CHỮ trên ảnh CCCD bằng OCR (Tesseract, tiếng Việt) ---------- */
  // chỉ loại chữ CHẮC CHẮN là tiêu đề — KHÔNG loại Công/Dân/Nam/Việt vì hay có trong tên thật
  var BAD_NAME = {"CỘNG":1,"CƯỚC":1,"CARD":1,"REPUBLIC":1,"SOCIALIST":1,"IDENTITY":1,"CITIZEN":1,
    "FULL":1,"NAME":1,"HỌ":1,"TÊN":1,"HẠNH":1,"PHÚC":1,"ĐỘC":1,"LẬP":1,"FREEDOM":1,"INDEPENDENCE":1,"PLACE":1,"RESIDENCE":1};
  var VOWELS = "aàáảãạăằắẳẵặâầấẩẫậeèéẻẽẹêềếểễệiìíỉĩịoòóỏõọôồốổỗộơờớởỡợuùúủũụưừứửữựyỳýỷỹỵ";
  function hasVowel(w){ w=w.toLowerCase(); for(var i=0;i<w.length;i++){ if(VOWELS.indexOf(w[i])>=0) return true; } return false; }
  function nameFrom(s){
    var runs = s.match(/[A-ZÀ-ỸĐ]{2,}(?:\s+[A-ZÀ-ỸĐ]+)*/g) || [];
    var best = "";
    runs.forEach(function(r){
      // token tên tiếng Việt phải có nguyên âm → loại rác như "HP", "TT"
      var toks = r.split(/\s+/).filter(function(w){ return w.length>=2 && !BAD_NAME[w] && hasVowel(w); });
      if(toks.length>=1 && toks.length<=5){ var c = toks.join(" "); if(c.length>=6 && c.length>best.length) best = c; }
    });
    return best;
  }
  function cleanAddr(s){
    s = s.replace(/c[oó]\s*g[^,]*?\d{4}[^\s,]*/ig, "");           // bỏ "Có giá trị đến ...2037."
    var prev = null;
    while(prev !== s){
      prev = s;
      s = s.replace(/^[^0-9A-Za-zÀ-ỹ]+/, "");                     // rác đầu
      s = s.replace(/^(í|i|place|of|residence|ofresidence|resi|ofresi|ne|e)['’:\s]+/i, ""); // mảnh nhãn Anh
    }
    s = s.replace(/\S*[®©=|{}§¢∞~^`]\S*/g, " ");                  // bỏ cụm dính ký hiệu rác OCR
    s = s.replace(/(\d)\s+(\d)/g, "$1$2");                        // "291 0" -> "2910"
    s = s.replace(/\s+/g, " ").replace(/\s*,\s*/g, ", ").replace(/[\s,\-]+$/, "").trim();
    return s;
  }
  function fillFromOCR(text){
    var lines = text.split(/\r?\n/).map(function(s){return s.trim();}).filter(Boolean);
    var name = "";
    for(var i=0;i<lines.length;i++){
      if(/Họ và tên|Full\s*?name/i.test(lines[i])){
        [lines[i], lines[i+1]||""].forEach(function(c){ var n=nameFrom(c); if(n.length>name.length) name=n; });
        if(name) break;
      }
    }
    // CCCD: tìm dòng "Số/No", lấy số của dòng đó (12 số cuối, hoặc giữ nguyên nếu OCR thiếu số để NV bổ sung)
    var cccd = "";
    var numRx = /s[oốô]\s*\/?\s*no|\/no|no\s*[:.]|s[ốô]\b|^s[oố]|só/i;
    for(var j=0;j<lines.length;j++){
      if(numRx.test(lines[j])){ var dd=lines[j].replace(/\D/g,""); if(dd.length>=10){ cccd = dd.length>=12 ? dd.slice(-12) : dd; break; } }
    }
    if(!cccd){ for(var j2=0;j2<lines.length;j2++){ var d2=lines[j2].replace(/\D/g,""); if(d2.length>=12){ cccd=d2.slice(-12); break; } } }
    // Địa chỉ thường trú: bắt từ nhãn, nối dòng kế, rồi dọn mảnh nhãn Anh + chữ "Có giá trị đến"
    var addr = "";
    for(var k=0;k<lines.length;k++){
      var mk=lines[k].match(/(?:thường\s*trú|residence|resi)[^\n]*/i);
      if(mk){
        var raw = lines[k].slice(mk.index) + " " + (lines[k+1]||"");
        raw = raw.replace(/^(nơi\s*)?thường\s*trú/i,"");
        addr = cleanAddr(raw);
        break;
      }
    }
    var got = false;
    if(name){ setVal("ho_ten",name); $("ho_ten").classList.add("filled"); got=true; }
    if(cccd){ setVal("cccd",cccd); $("cccd").classList.add("filled"); got=true; }
    if(addr){ setVal("dia_chi",addr); setVal("cho_o",addr); $("dia_chi").classList.add("filled"); $("cho_o").classList.add("filled"); got=true; }
    return got;
  }
  // tiền xử lý: phóng to + xám hoá + kéo giãn tương phản → OCR ăn hơn nhiều
  function preprocess(img){
    var MAXW = 2200, scale = Math.min(2.2, MAXW/img.naturalWidth); if(scale < 1) scale = 1;
    var w = Math.round(img.naturalWidth*scale), h = Math.round(img.naturalHeight*scale);
    var c = document.createElement("canvas"); c.width = w; c.height = h;
    var ctx = c.getContext("2d"); ctx.drawImage(img, 0, 0, w, h);
    try{
      var id = ctx.getImageData(0,0,w,h), d = id.data, n = d.length, mn = 255, mx = 0, i, g;
      var gray = new Uint8Array(n>>2);
      for(i=0;i<n;i+=4){ g=(d[i]*0.299+d[i+1]*0.587+d[i+2]*0.114)|0; gray[i>>2]=g; if(g<mn)mn=g; if(g>mx)mx=g; }
      var lo=mn, hi=mx; if(hi-lo<40) hi=lo+40; var range=hi-lo||1;
      for(i=0;i<n;i+=4){ var v=(gray[i>>2]-lo)*255/range; v=v<0?0:v>255?255:v; d[i]=d[i+1]=d[i+2]=v; }
      ctx.putImageData(id,0,0);
    }catch(e){}
    return c;
  }
  function scanFromFile(ev){
    var file = ev.target.files && ev.target.files[0];
    if(!file) return;
    if(!window.Tesseract){ alert("Chưa tải được bộ đọc chữ (OCR)."); return; }
    msg("Đang đọc chữ trên thẻ… (lần đầu hơi lâu vì tải bộ đọc, ~10–20s)");
    var img = new Image();
    img.onload = function(){
      var input;
      try{ input = preprocess(img); }catch(e){ input = file; }
      Tesseract.recognize(input, "vie")
        .then(function(r){
          if(fillFromOCR(r.data.text)){ recompute(); msg("Đã đọc từ ảnh ✓ — kiểm lại giúp, OCR có thể sai vài chữ (nhất là địa chỉ)."); }
          else { msg("Chưa bóc được thông tin. Chụp lại rõ & thẳng mặt trước thẻ, đủ sáng, hoặc nhập tay."); }
          URL.revokeObjectURL(img.src);
        })
        .catch(function(e){ msg("Không đọc được ảnh: "+(e&&e.message||e)); URL.revokeObjectURL(img.src); });
    };
    img.onerror = function(){ msg("Không mở được ảnh này."); URL.revokeObjectURL(img.src); };
    img.src = URL.createObjectURL(file);
    ev.target.value = "";
  }

  /* ---------- gắn sự kiện ---------- */
  document.addEventListener("DOMContentLoaded", function(){
    $("btnScan").onclick = startScan;
    $("btnStop").onclick = stopScan;
    $("btnParse").onclick = function(){ if(fillFromQr($("qrText").value)) recompute(); };
    $("cccdFile").addEventListener("change", scanFromFile);
    $("tien_vay").addEventListener("input", recompute);
    $("ngay_vay").addEventListener("change", recompute);
    $("btnPair").onclick = function(){
      var d=collect(); msg("Đang tạo…");
      Promise.resolve()
        .then(function(){return genOne("camdo",d);})
        .then(function(){return genOne("thuemay",d);})
        .then(function(){ msg("Đã tạo cầm đồ + thuê máy ✓ (xem mục tải về)"); })
        .catch(function(e){ msg("Lỗi: "+e.message); });
    };
    $("btnMuaban").onclick = function(){
      var d=collect(); msg("Đang tạo…");
      genOne("muaban",d).then(function(){ msg("Đã tạo giấy mua bán ✓"); }).catch(function(e){ msg("Lỗi: "+e.message); });
    };
  });
})();
