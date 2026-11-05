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

  /* ---------- đọc QR từ ảnh CCCD tải lên (jsQR + canvas, thử nhiều cỡ) ---------- */
  function qrAt(img, tw, th, sx, sy, sw, sh){
    var c = document.createElement("canvas"); c.width = tw; c.height = th;
    var ctx = c.getContext("2d"); ctx.drawImage(img, sx, sy, sw, sh, 0, 0, tw, th);
    var d = ctx.getImageData(0, 0, tw, th);
    var r = window.jsQR ? window.jsQR(d.data, tw, th, {inversionAttempts:"attemptBoth"}) : null;
    return r && r.data;
  }
  function decodeImage(img){
    var W = img.naturalWidth, H = img.naturalHeight, MAX = 1600;
    // 1) toàn ảnh, chuẩn hoá cạnh dài về <=1600 (ảnh to thu nhỏ, ảnh nhỏ phóng to)
    var s = MAX / Math.max(W, H);
    var tries = [[Math.round(W*s), Math.round(H*s), 0, 0, W, H]];
    // 2) cắt vùng góc trên phải (nơi có mã QR) rồi phóng cho nét
    var cw = Math.round(W*0.42), ch = Math.round(H*0.52), cx = W-cw, cy = 0;
    var cs = 1000 / Math.max(cw, ch);
    tries.push([Math.round(cw*cs), Math.round(ch*cs), cx, cy, cw, ch]);
    tries.push([Math.round(cw*cs*1.6), Math.round(ch*cs*1.6), cx, cy, cw, ch]);
    for (var i=0;i<tries.length;i++){ var t=tries[i]; try{ var v=qrAt(img,t[0],t[1],t[2],t[3],t[4],t[5]); if(v) return v; }catch(e){} }
    return null;
  }
  function scanFromFile(ev){
    var file = ev.target.files && ev.target.files[0];
    if(!file) return;
    if(!window.jsQR){ alert("Chưa tải được bộ đọc mã."); return; }
    msg("Đang đọc ảnh…");
    var img = new Image();
    img.onload = function(){
      var txt = decodeImage(img);
      if(txt){ if(fillFromQr(txt)){ recompute(); msg("Đã đọc CCCD từ ảnh ✓"); } }
      else { msg("Ảnh chưa đọc được mã QR. Chụp lại CẬN và RÕ ngay mã QR (góc trên phải thẻ), đủ sáng, không loá — hoặc bấm Quét trực tiếp cho chắc."); }
      URL.revokeObjectURL(img.src);
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
