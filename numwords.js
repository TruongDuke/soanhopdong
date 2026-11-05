// Đổi số tiền (VNĐ) sang chữ tiếng Việt. Dùng được cả trong trình duyệt lẫn node.
(function(root){
  var D=["không","một","hai","ba","bốn","năm","sáu","bảy","tám","chín"];
  function doc3(n, full){ // đọc nhóm 3 chữ số; full=true nếu không phải nhóm đầu (ép đủ trăm/chục)
    var tram=Math.floor(n/100), chuc=Math.floor((n%100)/10), dv=n%10, s="";
    if(tram>0 || full){ s+=D[tram]+" trăm"; }
    if(chuc>1){
      s+=" "+D[chuc]+" mươi";
      if(dv===1) s+=" mốt"; else if(dv===5) s+=" lăm"; else if(dv>0) s+=" "+D[dv];
    } else if(chuc===1){
      s+=" mười";
      if(dv===5) s+=" lăm"; else if(dv>0) s+=" "+D[dv];
    } else { // chuc===0
      if(dv>0){ s+=(tram>0||full)?" lẻ "+D[dv]:D[dv]; }
    }
    return s.trim();
  }
  function docSoThanhChu(num){
    num=Math.round(Number(num)||0);
    if(num===0) return "Không đồng";
    var units=["","nghìn","triệu","tỷ"], groups=[], n=num;
    while(n>0){ groups.push(n%1000); n=Math.floor(n/1000); }
    // ghép từ nhóm cao xuống thấp
    var parts=[], L=groups.length;
    for(var i=L-1;i>=0;i--){
      var g=groups[i];
      if(g===0){ continue; }
      var isFirst=(i===L-1);
      var w=doc3(g, !isFirst); // nhóm sau phải đọc đủ "không trăm..."
      var unit=units[i%4];
      // xử lý tỷ lặp cho số rất lớn
      if(i>=4){ unit=units[i%4]; var tyCount=Math.floor(i/4); unit=unit+ " "+Array(tyCount+1).join("tỷ ").trim(); }
      parts.push((w+" "+unit).trim());
    }
    var s=parts.join(" ").replace(/\s+/g," ").trim();
    s=s.charAt(0).toUpperCase()+s.slice(1);
    return s+" đồng";
  }
  if(typeof module!=="undefined"&&module.exports) module.exports=docSoThanhChu;
  else root.docSoThanhChu=docSoThanhChu;
})(typeof window!=="undefined"?window:this);
