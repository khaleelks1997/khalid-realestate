import { useState, useEffect, useCallback, useRef } from "react";
import { db } from "./firebase";
import { collection, getDocs, addDoc, updateDoc, deleteDoc, doc, onSnapshot } from "firebase/firestore";

const USERS = [
  { username: "Khalil", password: "Khalilks1997", role: "admin", displayName: "خليل" },
  { username: "emp", password: "emp1234", role: "employee", displayName: "موظف 1" },
  { username: "emp2", password: "emp2234", role: "employee", displayName: "موظف 2" },
];
const WA_NUMBER = "966568300022";
const PHONE = "0568300022";

const PROPERTY_TYPES = ["شقة","فيلا","محل تجاري","مكتب","استوديو","دوبلكس","أرض"];
const STATUS_OPTIONS  = ["متوفر","مؤجر","مباع","قريب الانتهاء","صيانة"];
const DEAL_TYPES      = ["إيجار","بيع","إيجار وبيع"];

const SC = {
  "متوفر":         { c:"#4ade80", bg:"#4ade8015" },
  "مؤجر":          { c:"#93c5fd", bg:"#93c5fd15" },
  "مباع":          { c:"#f87171", bg:"#f8717115" },
  "قريب الانتهاء": { c:"#fbbf24", bg:"#fbbf2415" },
  "صيانة":         { c:"#e879f9", bg:"#e879f915" },
};

const emptyForm = {
  name:"", address:"", type:"شقة", dealType:"إيجار",
  salePrice:"", rentPrice:"", minPrice:"",
  area:"", builtArea:"", ownerName:"", ownerPhone:"",
  furnished:false, aptCode:"", buildingCode:"",
  status:"متوفر", notes:"", images:[], mapUrl:"",
  marketingContractNo:"", adLicenseNo:"", refNo:"",
};

const today = () => new Date().toLocaleDateString("en-GB");

// ── Export to Excel ───────────────────────────────────────────────────────────
function exportToExcel(props) {
  const headers = ["الرقم المرجعي","اسم العقار","العنوان","النوع","الصفقة","الحالة","سعر البيع","الإيجار","المساحة","الغرف","الحمامات","مفروش","المالك","جوال المالك","رخصة إعلانية","عقد تسويق","ملاحظات"];
  const rows = props.map(p => [
    p.refNo||"",p.name||"",p.address||"",p.type||"",p.dealType||"",p.status||"",
    p.salePrice||"",p.rentPrice||"",p.area||"",p.rooms||"",p.bathrooms||"",
    p.furnished?"نعم":"لا",p.ownerName||"",p.ownerPhone||"",
    p.adLicenseNo||"",p.marketingContractNo||"",p.notes||""
  ]);
  const csv = [headers,...rows].map(r=>r.map(c=>`"${String(c).replace(/"/g,'""')}"`).join(",")).join("\n");
  const BOM = "\uFEFF";
  const blob = new Blob([BOM+csv],{type:"text/csv;charset=utf-8;"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href=url; a.download=`عقارات-${new Date().toLocaleDateString("en-GB").replace(/\//g,"-")}.csv`;
  a.click(); URL.revokeObjectURL(url);
}

// ── WA Icon ──────────────────────────────────────────────────────────────────
const WaIcon = ({size=14,color="#25d366"}) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill={color}>
    <path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/>
  </svg>
);

// ── Share Modal with Image Generator ─────────────────────────────────────────
function ShareModal({ p, onClose }) {
  const [copied, setCopied] = useState(false);
  const [generating, setGenerating] = useState(false);
  const cardRef = useRef();

  const text = `🏠 ${p.name}\n📍 ${p.address}\n🏷️ ${p.type} | ${p.dealType}${p.rentPrice?"\n💰 إيجار: "+Number(p.rentPrice).toLocaleString()+" ﷼/سنة":""}${p.salePrice?"\n💰 بيع: "+Number(p.salePrice).toLocaleString()+" ﷼":""}${p.area?"\n📐 المساحة: "+p.area+" م²":""}${p.furnished?"\n🛋️ مفروش":""}${p.adLicenseNo?"\n🏛️ رخصة إعلانية: "+p.adLicenseNo:""}\n📞 للتواصل: ${PHONE}\n💬 واتساب: wa.me/${WA_NUMBER}`;
  const waMsg = encodeURIComponent(text);
  const copy = () => { navigator.clipboard.writeText(text).then(()=>{ setCopied(true); setTimeout(()=>setCopied(false),2000); }); };
  const sc = SC[p.status]||SC["متوفر"];

  const downloadImage = async () => {
    setGenerating(true);
    try {
      // load html2canvas dynamically
      if (!window.html2canvas) {
        await new Promise((res,rej)=>{
          const s=document.createElement("script");
          s.src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
          s.onload=res; s.onerror=rej;
          document.head.appendChild(s);
        });
      }
      const canvas = await window.html2canvas(cardRef.current, {
        scale:2, useCORS:true, backgroundColor:null,
        logging:false, allowTaint:true
      });
      const a = document.createElement("a");
      a.download = `${p.name||"عقار"}.png`;
      a.href = canvas.toDataURL("image/png");
      a.click();
    } catch(e) { alert("حدث خطأ في توليد الصورة"); }
    setGenerating(false);
  };

  return (
    <div style={{position:"fixed",inset:0,background:"#000c",zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",padding:20,overflowY:"auto"}} onClick={onClose}>
      <div style={{background:"linear-gradient(160deg,#ffffff,#f5f8ff)",border:"1px solid rgba(74,158,255,.35)",borderRadius:20,padding:26,maxWidth:440,width:"100%"}} onClick={e=>e.stopPropagation()}>
        
        {/* Header */}
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <div style={{fontWeight:900,fontSize:16,color:"#fff"}}>📤 مشاركة العقار</div>
          <button onClick={onClose} style={{background:"#1e3a7a",border:"none",color:"#aaa",width:30,height:30,borderRadius:"50%",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
        </div>

        {/* Preview Card - this gets converted to image */}
        <div ref={cardRef} style={{
          background:"linear-gradient(135deg,#07103a 0%,#0e2563 60%,#1a4faa 100%)",
          borderRadius:18, padding:24, marginBottom:14,
          fontFamily:"'Cairo',sans-serif", direction:"rtl",
          border:`2px solid ${sc.c}44`, position:"relative", overflow:"hidden"
        }}>
          {/* Background pattern */}
          <div style={{position:"absolute",inset:0,backgroundImage:"linear-gradient(rgba(255,255,255,.03) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.03) 1px,transparent 1px)",backgroundSize:"24px 24px",pointerEvents:"none"}}/>
          
          {/* Logo + Title */}
          <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:18,position:"relative"}}>
            <img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1778958489/WhatsApp_Image_2026-05-16_at_9.59.47_PM_zhmw6y.jpg"
              alt="Logo" style={{width:52,height:52,borderRadius:12,objectFit:"cover",flexShrink:0}}/>
            <div>
              <div style={{fontWeight:900,fontSize:13,color:"#fff",lineHeight:1.3}}>مؤسسة خالد محمد عبدالغفور الشيخ</div>
              <div style={{fontSize:10,color:"rgba(255,255,255,.5)"}}>للخدمات العقارية</div>
            </div>
            <div style={{marginRight:"auto",background:sc.bg,border:`1px solid ${sc.c}44`,color:sc.c,borderRadius:20,padding:"3px 12px",fontSize:11,fontWeight:700}}>{p.status}</div>
          </div>

          {/* Property image if exists */}
          {p.images&&p.images[0]&&(
            <img src={p.images[0]} alt="" style={{width:"100%",height:160,objectFit:"cover",borderRadius:12,marginBottom:14}}/>
          )}

          {/* Property name */}
          <div style={{fontWeight:900,fontSize:20,color:"#fff",marginBottom:6,position:"relative"}}>{p.name}</div>
          <div style={{fontSize:13,color:"rgba(255,255,255,.6)",marginBottom:14,position:"relative"}}>📍 {p.address}</div>

          {/* Details grid */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14,position:"relative"}}>
            {p.rentPrice&&<div style={{background:"rgba(255,255,255,.07)",borderRadius:10,padding:"8px 12px"}}>
              <div style={{fontSize:10,color:"rgba(255,255,255,.4)"}}>الإيجار السنوي</div>
              <div style={{fontSize:15,fontWeight:900,color:"#4ade80"}}>{Number(p.rentPrice).toLocaleString()} ﷼</div>
            </div>}
            {p.salePrice&&<div style={{background:"rgba(255,255,255,.07)",borderRadius:10,padding:"8px 12px"}}>
              <div style={{fontSize:10,color:"rgba(255,255,255,.4)"}}>سعر البيع</div>
              <div style={{fontSize:15,fontWeight:900,color:"#fbbf24"}}>{Number(p.salePrice).toLocaleString()} ﷼</div>
            </div>}
            {p.area&&<div style={{background:"rgba(255,255,255,.07)",borderRadius:10,padding:"8px 12px"}}>
              <div style={{fontSize:10,color:"rgba(255,255,255,.4)"}}>المساحة</div>
              <div style={{fontSize:15,fontWeight:900,color:"#2a4d9b"}}>{p.area} م²</div>
            </div>}
            {p.rooms&&<div style={{background:"rgba(255,255,255,.07)",borderRadius:10,padding:"8px 12px"}}>
              <div style={{fontSize:10,color:"rgba(255,255,255,.4)"}}>الغرف</div>
              <div style={{fontSize:15,fontWeight:900,color:"#2a4d9b"}}>{p.rooms} غرف</div>
            </div>}
          </div>

          {/* Tags */}
          <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:14,position:"relative"}}>
            <span style={{background:"#1a4faa44",color:"#2a4d9b",fontSize:11,fontWeight:700,padding:"3px 10px",borderRadius:14}}>{p.type}</span>
            <span style={{background:"#1a4faa44",color:"#2a4d9b",fontSize:11,fontWeight:700,padding:"3px 10px",borderRadius:14}}>{p.dealType}</span>
            {p.furnished&&<span style={{background:"#fbbf2420",color:"#fbbf24",fontSize:11,fontWeight:700,padding:"3px 10px",borderRadius:14}}>🛋️ مفروش</span>}
          </div>

          {/* Contact */}
          <div style={{background:"rgba(255,255,255,.07)",borderRadius:12,padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center",position:"relative"}}>
            <div>
              <div style={{fontSize:10,color:"rgba(255,255,255,.4)",marginBottom:2}}>للتواصل</div>
              <div style={{fontSize:15,fontWeight:900,color:"#fff"}}>📞 {PHONE}</div>
            </div>
            <div style={{fontSize:11,color:"#25d366",fontWeight:700}}>💬 wa.me/{WA_NUMBER}</div>
          </div>

          {/* Ad license */}
          {p.adLicenseNo&&<div style={{marginTop:8,fontSize:10,color:"rgba(255,255,255,.3)",textAlign:"center",position:"relative"}}>🏛️ رخصة إعلانية: {p.adLicenseNo}</div>}
        </div>

        {/* Action buttons */}
        <div style={{display:"flex",flexDirection:"column",gap:9}}>
          <button onClick={downloadImage} disabled={generating} style={{display:"flex",alignItems:"center",gap:10,background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",border:"none",color:"#fff",borderRadius:12,padding:"11px 18px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:generating?"not-allowed":"pointer",justifyContent:"center",opacity:generating?.7:1}}>
            {generating?"⏳ جاري التوليد...":"🖼️ حفظ كصورة"}
          </button>
          <a href={`https://wa.me/?text=${waMsg}`} target="_blank" rel="noopener noreferrer" style={{display:"flex",alignItems:"center",gap:10,background:"#25d36620",border:"1px solid #25d36640",color:"#25d366",borderRadius:12,padding:"11px 18px",textDecoration:"none",fontWeight:700,fontSize:14,justifyContent:"center"}}><WaIcon size={18}/> مشاركة عبر واتساب</a>
          <button onClick={copy} style={{display:"flex",alignItems:"center",gap:10,background:copied?"#4ade8020":"#1a4faa22",border:`1px solid ${copied?"#4ade8040":"#2563c740"}`,color:copied?"#4ade80":"#93c5fd",borderRadius:12,padding:"11px 18px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",justifyContent:"center"}}>{copied?"✅ تم النسخ!":"📋 نسخ النص"}</button>
        </div>
      </div>
    </div>
  );
}

// ── Lightbox ──────────────────────────────────────────────────────────────────
function Lightbox({ images, startIndex, onClose }) {
  const [idx, setIdx] = useState(startIndex);
  const touchStart = useRef(null);
  const dragStart = useRef(null);

  useEffect(()=>{
    const h=e=>{ if(e.key==="Escape") onClose(); if(e.key==="ArrowLeft") setIdx(i=>(i+1)%images.length); if(e.key==="ArrowRight") setIdx(i=>(i-1+images.length)%images.length); };
    window.addEventListener("keydown",h); return ()=>window.removeEventListener("keydown",h);
  },[images,onClose]);

  const onTouchStart = e => { touchStart.current = e.touches[0].clientX; };
  const onTouchEnd = e => {
    if(touchStart.current===null) return;
    const diff = touchStart.current - e.changedTouches[0].clientX;
    if(Math.abs(diff) > 40) { diff > 0 ? setIdx(i=>(i+1)%images.length) : setIdx(i=>(i-1+images.length)%images.length); }
    touchStart.current = null;
  };
  const onMouseDown = e => { dragStart.current = e.clientX; };
  const onMouseUp = e => {
    if(dragStart.current===null) return;
    const diff = dragStart.current - e.clientX;
    if(Math.abs(diff) > 40) { diff > 0 ? setIdx(i=>(i+1)%images.length) : setIdx(i=>(i-1+images.length)%images.length); }
    dragStart.current = null;
  };

  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"#000e",zIndex:3000,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
      {/* Close button */}
      <button onClick={onClose} style={{position:"absolute",top:16,right:16,width:36,height:36,background:"#ef4444",border:"none",borderRadius:"50%",color:"#fff",fontSize:18,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",zIndex:10}}>×</button>

      {/* Counter */}
      <div style={{position:"absolute",top:20,left:"50%",transform:"translateX(-50%)",background:"rgba(0,0,0,.6)",color:"#fff",padding:"4px 14px",borderRadius:20,fontSize:13,fontWeight:700,zIndex:10}}>{idx+1} / {images.length}</div>

      {/* Image */}
      <div style={{position:"relative",userSelect:"none",width:"100%",display:"flex",alignItems:"center",justifyContent:"center",padding:"0 56px",boxSizing:"border-box"}}
        onClick={e=>e.stopPropagation()}
        onTouchStart={onTouchStart} onTouchEnd={onTouchEnd}
        onMouseDown={onMouseDown} onMouseUp={onMouseUp}
      >
        <img src={images[idx]} alt="" style={{maxWidth:"100%",maxHeight:"70vh",borderRadius:14,objectFit:"contain",display:"block",cursor:"grab"}}/>
      </div>

      {/* Arrows - always inside screen */}
      {images.length>1&&<>
        <button onClick={e=>{e.stopPropagation();setIdx(i=>(i-1+images.length)%images.length);}} style={{position:"absolute",right:8,top:"50%",transform:"translateY(-50%)",background:"rgba(30,58,122,.9)",border:"1px solid rgba(74,158,255,.4)",color:"#fff",width:44,height:44,borderRadius:"50%",fontSize:24,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",zIndex:10}}>‹</button>
        <button onClick={e=>{e.stopPropagation();setIdx(i=>(i+1)%images.length);}} style={{position:"absolute",left:8,top:"50%",transform:"translateY(-50%)",background:"rgba(30,58,122,.9)",border:"1px solid rgba(74,158,255,.4)",color:"#fff",width:44,height:44,borderRadius:"50%",fontSize:24,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",zIndex:10}}>›</button>
      </>}

      {/* Thumbnails */}
      {images.length>1&&(<div style={{display:"flex",gap:6,marginTop:14,overflowX:"auto",maxWidth:"88vw",padding:"0 4px"}} onClick={e=>e.stopPropagation()}>{images.map((img,i)=>(<img key={i} src={img} onClick={()=>setIdx(i)} alt="" style={{width:54,height:54,objectFit:"cover",borderRadius:8,cursor:"pointer",flexShrink:0,border:i===idx?"2px solid #4a9eff":"2px solid transparent",opacity:i===idx?1:.5,transition:"all .2s"}}/>))}</div>)}
    </div>
  );
}

// ── Image Uploader ────────────────────────────────────────────────────────────
// ── Cloudinary Upload + Compression ──────────────────────────────────────────
const CLOUD_NAME = "dumtp0krl";
const UPLOAD_PRESET = "khalid_realestate";

async function compressImage(file, maxWidth=1200, quality=0.8) {
  return new Promise((resolve) => {
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => {
      const canvas = document.createElement("canvas");
      let w = img.width, h = img.height;
      if (w > maxWidth) { h = Math.round(h * maxWidth / w); w = maxWidth; }
      canvas.width = w; canvas.height = h;
      canvas.getContext("2d").drawImage(img, 0, 0, w, h);
      canvas.toBlob(blob => { URL.revokeObjectURL(url); resolve(blob); }, "image/jpeg", quality);
    };
    img.src = url;
  });
}

async function uploadToCloudinary(file) {
  const compressed = await compressImage(file);
  const fd = new FormData();
  fd.append("file", compressed, file.name);
  fd.append("upload_preset", UPLOAD_PRESET);
  fd.append("folder", "properties");
  const res = await fetch(`https://api.cloudinary.com/v1_1/${CLOUD_NAME}/image/upload`, { method:"POST", body:fd });
  const data = await res.json();
  if (data.secure_url) return data.secure_url;
  throw new Error(data.error?.message || "Upload failed");
}

function ImageUploader({ images, onChange }) {
  const ref = useRef();
  const [drag, setDrag] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [currentFile, setCurrentFile] = useState("");

  const process = async (files) => {
    const valid = Array.from(files).filter(f => f.type.startsWith("image/"));
    if (!valid.length) return;
    setUploading(true);
    for (let i = 0; i < valid.length; i++) {
      setCurrentFile(`${i+1}/${valid.length}`);
      setProgress(Math.round((i/valid.length)*100));
      try {
        const url = await uploadToCloudinary(valid[i]);
        onChange(p => [...p, url]);
      } catch(e) { alert("خطأ في رفع الصورة: " + e.message); }
    }
    setUploading(false); setCurrentFile(""); setProgress(0);
  };

  return (
    <div>
      <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);process(e.dataTransfer.files);}} onClick={()=>!uploading&&ref.current.click()}
        style={{border:`2px dashed ${drag?"#60a5fa":"#1e3a7a"}`,borderRadius:12,padding:"16px",textAlign:"center",cursor:uploading?"not-allowed":"pointer",background:drag?"#60a5fa0d":"transparent",transition:"all .2s",marginBottom:10}}>
        <div style={{fontSize:24,marginBottom:4}}>{uploading?"⏳":"📸"}</div>
        {uploading ? (
          <div>
            <div style={{color:"#60a5fa",fontSize:13,fontWeight:700,marginBottom:4}}>جاري الضغط والرفع... {progress}% ({currentFile})</div>
            <div style={{height:6,background:"#1e3a7a",borderRadius:4}}><div style={{height:"100%",background:"linear-gradient(90deg,#1a4faa,#60a5fa)",borderRadius:4,width:progress+"%",transition:"width .3s"}}/></div>
          </div>
        ) : (
          <div style={{color:"#5a6a90",fontSize:13}}>اسحب الصور أو <span style={{color:"#60a5fa",fontWeight:700}}>اضغط للاختيار</span><br/><span style={{fontSize:11,color:"#5a6a90"}}>يتم ضغط الصور تلقائياً ⚡</span></div>
        )}
        <input ref={ref} type="file" accept="image/*" multiple style={{display:"none"}} onChange={e=>process(e.target.files)}/>
      </div>
      {images.length>0&&(<div style={{display:"flex",flexWrap:"wrap",gap:8}}>{images.map((img,i)=>(<div key={i} style={{position:"relative",width:72,height:72}}><img src={img} alt="" style={{width:"100%",height:"100%",objectFit:"cover",borderRadius:9,border:"2px solid #1e3a7a"}}/><button onClick={()=>onChange(images.filter((_,j)=>j!==i))} style={{position:"absolute",top:-5,right:-5,width:18,height:18,background:"#ef4444",border:"none",borderRadius:"50%",color:"#fff",fontSize:11,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>{i===0&&<div style={{position:"absolute",bottom:0,left:0,right:0,background:"#1a4faa",fontSize:9,color:"#fff",textAlign:"center",borderRadius:"0 0 7px 7px",fontWeight:900}}>رئيسية</div>}</div>))}</div>)}
    </div>
  );
}

function CodeInput({ label, value, onChange }) {
  const [show, setShow] = useState(false);
  return (<div><div style={{fontSize:11,color:"#5a6a90",marginBottom:5,fontWeight:600}}>{label}</div><div style={{position:"relative"}}><input type={show?"text":"password"} value={value} onChange={onChange} placeholder="—" style={{width:"100%",boxSizing:"border-box",background:"#f0f4fc",border:"1px solid #fbbf2444",borderRadius:10,padding:"9px 36px 9px 12px",color:"#fbbf24",fontFamily:"inherit",fontSize:13,fontWeight:700}}/><button onClick={()=>setShow(s=>!s)} style={{position:"absolute",top:"50%",left:9,transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:14,padding:0}}>{show?"🙈":"👁️"}</button></div></div>);
}

// ── NAVBAR ────────────────────────────────────────────────────────────────────
function Navbar({ page, setPage, isAdmin, onLoginClick, onLogout, lang, setLang, scrolled, darkMode, setDarkMode, T, userRole }) {
  const isEn = lang==="en";
  const [menuOpen, setMenuOpen] = useState(false);
  const navItems = isEn
    ? [["home","🏠","Home"],["properties","🏘️","Properties"],["services","✦","Services"],["about","ℹ️","About"]]
    : [["home","🏠","الرئيسية"],["properties","🏘️","العقارات"],["services","✦","خدماتنا"],["about","ℹ️","عن المؤسسة"]];

  const go = (p) => { setPage(p); setMenuOpen(false); };

  const slideStyle = {
    position:"fixed", top:0, right: menuOpen ? 0 : -320,
    width:300, height:"100vh",
    background: darkMode ? "#1a2d6b" : "#ffffff",
    zIndex:500, transition:"right .3s cubic-bezier(.4,0,.2,1)",
    boxShadow: menuOpen ? "-4px 0 30px rgba(13,31,82,.25)" : "none",
    display:"flex", flexDirection:"column",
  };

  return (
    <>
      {/* Overlay */}
      {menuOpen && <div onClick={()=>setMenuOpen(false)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.45)",zIndex:499,backdropFilter:"blur(2px)"}}/>}

      {/* Slide menu */}
      <div style={slideStyle}>
        {/* Header */}
        <div style={{background:"#1e3a7a",padding:"14px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",borderBottom:"2px solid rgba(74,158,255,.25)",flexShrink:0}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:40,height:40,borderRadius:9,background:"white",overflow:"hidden",flexShrink:0}}>
              <img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1778958489/WhatsApp_Image_2026-05-16_at_9.59.47_PM_zhmw6y.jpg" alt="" style={{width:"100%",height:"100%",objectFit:"contain"}}/>
            </div>
            <div>
              <div style={{fontSize:11,fontWeight:900,color:"white",lineHeight:1.2}}>{isEn?"Khalid Al-Shaikh Est.":"مؤسسة خالد محمد"}</div>
              <div style={{fontSize:9,color:"#7ab8ff"}}>{isEn?"Real Estate":"للخدمات العقارية"}</div>
            </div>
          </div>
          <button onClick={()=>setMenuOpen(false)} style={{background:"rgba(255,255,255,.12)",border:"none",color:"white",width:32,height:32,borderRadius:"50%",cursor:"pointer",fontSize:18,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
        </div>

        {/* Body */}
        <div style={{flex:1,overflowY:"auto",padding:"10px"}}>
          {/* Nav section */}
          <div style={{fontSize:10,fontWeight:700,color:darkMode?"rgba(255,255,255,.3)":"#8899bb",padding:"6px 8px 4px",marginBottom:4}}>{isEn?"MENU":"القائمة"}</div>
          {navItems.map(([p,icon,label])=>(
            <button key={p} onClick={()=>go(p)} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 12px",borderRadius:12,border:"none",background:page===p?(darkMode?"rgba(74,158,255,.15)":"rgba(74,158,255,.08)"):"none",color:page===p?(darkMode?"#7ab8ff":"#1e3a7a"):(darkMode?"rgba(255,255,255,.6)":"#5a6a90"),fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",width:"100%",textAlign:"right",marginBottom:2,transition:"all .15s"}}>
              <div style={{width:36,height:36,borderRadius:10,background:darkMode?"rgba(255,255,255,.06)":"#edf1fb",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>{icon}</div>
              <span>{label}</span>
            </button>
          ))}
          {isAdmin&&(
            <>
              <button onClick={()=>go("clients")} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 12px",borderRadius:12,border:"none",background:page==="clients"?(darkMode?"rgba(74,158,255,.15)":"rgba(74,158,255,.08)"):"none",color:page==="clients"?(darkMode?"#7ab8ff":"#1e3a7a"):(darkMode?"rgba(255,255,255,.6)":"#5a6a90"),fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",width:"100%",textAlign:"right",marginBottom:2}}>
                <div style={{width:36,height:36,borderRadius:10,background:darkMode?"rgba(255,255,255,.06)":"#edf1fb",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>👥</div>
                <span>{isEn?"Clients":"سجل العملاء"}</span>
              </button>
              {userRole==="admin"&&<button onClick={()=>go("providers")} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 12px",borderRadius:12,border:"none",background:page==="providers"?(darkMode?"rgba(74,158,255,.15)":"rgba(74,158,255,.08)"):"none",color:page==="providers"?(darkMode?"#7ab8ff":"#1e3a7a"):(darkMode?"rgba(255,255,255,.6)":"#5a6a90"),fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",width:"100%",textAlign:"right",marginBottom:2}}>
                <div style={{width:36,height:36,borderRadius:10,background:darkMode?"rgba(255,255,255,.06)":"#edf1fb",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>🔧</div>
                <span>{isEn?"Providers":"مزودو الخدمات"}</span>
              </button>}
            </>
          )}

          {/* Divider */}
          <div style={{height:1,background:darkMode?"rgba(255,255,255,.07)":"rgba(74,158,255,.1)",margin:"10px 4px"}}/>

          {/* Settings section */}
          <div style={{fontSize:10,fontWeight:700,color:darkMode?"rgba(255,255,255,.3)":"#8899bb",padding:"6px 8px 4px",marginBottom:4}}>{isEn?"SETTINGS":"الإعدادات"}</div>
          <button onClick={()=>{setDarkMode(d=>!d);}} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 12px",borderRadius:12,border:"none",background:"none",color:darkMode?"rgba(255,255,255,.6)":"#5a6a90",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",width:"100%",textAlign:"right",marginBottom:2}}>
            <div style={{width:36,height:36,borderRadius:10,background:darkMode?"rgba(255,255,255,.06)":"#edf1fb",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>{darkMode?"☀️":"🌙"}</div>
            <span>{darkMode?(isEn?"Light Mode":"الوضع الفاتح"):(isEn?"Dark Mode":"الوضع الداكن")}</span>
          </button>
          <button onClick={()=>{setLang(isEn?"ar":"en");}} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 12px",borderRadius:12,border:"none",background:"none",color:darkMode?"rgba(255,255,255,.6)":"#5a6a90",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",width:"100%",textAlign:"right",marginBottom:2}}>
            <div style={{width:36,height:36,borderRadius:10,background:darkMode?"rgba(255,255,255,.06)":"#edf1fb",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>{isEn?"🇸🇦":"🇬🇧"}</div>
            <span>{isEn?"عربي":"English"}</span>
          </button>
        </div>

        {/* Footer */}
        <div style={{padding:"12px",borderTop:`1px solid ${darkMode?"rgba(255,255,255,.07)":"rgba(74,158,255,.1)"}`,flexShrink:0}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
            <a href={`tel:${PHONE}`} style={{display:"flex",alignItems:"center",justifyContent:"center",gap:6,background:"#1e3a7a",color:"#fff",borderRadius:11,padding:"11px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:12}}>📞 {PHONE}</a>
            <a href={`https://wa.me/${WA_NUMBER}`} target="_blank" rel="noopener noreferrer" style={{display:"flex",alignItems:"center",justifyContent:"center",gap:6,background:"rgba(37,211,102,.1)",border:"1px solid rgba(37,211,102,.25)",color:"#25d366",borderRadius:11,padding:"11px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:12}}><WaIcon size={13}/> WhatsApp</a>
          </div>
          {isAdmin
            ? <button onClick={()=>{onLogout();setMenuOpen(false);}} style={{width:"100%",background:"#ef444418",border:"1px solid #ef444430",color:"#f87171",borderRadius:11,padding:"11px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>🔓 {isEn?"Logout":"تسجيل خروج"}</button>
            : <button onClick={()=>{onLoginClick();setMenuOpen(false);}} style={{width:"100%",background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",border:"none",color:"#fff",borderRadius:11,padding:"11px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>🔐 {isEn?"Admin Login":"دخول الإدارة"}</button>
          }
        </div>
      </div>

      {/* Topbar */}
      <div style={{background:"#0a1538",padding:"7px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:11,color:"rgba(255,255,255,.4)"}}>
        <div style={{display:"flex",gap:20}}>
          <span>مرخصون من الهيئة العامة للعقار 🏛️</span>
          <span style={{color:"rgba(255,255,255,.2)"}}>|</span>
          <span>📍 المنطقة الشرقية، المملكة العربية السعودية</span>
        </div>
        <div style={{display:"flex",gap:16}}>
          <a href={`tel:${PHONE}`} style={{color:"rgba(255,255,255,.4)",textDecoration:"none",transition:"color .2s"}}>📞 {PHONE}</a>
          <a href={`https://wa.me/${WA_NUMBER}`} target="_blank" rel="noopener noreferrer" style={{color:"rgba(37,211,102,.6)",textDecoration:"none"}}>💬 WhatsApp</a>
        </div>
      </div>

      {/* Navbar bar */}
      <div style={{position:"fixed",top:0,left:0,right:0,zIndex:200,background:"#1e3a7a",backdropFilter:"blur(18px)",borderBottom:"2px solid rgba(74,158,255,.3)",transition:"background .3s"}}>
        <div style={{maxWidth:1200,margin:"0 auto",padding:"0 16px",display:"flex",alignItems:"center",justifyContent:"space-between",height:64}}>

          {/* Brand */}
          <button onClick={()=>go("home")} style={{display:"flex",alignItems:"center",gap:10,background:"none",border:"none",cursor:"pointer",padding:0,flexShrink:0}}>
            <div style={{width:44,height:44,borderRadius:10,overflow:"hidden",flexShrink:0,boxShadow:"0 2px 10px rgba(74,158,255,.3)"}}><img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1778958489/WhatsApp_Image_2026-05-16_at_9.59.47_PM_zhmw6y.jpg" alt="Logo" style={{width:"100%",height:"100%",objectFit:"cover"}}/></div>
            <div className="brand-text" style={{textAlign:"right"}}>
              <div style={{fontWeight:900,fontSize:11,color:"white",fontFamily:"'Cairo',sans-serif",lineHeight:1.2,whiteSpace:"nowrap"}}>{isEn?"Khalid Al-Shaikh Est.":"مؤسسة خالد محمد عبدالغفور الشيخ"}</div>
              <div style={{fontSize:9,color:"#7ab8ff",fontFamily:"'Cairo',sans-serif"}}>{isEn?"Real Estate Services":"للخدمات العقارية"}</div>
            </div>
          </button>

          {/* Desktop nav */}
          <div className="desktop-nav" style={{display:"flex",gap:2}}>
            {navItems.map(([p,icon,label])=>(<button key={p} onClick={()=>go(p)} style={{background:page===p?"rgba(74,158,255,.18)":"none",border:"none",color:page===p?"#7ab8ff":"rgba(255,255,255,.55)",borderRadius:8,padding:"6px 10px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,cursor:"pointer",whiteSpace:"nowrap"}}>{icon} {label}</button>))}
          </div>

          {/* Right actions */}
          <div style={{display:"flex",alignItems:"center",gap:6,flexShrink:0}}>
            <div className="desktop-actions" style={{display:"flex",alignItems:"center",gap:6}}>
              <button onClick={()=>setDarkMode(d=>!d)} style={{background:"rgba(255,255,255,.08)",border:"1px solid rgba(255,255,255,.14)",color:"white",borderRadius:8,padding:"5px 9px",cursor:"pointer",fontSize:13}}>{darkMode?"☀️":"🌙"}</button>
              <button onClick={()=>setLang(isEn?"ar":"en")} style={{background:"rgba(255,255,255,.08)",border:"1px solid rgba(255,255,255,.14)",color:"white",borderRadius:8,padding:"5px 9px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,cursor:"pointer"}}>{isEn?"🇸🇦":"🇬🇧 EN"}</button>
              <a href={`tel:${PHONE}`} style={{background:"#4a9eff",color:"#fff",borderRadius:8,padding:"7px 12px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,cursor:"pointer",textDecoration:"none",whiteSpace:"nowrap",boxShadow:"0 2px 10px rgba(74,158,255,.35)"}}>📞 {PHONE}</a>
              <a href={`https://wa.me/${WA_NUMBER}`} target="_blank" rel="noopener noreferrer" style={{background:"rgba(37,211,102,.15)",border:"1px solid rgba(37,211,102,.3)",color:"#25d366",borderRadius:8,padding:"6px 9px",textDecoration:"none",display:"flex",alignItems:"center"}}><WaIcon size={12}/></a>
              {isAdmin
                ? <button onClick={onLogout} style={{background:"#ef444422",border:"1px solid #ef444444",color:"#f87171",borderRadius:8,padding:"6px 10px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,cursor:"pointer"}}>🔓</button>
                : <button onClick={onLoginClick} style={{background:"rgba(255,255,255,.1)",border:"1px solid rgba(255,255,255,.2)",color:"#fff",borderRadius:8,padding:"7px 12px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,cursor:"pointer",whiteSpace:"nowrap"}}>🔐 {isEn?"Admin":"الإدارة"}</button>
              }
            </div>
            {/* Hamburger */}
            <button onClick={()=>setMenuOpen(s=>!s)} style={{background:"rgba(255,255,255,.08)",border:"1px solid rgba(255,255,255,.15)",color:"white",borderRadius:8,width:40,height:40,cursor:"pointer",fontSize:18,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
              {menuOpen?"✕":"☰"}
            </button>
          </div>
        </div>

        {isAdmin&&(<div style={{background:"rgba(74,158,255,.1)",borderTop:"1px solid rgba(74,158,255,.2)",padding:"5px 20px",display:"flex",alignItems:"center",justifyContent:"center",gap:10}}><span style={{fontSize:11,color:"#7ab8ff",fontWeight:600}}>{userRole==="admin"?"👑 وضع المدير":"👤 وضع الموظف"}</span></div>)}
      </div>
    </>
  );
}

// ── Login Modal ───────────────────────────────────────────────────────────────
function LoginModal({ onSuccess, onClose, lang }) {
  const isEn=lang==="en";
  const [un,setUn]=useState(""); const [pw,setPw]=useState(""); const [show,setShow]=useState(false); const [err,setErr]=useState(""); const [loading,setLoading]=useState(false);
  const login=()=>{ setLoading(true); setTimeout(()=>{ const ok=USERS.find(u=>u.username===un&&u.password===pw); if(ok) onSuccess(ok); else { setErr(isEn?"Incorrect credentials":"بيانات الدخول غير صحيحة"); setLoading(false); } },600); };
  const IST={width:"100%",boxSizing:"border-box",background:"#f0f4fc",border:"1px solid rgba(74,158,255,.2)",borderRadius:10,padding:"10px 13px",color:"#1e3a7a",fontFamily:"'Cairo',sans-serif",fontSize:14};
  return (
    <div style={{position:"fixed",inset:0,background:"#000c",zIndex:1500,display:"flex",alignItems:"center",justifyContent:"center",padding:20}} onClick={onClose}>
      <div style={{background:"linear-gradient(160deg,#ffffff,#f5f8ff)",border:"2px solid #2563c7",borderRadius:22,padding:32,maxWidth:380,width:"100%",boxShadow:"0 24px 60px #1a4faa33"}} onClick={e=>e.stopPropagation()}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{width:60,height:60,borderRadius:16,background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",margin:"0 auto 14px",display:"flex",alignItems:"center",justifyContent:"center",fontSize:28}}>🏛️</div>
          <div style={{fontWeight:900,fontSize:18,color:"#fff",marginBottom:4}}>{isEn?"Admin Login":"دخول الإدارة"}</div>
          <div style={{fontSize:11,color:"#5a6a90"}}>Khalid M. A. Ghafour Al-Shaikh Est.</div>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:14,marginBottom:16}}>
          <div><div style={{fontSize:12,color:"#5a6a90",marginBottom:6,fontWeight:600}}>{isEn?"Username":"اسم المستخدم"}</div><input value={un} onChange={e=>setUn(e.target.value)} onKeyDown={e=>e.key==="Enter"&&login()} style={IST}/></div>
          <div><div style={{fontSize:12,color:"#5a6a90",marginBottom:6,fontWeight:600}}>{isEn?"Password":"كلمة السر"}</div><div style={{position:"relative"}}><input type={show?"text":"password"} value={pw} onChange={e=>setPw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&login()} style={{...IST,paddingLeft:42}}/><button onClick={()=>setShow(s=>!s)} style={{position:"absolute",top:"50%",left:12,transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:16,padding:0}}>{show?"🙈":"👁️"}</button></div></div>
        </div>
        {err&&<div style={{background:"#ef444418",border:"1px solid #ef444430",color:"#f87171",borderRadius:10,padding:"9px 14px",fontSize:13,fontWeight:600,marginBottom:14,textAlign:"center"}}>⚠️ {err}</div>}
        <button onClick={login} disabled={loading} style={{width:"100%",background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",color:"#fff",border:"none",borderRadius:12,padding:"13px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:15,cursor:"pointer",opacity:loading?.7:1}}>{loading?(isEn?"Verifying...":"⏳ جاري التحقق..."):(isEn?"Login 🔓":"دخول 🔓")}</button>
      </div>
    </div>
  );
}

// ── Prop Form ─────────────────────────────────────────────────────────────────
function PropForm({ form, setForm, onSave, onClose, editId, T, userRole }) {
  const f=key=>e=>setForm(p=>({...p,[key]:e.target.value}));
  const IST={width:"100%",boxSizing:"border-box",background:"#f0f4fc",border:"1px solid rgba(74,158,255,.2)",borderRadius:10,padding:"9px 12px",color:"#1e3a7a",fontFamily:"'Cairo',sans-serif",fontSize:13};
  const Lbl=({c})=><div style={{fontSize:11,color:"#5a6a90",marginBottom:5,fontWeight:600}}>{c}</div>;
  const Sec=({c})=><div style={{fontSize:11,color:"#60a5fa",fontWeight:700,marginBottom:10,padding:"6px 12px",background:"#1a4faa18",borderRadius:9,border:"1px solid #1a4faa33"}}>{c}</div>;
  return (
    <div style={{position:"fixed",inset:0,background:"#000b",zIndex:600,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={onClose}>
      <div style={{background:"linear-gradient(160deg,#ffffff,#f5f8ff)",border:"1px solid #1a4faa",borderRadius:22,padding:24,maxWidth:600,width:"100%",maxHeight:"93vh",overflowY:"auto"}} onClick={e=>e.stopPropagation()}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20,paddingBottom:14,borderBottom:"1px solid #1e3a7a"}}>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <div style={{width:36,height:36,borderRadius:10,background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>🏢</div>
            <div><div style={{fontWeight:900,fontSize:14,color:"#1e3a7a"}}>{editId?"تعديل العقار":"إضافة عقار جديد"}</div><div style={{fontSize:10,color:"#5a6a90"}}>مؤسسة خالد محمد عبدالغفور الشيخ</div></div>
          </div>
          <button onClick={onClose} style={{background:"#1e3a7a",border:"none",color:"#aaa",width:30,height:30,borderRadius:"50%",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
        </div>

        <Sec c="📋 معلومات أساسية"/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:11,marginBottom:16}}>
          <div style={{gridColumn:"1/-1"}}><Lbl c="اسم العقار *"/><input value={form.name} onChange={f("name")} style={IST}/></div>
          <div style={{gridColumn:"1/-1"}}><Lbl c="العنوان"/><input value={form.address} onChange={f("address")} style={IST}/></div>
          {[["نوع العقار","type",PROPERTY_TYPES],["نوع الصفقة","dealType",DEAL_TYPES],["الحالة","status",STATUS_OPTIONS]].map(([lbl,key,opts])=>(<div key={key}><Lbl c={lbl}/><select value={form[key]} onChange={f(key)} style={IST}>{opts.map(o=><option key={o}>{o}</option>)}</select></div>))}
          <div><Lbl c="مفروش؟"/><div style={{display:"flex",gap:7}}>{["نعم","لا"].map(opt=>(<button key={opt} onClick={()=>setForm(p=>({...p,furnished:opt==="نعم"}))} style={{flex:1,padding:"8px",borderRadius:9,border:`1px solid ${(opt==="نعم")===form.furnished?"#2563c7":"#1e3a7a"}`,background:(opt==="نعم")===form.furnished?"#1a4faa":"#071840",color:(opt==="نعم")===form.furnished?"#fff":"#4a6fa5",cursor:"pointer",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:12}}>{opt==="نعم"?"🛋️ نعم":"❌ لا"}</button>))}</div></div>
          <div><Lbl c="الرقم المرجعي"/><input value={form.refNo} onChange={f("refNo")} style={IST} placeholder="REF-001"/></div>
        </div>

        <Sec c="🏛️ الأرقام الرسمية"/>
        <div style={{background:"#ef444408",border:"1px solid #ef444420",borderRadius:9,padding:"8px 12px",marginBottom:11,fontSize:11,color:"#f87171"}}>⚠️ رقم الترخيص الإعلاني إلزامي — اشتراط الهيئة العامة للعقار</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:11,marginBottom:16}}>
          <div><Lbl c="رقم الترخيص الإعلاني *"/><input value={form.adLicenseNo} onChange={f("adLicenseNo")} style={{...IST,borderColor:form.adLicenseNo?"#4ade8040":"#ef444440",color:form.adLicenseNo?"#4ade80":"#f87171"}} placeholder="1010123456"/></div>
          <div><Lbl c="رقم عقد التسويق"/><input value={form.marketingContractNo} onChange={f("marketingContractNo")} style={IST} placeholder="MC-2024-001"/></div>
        </div>

        <Sec c="💰 الأسعار"/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:11,marginBottom:16}}>
          {[["سعر البيع (﷼)","salePrice"],["قيمة الإيجار السنوي (﷼)","rentPrice"],["الحد الأدنى (﷼)","minPrice"]].map(([lbl,key])=>(<div key={key}><Lbl c={lbl}/><input type="number" value={form[key]} onChange={f(key)} style={IST}/></div>))}
        </div>

        <Sec c="📐 المساحات"/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:11,marginBottom:16}}>
          {[["المساحة الصافية (م²)","area"],["المسطح البنائي (م²)","builtArea"]].map(([lbl,key])=>(<div key={key}><Lbl c={lbl}/><input type="number" value={form[key]} onChange={f(key)} style={IST}/></div>))}
        </div>

        <Sec c="👤 المالك"/>
        {userRole==="admin" ? (
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:11,marginBottom:16}}>
            <div><Lbl c="اسم المالك"/><input value={form.ownerName} onChange={f("ownerName")} style={IST}/></div>
            <div><Lbl c="رقم الجوال"/><input value={form.ownerPhone} onChange={f("ownerPhone")} style={IST}/></div>
          </div>
        ) : (
          <div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:12,padding:"12px 14px",marginBottom:16,display:"flex",alignItems:"center",gap:10}}>
            <span style={{fontSize:20}}>🔒</span>
            <div><div style={{fontSize:13,color:"#ef4444",fontWeight:700}}>بيانات المالك محجوبة</div><div style={{fontSize:11,color:"#f87171",marginTop:2}}>فقط المدير يستطيع مشاهدة وتعديل بيانات المالك</div></div>
          </div>
        )}

        <Sec c="🔐 رموز الدخول"/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:11,marginBottom:16}}>
          <CodeInput label="رمز دخول الشقة" value={form.aptCode} onChange={f("aptCode")}/>
          <CodeInput label="رمز دخول المبنى" value={form.buildingCode} onChange={f("buildingCode")}/>
        </div>

        <Sec c="📍 موقع العقار"/>
        <div style={{marginBottom:14}}>
          <Lbl c="رابط قوقل ماب"/>
          <input value={form.mapUrl} onChange={f("mapUrl")} placeholder="https://maps.google.com/?q=..." style={{...IST,borderColor:"#16a34a33",color:"#4ade80"}}/>
          <div style={{fontSize:10,color:"#5a6a90",marginTop:5}}>💡 قوقل ماب → ابحث عن الموقع → مشاركة → نسخ الرابط</div>
        </div>

        <div style={{marginBottom:14}}><Lbl c="📝 ملاحظات"/><textarea value={form.notes} onChange={f("notes")} rows={5} placeholder="كل سطر يظهر منفصلاً..." style={{...IST,resize:"vertical",whiteSpace:"pre-wrap",lineHeight:1.7}}/></div>

        <Sec c="🖼️ صور العقار"/>
        <div style={{marginBottom:20}}><ImageUploader images={form.images||[]} onChange={imgs=>setForm(p=>({...p,images:typeof imgs==="function"?imgs(p.images||[]):imgs}))}/></div>

        <div style={{display:"flex",gap:10}}>
          <button onClick={onSave} style={{flex:1,background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",color:"#fff",border:"none",borderRadius:11,padding:"12px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer"}}>{editId?"حفظ التعديلات":"إضافة العقار"}</button>
          <button onClick={onClose} style={{background:"#f0f4fc",border:"1px solid rgba(74,158,255,.2)",color:"#5a6a90",borderRadius:11,padding:"12px 18px",fontFamily:"'Cairo',sans-serif",cursor:"pointer",fontWeight:600}}>إلغاء</button>
        </div>
      </div>
    </div>
  );
}

// ── Public Card ───────────────────────────────────────────────────────────────
// ── Property Detail Modal ─────────────────────────────────────────────────────
function PropertyModal({ p, onClose, setLightbox, onShare, lang }) {
  const isEn=lang==="en";
  const imgs=p.images||[]; const sc=SC[p.status]||SC["متوفر"];
  const waMsg=encodeURIComponent(`مرحباً، أود الاستفسار عن ${p.name} - ${p.address}${p.rentPrice?" | "+Number(p.rentPrice).toLocaleString()+" ﷼/سنة":""}${p.salePrice?" | "+Number(p.salePrice).toLocaleString()+" ﷼":""}`);
  const statusEn={"متوفر":"Available","مؤجر":"Rented","مباع":"Sold","قريب الانتهاء":"Expiring","صيانة":"Maintenance"};

  useEffect(()=>{ document.body.style.overflow="hidden"; return()=>{ document.body.style.overflow=""; }; },[]);

  return (
    <div style={{position:"fixed",inset:0,background:"#000d",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16,overflowY:"auto"}} onClick={onClose}>
      <div style={{background:"linear-gradient(160deg,#ffffff,#f5f8ff)",border:`1px solid ${sc.c}33`,borderRadius:22,maxWidth:600,width:"100%",overflow:"hidden",maxHeight:"90vh",overflowY:"auto"}} onClick={e=>e.stopPropagation()}>

        {/* Header image */}
        <div style={{position:"relative",height:240,background:"#edf1fb",cursor:imgs.length>0?"pointer":"default"}} onClick={()=>imgs.length>0&&setLightbox({images:imgs,idx:0})}>
          {imgs.length>0?(<><img src={imgs[0]} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/><div style={{position:"absolute",inset:0,background:"linear-gradient(to top,#071840cc,transparent 60%)",pointerEvents:"none"}}/></>):(<div style={{width:"100%",height:"100%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:52,color:"#1e3a7a"}}>🏠</div>)}
          <button onClick={onClose} style={{position:"absolute",top:12,right:12,background:"rgba(0,0,0,.6)",border:"none",color:"#fff",width:34,height:34,borderRadius:"50%",cursor:"pointer",fontSize:18,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
          {imgs.length>1&&<div style={{position:"absolute",bottom:12,left:12,background:"rgba(0,0,0,.6)",color:"#fff",fontSize:11,fontWeight:700,padding:"3px 10px",borderRadius:18}}>📷 {imgs.length} {isEn?"photos":"صور"}</div>}
          <div style={{position:"absolute",top:12,left:12}}><span style={{display:"inline-flex",alignItems:"center",gap:5,padding:"4px 11px",borderRadius:20,background:sc.bg,color:sc.c,fontSize:11,fontWeight:700,border:`1px solid ${sc.c}44`,backdropFilter:"blur(4px)"}}><span style={{width:6,height:6,borderRadius:"50%",background:sc.c}}/>{isEn?statusEn[p.status]:p.status}</span></div>
        </div>

        <div style={{padding:"20px 22px"}}>
          {/* Title */}
          <div style={{fontWeight:900,fontSize:20,color:"#1e3a7a",marginBottom:4}}>{p.name}</div>
          <div style={{fontSize:12,color:"#5a6a90",marginBottom:14}}>📍 {p.address}</div>

          {/* Price */}
          <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}>
            {p.rentPrice&&<div style={{background:"rgba(22,163,74,.1)",border:"1px solid rgba(22,163,74,.25)",borderRadius:10,padding:"8px 14px",fontSize:14,color:"#16a34a",fontWeight:900}}>🏠 {Number(p.rentPrice).toLocaleString()} {isEn?"SAR/yr":"﷼/سنة"}</div>}
            {p.salePrice&&<div style={{background:"rgba(180,83,9,.08)",border:"1px solid rgba(180,83,9,.2)",borderRadius:10,padding:"8px 14px",fontSize:14,color:"#b45309",fontWeight:900}}>💰 {Number(p.salePrice).toLocaleString()} {isEn?"SAR":"﷼"}</div>}
          </div>

          {/* Details grid */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:16}}>
            {[
              p.type&&{l:isEn?"Type":"النوع",v:p.type,i:"🏢"},
              p.dealType&&{l:isEn?"Deal":"الصفقة",v:p.dealType,i:"📋"},
              p.area&&{l:isEn?"Area":"المساحة",v:p.area+" م²",i:"📐"},
              p.builtArea&&{l:isEn?"Built":"المبني",v:p.builtArea+" م²",i:"🏗️"},
              p.furnished!==undefined&&{l:isEn?"Furnished":"مفروش",v:p.furnished?(isEn?"Yes":"نعم"):(isEn?"No":"لا"),i:"🛋️"},
              p.refNo&&{l:isEn?"Ref":"المرجع",v:p.refNo,i:"🔖"},
            ].filter(Boolean).map((s,i)=>(
              <div key={i} style={{background:"#edf1fb",borderRadius:10,padding:"10px 12px",border:"1px solid rgba(74,158,255,.2)"}}>
                <div style={{fontSize:10,color:"#5a6a90",marginBottom:3}}>{s.i} {s.l}</div>
                <div style={{fontSize:13,fontWeight:700,color:"#2a4d9b"}}>{s.v}</div>
              </div>
            ))}
          </div>

          {/* License */}
          {p.adLicenseNo&&<div style={{background:"#1a4faa18",border:"1px solid #2563c740",borderRadius:10,padding:"8px 12px",marginBottom:14,fontSize:11,color:"#2a4d9b",fontWeight:700}}>🏛️ {isEn?"Ad License:":"رخصة إعلانية:"} {p.adLicenseNo}</div>}

          {/* Notes */}
          {p.notes&&<div style={{background:"#edf1fb",border:"1px solid rgba(74,158,255,.2)",borderRadius:12,padding:"12px 14px",marginBottom:16,fontSize:13,color:"#2a4d9b",lineHeight:1.9,whiteSpace:"pre-line",direction:"auto"}}>{p.notes}</div>}

          {/* Map */}
          {p.mapUrl&&<a href={p.mapUrl} target="_blank" rel="noopener noreferrer" style={{display:"flex",alignItems:"center",justifyContent:"center",gap:6,background:"rgba(22,163,74,.1)",border:"1px solid rgba(22,163,74,.3)",color:"#16a34a",borderRadius:11,padding:"10px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,width:"100%",boxSizing:"border-box",marginBottom:12}}>📍 {isEn?"View on Map":"عرض على الخريطة"}</a>}

          {/* Actions */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
            <a href={`tel:${PHONE}`} style={{display:"flex",alignItems:"center",justifyContent:"center",gap:4,background:"#1e3a7a",color:"#fff",borderRadius:11,padding:"11px 6px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:12}}>📞 {isEn?"Call":"اتصال"}</a>
            <a href={`https://wa.me/${WA_NUMBER}?text=${waMsg}`} target="_blank" rel="noopener noreferrer" style={{display:"flex",alignItems:"center",justifyContent:"center",gap:4,background:"rgba(37,211,102,.1)",border:"1px solid rgba(37,211,102,.3)",color:"#16a34a",borderRadius:11,padding:"11px 6px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:12}}><WaIcon size={13}/> {isEn?"WhatsApp":"واتساب"}</a>
            <button onClick={()=>onShare(p)} style={{display:"flex",alignItems:"center",justifyContent:"center",gap:4,background:"rgba(74,158,255,.1)",border:"1px solid rgba(74,158,255,.25)",color:"#2a4d9b",borderRadius:11,padding:"11px 6px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:12,cursor:"pointer"}}>📤 {isEn?"Share":"شارك"}</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function PublicCard({ p, setLightbox, onShare, lang }) {
  const [hov,setHov]=useState(false);
  const [showModal,setShowModal]=useState(false);
  const isEn=lang==="en";
  const imgs=p.images||[]; const sc=SC[p.status]||SC["متوفر"];

  const trackView = async () => {
    try { await updateDoc(doc(db,"properties",p.id),{views:(p.views||0)+1}); } catch(e) {}
  };
  const waMsg=encodeURIComponent(`${isEn?"Hello, I'd like to inquire about":"مرحباً، أود الاستفسار عن"} ${p.name} - ${p.address}${p.rentPrice?" | "+Number(p.rentPrice).toLocaleString()+" ﷼/سنة":""}${p.salePrice?" | "+Number(p.salePrice).toLocaleString()+" ﷼":""}`);
  const statusEn={"متوفر":"Available","مؤجر":"Rented","مباع":"Sold","قريب الانتهاء":"Expiring","صيانة":"Maintenance"};
  return (
    <>
    {showModal&&<PropertyModal p={p} onClose={()=>setShowModal(false)} setLightbox={setLightbox} onShare={onShare} lang={lang}/>}
    <div style={{background:"linear-gradient(160deg,#ffffff,#f5f8ff)",border:`1px solid ${sc.c}25`,borderRadius:18,overflow:"hidden",transition:"all .15s",transform:hov?"translateY(-3px)":"none",boxShadow:hov?`0 12px 36px ${sc.c}18`:"none",display:"flex",flexDirection:"column"}}
      onMouseEnter={()=>setHov(true)} onMouseLeave={()=>setHov(false)}>
      <div style={{position:"relative",height:185,background:"#edf1fb",cursor:imgs.length>0?"pointer":"default"}} onClick={()=>{if(imgs.length>0){setLightbox({images:imgs,idx:0});trackView();}}}>
        {imgs.length>0?(<><img src={imgs[0]} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/><div style={{position:"absolute",inset:0,background:"linear-gradient(to top,#07184088,transparent 55%)",pointerEvents:"none"}}/>{imgs.length>1&&<div style={{position:"absolute",bottom:10,left:10,background:"#000a",color:"#fff",fontSize:11,fontWeight:700,padding:"3px 10px",borderRadius:18}}>📷 {imgs.length}</div>}</>):(
          <div style={{width:"100%",height:"100%",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",color:"#1e3a7a",gap:8}}><span style={{fontSize:38}}>🏠</span></div>
        )}
        <div style={{position:"absolute",top:10,right:10}}><span style={{display:"inline-flex",alignItems:"center",gap:5,padding:"4px 11px",borderRadius:20,background:sc.bg,color:sc.c,fontSize:11,fontWeight:700,border:`1px solid ${sc.c}33`}}><span style={{width:6,height:6,borderRadius:"50%",background:sc.c}}/>{isEn?statusEn[p.status]:p.status}</span></div>
        {p.furnished&&<div style={{position:"absolute",top:10,left:10,background:"#fbbf2420",color:"#fbbf24",fontSize:10,fontWeight:700,padding:"3px 9px",borderRadius:18,border:"1px solid #fbbf2440"}}>🛋️ {isEn?"Furnished":"مفروش"}</div>}
        <button onClick={e=>{e.stopPropagation();onShare(p);}} style={{position:"absolute",bottom:10,left:10,background:"#1a4faa",border:"none",color:"#fff",borderRadius:20,padding:"4px 10px",fontSize:11,fontWeight:700,cursor:"pointer"}}>📤 {isEn?"Share":"مشاركة"}</button>
      </div>
      <div style={{padding:"14px 16px",display:"flex",flexDirection:"column",flex:1}}>
        <div style={{fontWeight:900,fontSize:15,color:"#1e3a7a",marginBottom:3}}>{p.name}</div>
        <div style={{fontSize:11,color:"#5a6a90",marginBottom:8}}>📍 {p.address}</div>
        <div style={{marginBottom:9}}>{p.adLicenseNo?<span style={{display:"inline-flex",alignItems:"center",gap:5,background:"#1a4faa22",border:"1px solid #2563c744",color:"#2a4d9b",borderRadius:8,padding:"3px 10px",fontSize:10,fontWeight:700}}>🏛️ {isEn?"Ad License:":"رخصة إعلانية:"} {p.adLicenseNo}</span>:<span style={{display:"inline-flex",alignItems:"center",gap:5,background:"#ef444418",border:"1px solid #ef444430",color:"#f87171",borderRadius:8,padding:"3px 10px",fontSize:10,fontWeight:700}}>⚠️ {isEn?"Pending License":"قيد الترخيص"}</span>}</div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:9}}>
          {p.rentPrice&&<div style={{background:"rgba(22,163,74,.1)",border:"1px solid rgba(22,163,74,.25)",borderRadius:9,padding:"5px 10px",fontSize:11,color:"#16a34a",fontWeight:700}}>🏠 {Number(p.rentPrice).toLocaleString()} {isEn?"SAR/yr":"﷼/سنة"}</div>}
          {p.salePrice&&<div style={{background:"rgba(180,83,9,.08)",border:"1px solid rgba(180,83,9,.2)",borderRadius:9,padding:"5px 10px",fontSize:11,color:"#b45309",fontWeight:700}}>💰 {Number(p.salePrice).toLocaleString()} {isEn?"SAR":"﷼"}</div>}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:10}}>
          {p.area&&<div style={{background:"#f0f4fc",borderRadius:8,padding:"5px 9px",fontSize:11,color:"#5a6a90"}}>📐 {p.area} م²</div>}
          {p.builtArea&&<div style={{background:"#f0f4fc",borderRadius:8,padding:"5px 9px",fontSize:11,color:"#5a6a90"}}>🏗️ {p.builtArea} م²</div>}
        </div>

        {/* Notes preview - 2 lines only */}
        {p.notes&&<div style={{fontSize:11,color:"#5a6a90",background:"#f0f4fc",borderRadius:8,padding:"6px 10px",marginBottom:10,overflow:"hidden",display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical",lineHeight:1.7}}>💬 {p.notes}</div>}

        {/* View details button */}
        <button onClick={()=>{setShowModal(true);trackView();}} style={{width:"100%",background:"#edf1fb",border:"1px solid rgba(74,158,255,.2)",color:"#1e3a7a",borderRadius:10,padding:"9px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:12,cursor:"pointer",marginBottom:9,transition:"background .2s"}}
          onMouseEnter={e=>e.currentTarget.style.background="#dde8f8"}
          onMouseLeave={e=>e.currentTarget.style.background="#edf1fb"}
        >🔍 {isEn?"View Full Details":"عرض التفاصيل كاملة"}</button>

        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:7,marginBottom:p.mapUrl?8:0}}>
          <a href={`tel:${PHONE}`} style={{display:"flex",alignItems:"center",justifyContent:"center",gap:4,background:"#1e3a7a",color:"#fff",borderRadius:10,padding:"9px 6px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11}}>📞 {isEn?"Call":"اتصال"}</a>
          <a href={`https://wa.me/${WA_NUMBER}?text=${waMsg}`} target="_blank" rel="noopener noreferrer" style={{display:"flex",alignItems:"center",justifyContent:"center",gap:4,background:"rgba(37,211,102,.1)",border:"1px solid rgba(37,211,102,.35)",color:"#16a34a",borderRadius:10,padding:"9px 6px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11}}><WaIcon size={12}/> {isEn?"WhatsApp":"واتساب"}</a>
          <button onClick={()=>onShare(p)} style={{display:"flex",alignItems:"center",justifyContent:"center",gap:4,background:"rgba(74,158,255,.1)",border:"1px solid rgba(74,158,255,.25)",color:"#2a4d9b",borderRadius:10,padding:"9px 6px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,cursor:"pointer"}}>📤 {isEn?"Share":"شارك"}</button>
        </div>
        {p.mapUrl&&<a href={p.mapUrl} target="_blank" rel="noopener noreferrer" style={{display:"flex",alignItems:"center",justifyContent:"center",gap:6,background:"rgba(22,163,74,.1)",border:"1px solid rgba(22,163,74,.3)",color:"#16a34a",borderRadius:10,padding:"8px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:12,width:"100%",boxSizing:"border-box"}}>📍 {isEn?"View on Map":"عرض على الخريطة"}</a>}
      </div>
    </div>
    </>
  );
}

// ── Admin Card ────────────────────────────────────────────────────────────────
function AdminCard({ p, onEdit, onDelete, onChangeStatus, setLightbox, onShare, userRole }) {
  const [hov,setHov]=useState(false);
  const imgs=p.images||[]; const sc=SC[p.status]||SC["متوفر"];
  return (
    <div style={{background:"linear-gradient(160deg,#ffffff,#f5f8ff)",border:`1px solid ${!p.adLicenseNo?"#ef444435":sc.c+"25"}`,borderRadius:18,overflow:"hidden",transition:"all .15s",transform:hov?"translateY(-3px)":"none"}}
      onMouseEnter={()=>setHov(true)} onMouseLeave={()=>setHov(false)}>
      {!p.adLicenseNo&&<div style={{background:"#ef444415",borderBottom:"1px solid #ef444428",padding:"5px 12px",fontSize:10,color:"#f87171",fontWeight:700}}>⚠️ يجب إضافة رقم الترخيص الإعلاني</div>}
      <div style={{position:"relative",height:148,background:"#edf1fb",cursor:imgs.length>0?"pointer":"default"}} onClick={()=>imgs.length>0&&setLightbox({images:imgs,idx:0})}>
        {imgs.length>0?(<><img src={imgs[0]} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/><div style={{position:"absolute",inset:0,background:"linear-gradient(to top,#07184088,transparent 55%)",pointerEvents:"none"}}/></>):(<div style={{width:"100%",height:"100%",display:"flex",alignItems:"center",justifyContent:"center",color:"#1e3a7a",fontSize:32}}>🏠</div>)}
        <div style={{position:"absolute",top:8,right:8}}><span style={{display:"inline-flex",alignItems:"center",gap:4,padding:"3px 10px",borderRadius:20,background:sc.bg,color:sc.c,fontSize:10,fontWeight:700,border:`1px solid ${sc.c}33`}}><span style={{width:5,height:5,borderRadius:"50%",background:sc.c}}/>{p.status}</span></div>
        {p.refNo&&<div style={{position:"absolute",bottom:8,left:8,background:"#1a4faa",color:"#fff",fontSize:9,padding:"2px 7px",borderRadius:12,fontWeight:700}}>{p.refNo}</div>}
      </div>
      <div style={{padding:"11px 13px"}}>
        <div style={{fontWeight:900,fontSize:13,color:"#1e3a7a",marginBottom:2}}>{p.name}</div>
        <div style={{fontSize:10,color:"#5a6a90",marginBottom:7}}>📍 {p.address}</div>
        <div style={{background:"#f0f4fc",borderRadius:9,padding:"8px 10px",marginBottom:7,border:"1px solid rgba(74,158,255,.15)"}}>
          <div style={{fontSize:9,color:p.adLicenseNo?"#16a34a":"#ef4444",marginBottom:2}}>🏛️ رخصة: {p.adLicenseNo||"غير مُدخل ⚠️"}</div>
          <div style={{fontSize:9,color:"#5a6a90",marginBottom:2}}>📋 عقد: {p.marketingContractNo||"—"}</div>
          {userRole==="admin"&&p.ownerName&&<div style={{fontSize:9,color:"#1e3a7a",marginBottom:2}}>👤 {p.ownerName} {p.ownerPhone&&"— "+p.ownerPhone}</div>}
          {userRole==="employee"&&<div style={{fontSize:9,color:"#ef4444",marginBottom:2}}>🔒 بيانات المالك محجوبة</div>}
          <div style={{fontSize:9,color:"#2a4d9b",fontWeight:700}}>👁️ {p.views||0} مشاهدة</div>
        </div>
        <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:7}}>
          {p.rentPrice&&<div style={{background:"#4ade8015",border:"1px solid #4ade8025",borderRadius:7,padding:"3px 8px",fontSize:9,color:"#4ade80",fontWeight:700}}>إيجار: {Number(p.rentPrice).toLocaleString()}</div>}
          {p.salePrice&&<div style={{background:"#fbbf2415",border:"1px solid #fbbf2425",borderRadius:7,padding:"3px 8px",fontSize:9,color:"#fbbf24",fontWeight:700}}>بيع: {Number(p.salePrice).toLocaleString()}</div>}
          {p.minPrice&&<div style={{background:"#f8717115",border:"1px solid #f8717125",borderRadius:7,padding:"3px 8px",fontSize:9,color:"#f87171",fontWeight:700}}>أدنى: {Number(p.minPrice).toLocaleString()}</div>}
        </div>
        <div style={{display:"flex",gap:4,flexWrap:"wrap",marginBottom:8}}>{STATUS_OPTIONS.map(s=>(<button key={s} onClick={()=>onChangeStatus(p.id,s)} style={{padding:"3px 8px",borderRadius:6,border:"none",cursor:"pointer",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:9,background:p.status===s?"linear-gradient(135deg,#1e3a7a,#2a4d9b)":"#0e2050",color:p.status===s?"#fff":"#4a6fa5"}}>{s}</button>))}</div>
        <div style={{display:"flex",gap:6}}>
          {userRole==="admin"&&<button onClick={()=>onEdit(p)} style={{flex:1,background:"#ddeeff",border:"1px solid rgba(74,158,255,.35)",color:"#2a4d9b",borderRadius:8,padding:"7px",fontFamily:"'Cairo',sans-serif",fontWeight:600,fontSize:11,cursor:"pointer"}}>✏️ تعديل</button>}
          <button onClick={()=>onShare(p)} style={{background:"#6366f118",border:"1px solid #6366f130",color:"#a5b4fc",borderRadius:8,padding:"7px 10px",cursor:"pointer",fontSize:11}}>📤</button>
          {imgs.length>0&&<button onClick={()=>setLightbox({images:imgs,idx:0})} style={{background:"#1a4faa22",border:"1px solid #2563c740",color:"#2a4d9b",borderRadius:8,padding:"7px 10px",cursor:"pointer",fontSize:11}}>🖼️</button>}
          {p.mapUrl&&<a href={p.mapUrl} target="_blank" rel="noopener noreferrer" style={{background:"#16a34a18",border:"1px solid #16a34a40",color:"#4ade80",borderRadius:8,padding:"7px 10px",fontSize:11,textDecoration:"none",display:"flex",alignItems:"center"}}>📍</a>}
          {userRole==="admin"&&<button onClick={()=>onDelete(p.id)} style={{background:"#ef444414",border:"1px solid #ef444428",color:"#f87171",borderRadius:8,padding:"7px 10px",cursor:"pointer",fontSize:11}}>🗑️</button>}
        </div>
      </div>
    </div>
  );
}

// ── Pages ─────────────────────────────────────────────────────────────────────
function HomePage({ setPage, lang, darkMode, T }) {
  const isEn=lang==="en";
  const services=isEn?[
    {icon:"🏢",t:"Buy & Rent",d:"Best real estate opportunities at competitive prices"},
    {icon:"🗝️",t:"Property Management",d:"Full management — rent collection, contracts, tenant relations"},
    {icon:"📊",t:"Valuation",d:"Professional valuation per Real Estate General Authority standards"},
    {icon:"💡",t:"Consulting",d:"Market analysis and investment opportunity guidance"},
    {icon:"🔧",t:"Maintenance",d:"Comprehensive maintenance services for your properties"},
    {icon:"🛋️",t:"Furnished Rental",d:"Fully furnished units ready to move in immediately"},
    {icon:"📋",t:"Registry Services",d:"Property subdivisions, deed updates, official procedures"},
    {icon:"🏛️",t:"Ejar Platform",d:"Lease contract documentation and legal preparation"},
  ]:[
    {icon:"🏢",t:"بيع وإيجار العقارات",d:"أفضل الفرص العقارية من شقق وفلل ومحلات بأسعار تنافسية"},
    {icon:"🗝️",t:"إدارة العقارات",d:"إدارة متكاملة — تحصيل الإيجارات، إدارة العقود، متابعة المستأجرين"},
    {icon:"📊",t:"تقييم العقارات",d:"تقييم احترافي وفق معايير الهيئة العامة للعقار"},
    {icon:"💡",t:"الاستشارات العقارية",d:"استشارات متخصصة تشمل تحليل السوق وتقييم الفرص الاستثمارية"},
    {icon:"🔧",t:"صيانة العقارات",d:"خدمات صيانة شاملة للحفاظ على قيمة عقاراتكم"},
    {icon:"🛋️",t:"تأجير المفروش",d:"وحدات مفروشة متكاملة جاهزة للسكن فوراً"},
    {icon:"📋",t:"فرز ودمج العقارات",d:"إجراءات فرز العقارات ودمجها وتحديث الصكوك"},
    {icon:"🏛️",t:"خدمات السجل العقاري",d:"إفراغات، تحديث صكوك، عقود إيجار في منصة إيجار"},
  ];

  return (
    <div style={{paddingTop:96,background:T.bg}}>

      {/* ── HERO ── */}
      <div style={{minHeight:"calc(100vh - 64px)",background:"linear-gradient(150deg,#0a1538 0%,#1e3a7a 55%,#2a4d9b 100%)",display:"flex",alignItems:"center",justifyContent:"center",position:"relative",overflow:"hidden",padding:"60px 24px"}}>
        <div style={{position:"absolute",inset:0,backgroundImage:"radial-gradient(circle at 25% 75%, rgba(74,158,255,.1) 0%,transparent 50%),radial-gradient(circle at 75% 25%, rgba(74,158,255,.07) 0%,transparent 50%)"}}/>
        <div style={{position:"absolute",inset:0,backgroundImage:"linear-gradient(rgba(74,158,255,.04) 1px,transparent 1px),linear-gradient(90deg,rgba(74,158,255,.04) 1px,transparent 1px)",backgroundSize:"50px 50px"}}/>
        <div style={{position:"relative",zIndex:2,textAlign:"center",maxWidth:800}}>
          <div style={{display:"inline-flex",alignItems:"center",gap:8,background:"rgba(74,158,255,.1)",border:"1px solid rgba(74,158,255,.25)",borderRadius:30,padding:"6px 20px",fontSize:12,color:"#7ab8ff",marginBottom:32}}>
            <span style={{width:7,height:7,borderRadius:"50%",background:"#4a9eff",boxShadow:"0 0 8px #4a9eff",display:"inline-block"}}/>
            {isEn?"Licensed by Real Estate General Authority":"مرخصون من الهيئة العامة للعقار"}
          </div>

          {/* Logo */}
          <div style={{width:170,height:170,borderRadius:24,background:"white",margin:"0 auto 30px",padding:12,boxShadow:"0 12px 50px rgba(74,158,255,.22),0 0 0 1px rgba(74,158,255,.12)"}}>
            <img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1778958489/WhatsApp_Image_2026-05-16_at_9.59.47_PM_zhmw6y.jpg" alt="Logo" style={{width:"100%",height:"100%",objectFit:"contain"}}/>
          </div>

          <p style={{fontSize:13,color:"rgba(255,255,255,.38)",letterSpacing:.5,marginBottom:14}}>Khalid M. A. Ghafour Al-Shaikh Est. | Real Estate Services</p>
          <h1 style={{fontSize:"clamp(30px,5.5vw,56px)",fontWeight:900,color:"white",lineHeight:1.15,marginBottom:18}}>
            {isEn?"Khalid M. A. Ghafour":"مؤسسة خالد محمد"}<br/>
            <span style={{color:"#7ab8ff"}}>{isEn?"Al-Shaikh Est.":"عبدالغفور الشيخ"}</span>
          </h1>
          <p style={{fontSize:15,color:"rgba(255,255,255,.55)",lineHeight:1.8,maxWidth:520,margin:"0 auto 40px"}}>
            {isEn?"Your trusted partner for all real estate services since 1997":"شريكك الموثوق في جميع الخدمات العقارية منذ عام 1997م"}
          </p>
          <div style={{display:"flex",gap:14,justifyContent:"center",flexWrap:"wrap",marginBottom:58}}>
            <button onClick={()=>setPage("properties")} style={{background:"#4a9eff",color:"#fff",border:"none",borderRadius:14,padding:"14px 30px",fontFamily:"'Cairo',sans-serif",fontWeight:900,fontSize:15,cursor:"pointer",boxShadow:"0 6px 22px rgba(74,158,255,.45)"}}>🏘️ {isEn?"Browse Properties":"تصفح العقارات"}</button>
            <a href={`tel:${PHONE}`} style={{background:"rgba(255,255,255,.08)",border:"1px solid rgba(255,255,255,.2)",color:"#fff",borderRadius:14,padding:"14px 24px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:15,display:"flex",alignItems:"center",gap:8}}>📞 {isEn?"Contact Now":"تواصل الآن"}</a>
          </div>
          <div style={{display:"flex",justifyContent:"center",gap:0,borderTop:"1px solid rgba(74,158,255,.12)",paddingTop:36,flexWrap:"wrap"}}>
            {[[isEn?"1997":"١٩٩٧",isEn?"Est. Year":"سنة التأسيس"],[isEn?"27+":"٢٧+",isEn?"Years":"سنة خبرة"],[isEn?"100%":"١٠٠٪",isEn?"Compliant":"امتثال للهيئة"],[isEn?"8+":"٨+",isEn?"Services":"خدمة عقارية"]].map(([n,l],i)=>(
              <div key={i} style={{textAlign:"center",padding:"0 28px",borderLeft:i>0?"1px solid rgba(74,158,255,.12)":"none"}}>
                <div style={{fontSize:28,fontWeight:900,color:"#7ab8ff"}}>{n}</div>
                <div style={{fontSize:11,color:"rgba(255,255,255,.36)",marginTop:3}}>{l}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* ── SERVICES BAND ── */}
      <div style={{background:"#162f63",borderTop:"1px solid rgba(74,158,255,.15)",borderBottom:"1px solid rgba(74,158,255,.15)"}}>
        <div style={{maxWidth:1100,margin:"0 auto",display:"flex",overflowX:"auto",scrollbarWidth:"none"}}>
          {[
            ["🏢",isEn?"Buy & Rent":"بيع وإيجار"],
            ["🗝️",isEn?"Management":"إدارة العقارات"],
            ["📊",isEn?"Valuation":"التقييم العقاري"],
            ["💡",isEn?"Consulting":"الاستشارات"],
            ["🔧",isEn?"Maintenance":"الصيانة"],
            ["🛋️",isEn?"Furnished":"تأجير المفروش"],
            ["📋",isEn?"Registry":"خدمات السجل"],
            ["🏛️",isEn?"Ejar Platform":"منصة إيجار"],
          ].map(([icon,label],i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",gap:10,padding:"14px 24px",borderLeft:i>0?"1px solid rgba(255,255,255,.07)":"none",flexShrink:0,cursor:"pointer",transition:"background .2s"}}
              onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.05)"}
              onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
              <div style={{width:32,height:32,borderRadius:7,background:"rgba(74,158,255,.12)",border:"1px solid rgba(74,158,255,.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,flexShrink:0}}>{icon}</div>
              <span style={{fontSize:12,fontWeight:600,color:"rgba(255,255,255,.65)",whiteSpace:"nowrap"}}>{label}</span>
            </div>
          ))}
        </div>
      </div>

      {/* ── SERVICES ── */}
      <div style={{height:2,background:"linear-gradient(90deg,transparent,#4a9eff,transparent)",opacity:.3}}/>
      <div style={{background:T.bg,padding:"74px 24px"}}>
        <div style={{maxWidth:1100,margin:"0 auto"}}>
          <div style={{textAlign:"center",marginBottom:46}}>
            <div style={{display:"inline-block",background:"rgba(74,158,255,.08)",border:"1px solid rgba(74,158,255,.18)",color:"#2a4d9b",borderRadius:20,padding:"5px 16px",fontSize:11,fontWeight:700,marginBottom:12}}>{isEn?"Our Services":"خدماتنا"}</div>
            <h2 style={{fontSize:"clamp(20px,3.5vw,32px)",fontWeight:900,color:T.text}}>{isEn?"Complete Real Estate Services":"خدمات عقارية متكاملة تحت سقف واحد"}</h2>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:14}}>
            {services.map((s,i)=>(<div key={i} style={{background:darkMode?"#1a2d6b":"white",border:`1px solid rgba(74,158,255,.${darkMode?"1":"07"})`,borderRadius:17,padding:"24px 20px",transition:"all .25s",cursor:"pointer"}} onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-5px)";e.currentTarget.style.borderColor="rgba(74,158,255,.3)";e.currentTarget.style.boxShadow="0 10px 30px rgba(74,158,255,.1)";}} onMouseLeave={e=>{e.currentTarget.style.transform="";e.currentTarget.style.borderColor="";e.currentTarget.style.boxShadow="";}}>
              <div style={{width:48,height:48,borderRadius:14,background:darkMode?"rgba(74,158,255,.1)":"#edf1fb",border:`1px solid rgba(74,158,255,.${darkMode?"18":"1"})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,marginBottom:14}}>{s.icon}</div>
              <div style={{fontSize:14,fontWeight:800,color:T.text,marginBottom:7}}>{s.t}</div>
              <div style={{fontSize:12,color:T.text3,lineHeight:1.8}}>{s.d}</div>
            </div>))}
          </div>
        </div>
      </div>

      {/* ── NUMBERS ── */}
      <div style={{height:2,background:"linear-gradient(90deg,transparent,#4a9eff,transparent)",opacity:.3}}/>
      <div style={{background:"#1e3a7a",padding:"64px 24px"}}>
        <div style={{maxWidth:900,margin:"0 auto",textAlign:"center"}}>
          <div style={{display:"inline-block",background:"rgba(74,158,255,.1)",border:"1px solid rgba(74,158,255,.25)",color:"#7ab8ff",borderRadius:20,padding:"5px 16px",fontSize:11,fontWeight:700,marginBottom:12}}>{isEn?"Our Numbers":"أرقامنا"}</div>
          <h2 style={{fontSize:"clamp(20px,3.5vw,30px)",fontWeight:900,color:"white",marginBottom:44}}>{isEn?"Trust Built Over Years of Experience":"ثقة مبنية على سنوات من الخبرة"}</h2>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:20}}>
            {[[isEn?"27+":"٢٧+",isEn?"Years in Market":"سنة في السوق"],[isEn?"500+":"٥٠٠+",isEn?"Properties":"عقار تم تسويقه"],[isEn?"1000+":"١٠٠٠+",isEn?"Clients":"عميل راضٍ"],[isEn?"100%":"١٠٠٪",isEn?"Authority Compliant":"امتثال للهيئة"]].map(([n,l],i)=>(
              <div key={i} style={{textAlign:"center"}}>
                <div style={{fontSize:"clamp(24px,3vw,36px)",fontWeight:900,color:"#7ab8ff",marginBottom:6}}>{n}</div>
                <div style={{fontSize:12,color:"rgba(255,255,255,.45)"}}>{l}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* ── NUMBERS ── */}
      <div style={{height:2,background:"linear-gradient(90deg,transparent,#4a9eff,transparent)",opacity:.3}}/>
      <div style={{background:"#1e3a7a",padding:"64px 24px"}}>
        <div style={{maxWidth:1100,margin:"0 auto",textAlign:"center"}}>
          <div style={{fontSize:11,fontWeight:700,color:"rgba(122,184,255,.6)",letterSpacing:1.5,marginBottom:10,textTransform:"uppercase"}}>{isEn?"Our Numbers":"أرقامنا"}</div>
          <h2 style={{fontSize:"clamp(20px,3vw,30px)",fontWeight:900,color:"white",marginBottom:44,fontFamily:"'Cairo',sans-serif"}}>{isEn?"Trust Built Over Years":"ثقة مبنية على سنوات من الإنجاز"}</h2>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:0,border:"1px solid rgba(255,255,255,.08)",borderRadius:16,overflow:"hidden"}}>
            {[[isEn?"1997":"١٩٩٧",isEn?"Est. Year":"سنة التأسيس"],[isEn?"27+":"٢٧+",isEn?"Years Experience":"سنة خبرة"],[isEn?"500+":"٥٠٠+",isEn?"Properties":"عقار تم تسويقه"],[isEn?"100%":"١٠٠٪",isEn?"Authority Compliant":"امتثال للهيئة"]].map(([n,l],i)=>(
              <div key={i} style={{padding:"36px 20px",textAlign:"center",borderLeft:i>0?"1px solid rgba(255,255,255,.08)":"none",transition:"background .2s"}}
                onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,.04)"}
                onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                <div style={{fontSize:"clamp(28px,3.5vw,44px)",fontWeight:900,color:"#7ab8ff",marginBottom:8,fontFamily:"'Cairo',sans-serif"}}>{n}</div>
                <div style={{fontSize:12,color:"rgba(255,255,255,.4)",fontWeight:300}}>{l}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* ── CTA ── */}
      <div style={{height:2,background:"linear-gradient(90deg,transparent,#4a9eff,transparent)",opacity:.3}}/>
      <div style={{background:T.bg,padding:"70px 24px",textAlign:"center"}}>
        <h2 style={{fontSize:"clamp(22px,4vw,34px)",fontWeight:900,color:T.text,marginBottom:10}}>{isEn?"Looking for a property?":"هل تبحث عن عقار؟"}</h2>
        <p style={{color:T.text3,fontSize:14,marginBottom:30}}>{isEn?"Our team is ready 24/7":"فريقنا جاهز لمساعدتك على مدار الساعة"}</p>
        <div style={{display:"flex",gap:12,justifyContent:"center",flexWrap:"wrap"}}>
          <a href={`tel:${PHONE}`} style={{background:"#1e3a7a",color:"white",border:"none",borderRadius:13,padding:"13px 30px",fontFamily:"'Cairo',sans-serif",fontWeight:900,fontSize:15,cursor:"pointer",textDecoration:"none"}}>📞 {PHONE}</a>
          <button onClick={()=>setPage("properties")} style={{background:"#4a9eff",color:"white",border:"none",borderRadius:13,padding:"13px 26px",fontFamily:"'Cairo',sans-serif",fontWeight:900,fontSize:15,cursor:"pointer",boxShadow:"0 4px 16px rgba(74,158,255,.35)"}}>🏘️ {isEn?"Browse Properties":"تصفح العقارات"}</button>
          <a href={`https://wa.me/${WA_NUMBER}`} target="_blank" rel="noopener noreferrer" style={{background:"rgba(37,211,102,.1)",border:"1px solid rgba(37,211,102,.3)",color:"#25d366",borderRadius:13,padding:"13px 22px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:15,cursor:"pointer",textDecoration:"none",display:"flex",alignItems:"center",gap:8}}><WaIcon size={16}/> WhatsApp</a>
        </div>
      </div>

      {/* ── FOOTER ── */}
      <div style={{background:"#0a1538",borderTop:"2px solid rgba(74,158,255,.15)",padding:"60px 24px 0"}}>
        <div style={{maxWidth:1100,margin:"0 auto"}}>
          {/* Footer grid */}
          <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:48,marginBottom:48,flexWrap:"wrap"}}>
            {/* Brand col */}
            <div>
              <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:18}}>
                <div style={{width:46,height:46,borderRadius:10,background:"white",padding:4,overflow:"hidden",flexShrink:0}}>
                  <img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1778958489/WhatsApp_Image_2026-05-16_at_9.59.47_PM_zhmw6y.jpg" alt="" style={{width:"100%",height:"100%",objectFit:"contain"}}/>
                </div>
                <div>
                  <div style={{fontSize:12,fontWeight:900,color:"white",fontFamily:"'Cairo',sans-serif",lineHeight:1.3}}>مؤسسة خالد محمد عبدالغفور الشيخ</div>
                  <div style={{fontSize:9,color:"#7ab8ff",letterSpacing:.5}}>KHALID M. A. GHAFOUR AL-SHAIKH EST.</div>
                </div>
              </div>
              <p style={{fontSize:12,color:"rgba(255,255,255,.38)",lineHeight:1.9,fontWeight:300,marginBottom:22}}>
                {isEn?"A specialized real estate establishment founded in 1997 in the Eastern Province of Saudi Arabia. We provide comprehensive real estate services with transparency and professionalism.":"مؤسسة عقارية متخصصة تأسست عام 1997م في المنطقة الشرقية، نقدم خدمات عقارية شاملة للأفراد والشركات والمستثمرين بالشفافية والمهنية."}
              </p>
              <div style={{display:"flex",flexDirection:"column",gap:10}}>
                {[["📞",PHONE],["💬","WhatsApp: "+WA_NUMBER],["📍",isEn?"Eastern Province, KSA":"المنطقة الشرقية، المملكة العربية السعودية"]].map(([i,v])=>(
                  <div key={v} style={{fontSize:11,color:"rgba(255,255,255,.4)",display:"flex",alignItems:"center",gap:8}}><span style={{color:"#4a9eff"}}>{i}</span>{v}</div>
                ))}
              </div>
            </div>
            {/* Links cols */}
            {[
              {title:isEn?"Quick Links":"روابط سريعة",links:[[isEn?"Home":"الرئيسية","home"],[isEn?"Properties":"العقارات","properties"],[isEn?"Services":"خدماتنا","services"],[isEn?"About":"عن المؤسسة","about"]]},
              {title:isEn?"Services":"خدماتنا",links:[[isEn?"Buy & Rent":"بيع وإيجار",""],[isEn?"Management":"إدارة العقارات",""],[isEn?"Valuation":"التقييم العقاري",""],[isEn?"Consulting":"الاستشارات",""],[isEn?"Maintenance":"الصيانة",""]]}  ,
              {title:isEn?"Property Types":"أنواع العقارات",links:[[isEn?"Apartments":"شقق سكنية",""],[isEn?"Villas":"فلل وقصور",""],[isEn?"Commercial":"محلات تجارية",""],[isEn?"Offices":"مكاتب",""],[isEn?"Land":"أراضي",""]]}
            ].map((col,i)=>(
              <div key={i}>
                <div style={{fontSize:11,fontWeight:700,color:"white",marginBottom:18,letterSpacing:.5}}>{col.title}</div>
                <div style={{display:"flex",flexDirection:"column",gap:10}}>
                  {col.links.map(([label,p])=>(
                    <span key={label} onClick={()=>p&&setPage(p)} style={{fontSize:12,color:"rgba(255,255,255,.38)",cursor:p?"pointer":"default",transition:"color .2s"}}
                      onMouseEnter={e=>{if(p)e.target.style.color="#4a9eff";}}
                      onMouseLeave={e=>{e.target.style.color="rgba(255,255,255,.38)";}}>
                      {label}
                    </span>
                  ))}
                </div>
              </div>
            ))}
          </div>
          {/* Bottom bar */}
          <div style={{borderTop:"1px solid rgba(255,255,255,.07)",padding:"22px 0",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
            <div style={{fontSize:11,color:"rgba(255,255,255,.2)"}}>© 2025 {isEn?"Khalid M. A. Ghafour Al-Shaikh Est. All rights reserved.":"مؤسسة خالد محمد عبدالغفور الشيخ — جميع الحقوق محفوظة"}</div>
            <div style={{display:"inline-flex",alignItems:"center",gap:6,background:"rgba(74,158,255,.08)",border:"1px solid rgba(74,158,255,.15)",borderRadius:4,padding:"4px 14px",fontSize:11,color:"#7ab8ff"}}>🏛️ {isEn?"Licensed by Real Estate General Authority":"مرخصة من الهيئة العامة للعقار"}</div>
          </div>
        </div>
      </div>
    </div>
  );
}

function PropertiesPage({ props, isAdmin, userRole, onEdit, onDelete, onChangeStatus, setLightbox, onShare, onOpenAdd, lang, darkMode, T }) {
  const isEn=lang==="en";

  const dealLabels = isEn
    ? {all:"All","إيجار":"Rent","بيع":"Sale","إيجار وبيع":"Rent & Sale"}
    : {all:"الكل","إيجار":"إيجار","بيع":"بيع","إيجار وبيع":"إيجار وبيع"};

  const typeLabels = isEn
    ? {all:"All Types","شقة":"Apartment","فيلا":"Villa","محل تجاري":"Shop","مكتب":"Office","استوديو":"Studio","دوبلكس":"Duplex","أرض":"Land"}
    : {all:"كل الأنواع","شقة":"شقة","فيلا":"فيلا","محل تجاري":"محل","مكتب":"مكتب","استوديو":"استوديو","دوبلكس":"دوبلكس","أرض":"أرض"};

  const statusLabels = isEn
    ? {all:"All","متوفر":"Available","مؤجر":"Rented","مباع":"Sold","قريب الانتهاء":"Expiring","صيانة":"Maintenance"}
    : {all:"الكل","متوفر":"متوفر","مؤجر":"مؤجر","مباع":"مباع","قريب الانتهاء":"قريب الانتهاء","صيانة":"صيانة"};

  const [dealF,setDealF]   = useState("all");
  const [typeF,setTypeF]   = useState("all");
  const [statusF,setStatusF] = useState("all");
  const [search,setSearch] = useState("");
  const [sortF,setSortF]   = useState("default");

  const filtered=props.filter(p=>
    (dealF==="all"||p.dealType===dealF)&&
    (typeF==="all"||p.type===typeF)&&
    (statusF==="all"||p.status===statusF)&&
    (p.name.includes(search)||p.address.includes(search)||(p.ownerName||"").includes(search))
  ).sort((a,b)=>{
    if(sortF==="price_asc")  { const pa=Number(a.rentPrice||a.salePrice||0); const pb=Number(b.rentPrice||b.salePrice||0); return pa-pb; }
    if(sortF==="price_desc") { const pa=Number(a.rentPrice||a.salePrice||0); const pb=Number(b.rentPrice||b.salePrice||0); return pb-pa; }
    return 0;
  });

  const BtnStyle = (active, color="#1a4faa") => ({
    padding:"7px 13px",borderRadius:20,cursor:"pointer",
    fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,
    background:active?`linear-gradient(135deg,${color},${color}dd)`:T.bg2,
    color:active?"#fff":T.text3,
    border:active?"none":`1px solid ${T.border}`,
    whiteSpace:"nowrap", transition:"all .2s"
  });

  return (
    <div style={{paddingTop:isAdmin?90:64,minHeight:"100vh",background:T.bg}}>
      <div style={{background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",padding:"28px 24px 24px"}}>
        <div style={{maxWidth:1200,margin:"0 auto",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
          <div><div style={{fontWeight:900,fontSize:22,color:"#fff",marginBottom:3}}>{isAdmin?(isEn?"🔒 Admin Panel":"🔒 لوحة الإدارة"):(isEn?"🏘️ Available Properties":"🏘️ العقارات المتاحة")}</div><div style={{fontSize:12,color:"rgba(255,255,255,.5)"}}>Khalid M. A. Ghafour Al-Shaikh Est.</div></div>
          {isAdmin&&<div style={{display:"flex",gap:8}}>
            {userRole==="admin"&&<button onClick={()=>exportToExcel(props)} style={{background:"#16a34a22",border:"1px solid #16a34a44",color:"#4ade80",borderRadius:11,padding:"10px 18px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>📊 {isEn?"Export Excel":"تصدير Excel"}</button>}
            <button onClick={onOpenAdd} style={{background:"rgba(255,255,255,.15)",border:"1px solid rgba(255,255,255,.3)",color:"#fff",borderRadius:11,padding:"10px 22px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>+ {isEn?"Add Property":"إضافة عقار"}</button>
          </div>}
        </div>
      </div>
      <div style={{maxWidth:1200,margin:"0 auto",padding:"22px"}}>
        {isAdmin&&(()=>{const st={total:props.length,available:props.filter(p=>p.status==="متوفر").length,rented:props.filter(p=>p.status==="مؤجر").length,income:props.filter(p=>p.rentPrice).reduce((s,p)=>s+Number(p.rentPrice),0),noLicense:props.filter(p=>!p.adLicenseNo).length};return(<div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:9,marginBottom:20}}>{[{l:isEn?"Total":"الإجمالي",v:st.total,i:"🏢",c:"#2a4d9b"},{l:isEn?"Available":"متوفر",v:st.available,i:"✅",c:"#16a34a"},{l:isEn?"Rented":"مؤجر",v:st.rented,i:"🔑",c:"#7c3aed"},{l:isEn?"Annual Income":"الدخل السنوي",v:st.income.toLocaleString()+" ﷼",i:"💰",c:"#b45309"},{l:isEn?"No License":"بدون ترخيص",v:st.noLicense,i:"⚠️",c:"#dc2626"}].map((s,i)=>(<div key={i} style={{background:"white",border:`1px solid ${s.c}22`,borderRadius:12,padding:"12px 10px",position:"relative",overflow:"hidden",boxShadow:"0 2px 8px rgba(30,58,122,.07)"}}><div style={{position:"absolute",top:-8,left:-8,width:32,height:32,background:s.c+"15",borderRadius:"50%"}}/><div style={{fontSize:17,marginBottom:5}}>{s.i}</div><div style={{fontWeight:900,fontSize:15,color:s.c}}>{s.v}</div><div style={{fontSize:10,color:"#5a6a90",marginTop:1}}>{s.l}</div></div>))}</div>);})()}

        {/* Search */}
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder={isEn?"🔍 Search...":"🔍 ابحث بالاسم أو الموقع..."} style={{width:"100%",boxSizing:"border-box",background:T.bg2,border:`1px solid ${T.border}`,borderRadius:12,padding:"10px 16px",color:T.text,fontFamily:"'Cairo',sans-serif",fontSize:13,marginBottom:12}}/>

        {/* Row 1: Deal type */}
        <div style={{display:"flex",gap:7,marginBottom:8,flexWrap:"wrap"}}>
          {["all","إيجار","بيع","إيجار وبيع"].map(f=>(
            <button key={f} onClick={()=>setDealF(f)} style={BtnStyle(dealF===f,"#1a4faa")}>
              {f==="all"?dealLabels.all:dealLabels[f]}
            </button>
          ))}
        </div>

        {/* Row 2: Property type */}
        <div style={{display:"flex",gap:7,marginBottom:8,flexWrap:"wrap"}}>
          {["all",...PROPERTY_TYPES].map(f=>(
            <button key={f} onClick={()=>setTypeF(f)} style={BtnStyle(typeF===f,"#7c3aed")}>
              {f==="all"?typeLabels.all:typeLabels[f]}
            </button>
          ))}
        </div>

        {/* Row 3: Status - admin only */}
        {isAdmin&&<div style={{display:"flex",gap:7,marginBottom:12,flexWrap:"wrap"}}>
          {["all",...STATUS_OPTIONS].map(f=>(
            <button key={f} onClick={()=>setStatusF(f)} style={BtnStyle(statusF===f,"#0e7490")}>
              {f==="all"?statusLabels.all:statusLabels[f]}
            </button>
          ))}
        </div>}

        {/* Sort row */}
        <div style={{display:"flex",gap:7,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
          <span style={{fontSize:11,color:T.text3,fontWeight:700}}>{isEn?"Sort:":"ترتيب:"}</span>
          {[["default",isEn?"Default":"الافتراضي"],["price_asc",isEn?"Price ↑":"السعر من الأقل"],["price_desc",isEn?"Price ↓":"السعر من الأعلى"]].map(([v,l])=>(
            <button key={v} onClick={()=>setSortF(v)} style={BtnStyle(sortF===v,"#7c3aed")}>
              {l}
            </button>
          ))}
        </div>

        <div style={{fontSize:11,color:T.text3,marginBottom:12}}>{filtered.length} {isEn?"properties":"عقار"}</div>
        {filtered.length===0?(<div style={{textAlign:"center",padding:"60px 0",color:"#1e3a7a"}}><div style={{fontSize:46,marginBottom:10}}>🏚️</div><div style={{fontSize:13,fontWeight:600}}>{isEn?"No results":"لا توجد نتائج"}</div></div>):(
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(295px,1fr))",gap:15}}>
            {filtered.map(p=>isAdmin
              ? <AdminCard key={p.id} p={p} onEdit={onEdit} onDelete={onDelete} onChangeStatus={onChangeStatus} setLightbox={setLightbox} onShare={onShare} userRole={userRole}/>
              : <PublicCard key={p.id} p={p} setLightbox={setLightbox} onShare={onShare} lang={lang}/>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

// ── Clients Page ──────────────────────────────────────────────────────────────
function ClientsPage({ lang, T, darkMode, userRole, currentUser }) {
  const isEn = lang==="en";
  const isManager = userRole==="admin";
  const employees = USERS.filter(u=>u.role==="employee");
  const [clients, setClients] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [editClientId, setEditClientId] = useState(null);
  const [delId, setDelId] = useState(null);
  const [commentId, setCommentId] = useState(null);
  const [commentText, setCommentText] = useState("");
  const [assignModal, setAssignModal] = useState(null); // client to assign
  const [assignTo, setAssignTo] = useState("");
  const [assignMsg, setAssignMsg] = useState("");
  const emptyClient = { name:"", phone:"", area:"", clientType:"مستأجر", requestType:"إيجار", budget:"", paymentType:"", ownerPropertyType:"", investorType:"", notes:"", contactDate:"", clientStatus:"معلق" };
  const [form, setForm] = useState(emptyClient);

  useEffect(()=>{
    const unsub = onSnapshot(collection(db,"clients"), snap=>{
      const data = snap.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>b.createdAt?.localeCompare?.(a.createdAt)||0);
      setClients(data); setLoaded(true);
    }, ()=>setLoaded(true));
    return ()=>unsub();
  },[]);

  const openAdd = () => { setForm(emptyClient); setEditClientId(null); setShowForm(true); };
  const openEdit = (c) => { 
    const fd = {...emptyClient,...c};
    if(isManager===false && c.clientType==="مالك") fd.phone="";
    setForm(fd); setEditClientId(c.id); setShowForm(true); 
  };

  const save = async () => {
    if(!form.name||!form.phone) return alert("الاسم ورقم الجوال مطلوبان");
    const duplicate = clients.find(c=>c.phone===form.phone&&c.id!==editClientId);
    if(duplicate) { const ok=window.confirm(`⚠️ رقم الجوال مكرر!\nمسجل للعميل: "${duplicate.name}" — #${duplicate.clientNo}\nهل تريد الحفظ رغم التكرار؟`); if(!ok) return; }
    if(editClientId) {
      await updateDoc(doc(db,"clients",editClientId),{...form, updatedAt:new Date().toISOString(), isDuplicate:!!duplicate});
    } else {
      const maxNum = clients.reduce((max,c)=>Math.max(max,c.clientNo||0),0);
      await addDoc(collection(db,"clients"),{
        ...form, clientNo:maxNum+1,
        createdAt:new Date().toISOString(),
        createdDate:new Date().toLocaleDateString("en-GB"),
        comments:[], isDuplicate:!!duplicate,
        assignedTo: currentUser?.username||"",
        assignedToName: currentUser?.displayName||""
      });
    }
    setForm(emptyClient); setShowForm(false); setEditClientId(null);
  };

  const assignClient = async () => {
    if(!assignTo||!assignModal) return;
    const emp = USERS.find(u=>u.username===assignTo);
    const comment = { text:`📋 تم تحويل العميل إلى ${emp?.displayName||assignTo}${assignMsg?` — رسالة: ${assignMsg}`:""}`, date:new Date().toLocaleDateString("en-GB"), time:new Date().toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}), system:true };
    const existing = clients.find(c=>c.id===assignModal);
    const comments = [...(existing?.comments||[]), comment];
    await updateDoc(doc(db,"clients",assignModal),{ assignedTo:assignTo, assignedToName:emp?.displayName||assignTo, comments });
    setAssignModal(null); setAssignTo(""); setAssignMsg("");
  };

  const printClientCard = (c) => {
    const ratingColor = {"مستأجر":"#16a34a","مشتري":"#b45309","مالك":"#2a4d9b","مستثمر":"#7c3aed"};
    const win = window.open('','_blank','width=420,height=600');
    win.document.write(`
      <!DOCTYPE html><html dir="rtl" lang="ar">
      <head><meta charset="UTF-8">
      <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet">
      <style>
        *{box-sizing:border-box;margin:0;padding:0;}
        body{font-family:'Cairo',sans-serif;background:#f5f8ff;display:flex;justify-content:center;align-items:center;min-height:100vh;padding:20px;}
        .card{background:white;border-radius:16px;padding:28px;width:360px;box-shadow:0 4px 24px rgba(30,58,122,.12);border:1px solid rgba(74,158,255,.15);}
        .header{background:linear-gradient(135deg,#1e3a7a,#2a4d9b);border-radius:12px;padding:18px;margin-bottom:18px;display:flex;align-items:center;gap:14px;}
        .logo{width:50px;height:50px;border-radius:10px;background:white;overflow:hidden;flex-shrink:0;}
        .logo img{width:100%;height:100%;object-fit:contain;}
        .brand{color:white;}
        .brand-n{font-size:11px;font-weight:900;line-height:1.3;}
        .brand-s{font-size:9px;color:rgba(255,255,255,.55);}
        .client-num{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;}
        .num-badge{width:44px;height:44px;border-radius:12px;background:rgba(74,158,255,.1);border:1px solid rgba(74,158,255,.2);display:flex;align-items:center;justify-content:center;font-weight:900;color:#1e3a7a;font-size:14px;}
        .client-name{font-size:20px;font-weight:900;color:#1e3a7a;}
        .client-phone{font-size:13px;color:#5a6a90;margin-top:3px;}
        .type-badge{display:inline-flex;align-items:center;gap:6px;border-radius:20px;padding:4px 14px;font-size:11px;font-weight:700;margin-top:8px;}
        .divider{height:1px;background:rgba(74,158,255,.1);margin:14px 0;}
        .info-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px;}
        .info-box{background:#f5f8ff;border-radius:8px;padding:10px 12px;}
        .info-lbl{font-size:9px;color:#8899bb;margin-bottom:3px;font-weight:600;}
        .info-val{font-size:12px;font-weight:700;color:#1e3a7a;}
        .footer-bar{background:#f0f4fc;border-radius:10px;padding:12px 14px;display:flex;justify-content:space-between;align-items:center;}
        .footer-txt{font-size:10px;color:#8899bb;}
        .footer-phone{font-size:12px;font-weight:700;color:#1e3a7a;}
        @media print{body{background:white;padding:0;}.card{box-shadow:none;border:1px solid #ddd;}}
      </style>
      </head><body>
      <div class="card">
        <div class="header">
          <div class="logo"><img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1778958489/WhatsApp_Image_2026-05-16_at_9.59.47_PM_zhmw6y.jpg"/></div>
          <div class="brand">
            <div class="brand-n">مؤسسة خالد محمد عبدالغفور الشيخ</div>
            <div class="brand-s">Khalid M. A. Ghafour Al-Shaikh Est.</div>
          </div>
        </div>
        <div class="client-num">
          <div>
            <div class="client-name">${c.name||""}</div>
            <div class="client-phone">📞 ${c.phone||""}</div>
            <div class="type-badge" style="background:${(ratingColor[c.clientType]||"#2a4d9b")}18;color:${ratingColor[c.clientType]||"#2a4d9b"};border:1px solid ${ratingColor[c.clientType]||"#2a4d9b"}33">${c.clientType||""}</div>
          </div>
          <div class="num-badge">#${c.clientNo||"—"}</div>
        </div>
        <div class="divider"></div>
        <div class="info-grid">
          <div class="info-box"><div class="info-lbl">📍 الحي / الموقع</div><div class="info-val">${c.area||"—"}</div></div>
          <div class="info-box"><div class="info-lbl">💰 الميزانية</div><div class="info-val">${c.budget?Number(c.budget).toLocaleString()+" ﷼":"—"}</div></div>
          <div class="info-box"><div class="info-lbl">📋 الحالة</div><div class="info-val">${c.clientStatus||"معلق"}</div></div>
          <div class="info-box"><div class="info-lbl">📅 تاريخ التسجيل</div><div class="info-val">${c.createdDate||"—"}</div></div>
          ${c.ownerPropertyType?`<div class="info-box"><div class="info-lbl">🏠 نوع العقار</div><div class="info-val">${c.ownerPropertyType}</div></div>`:""}
          ${c.investorType?`<div class="info-box"><div class="info-lbl">📈 الاستثمار</div><div class="info-val">${c.investorType}</div></div>`:""}
          ${c.paymentType?`<div class="info-box"><div class="info-lbl">💳 طريقة الدفع</div><div class="info-val">${c.paymentType}</div></div>`:""}
          ${c.contactDate?`<div class="info-box"><div class="info-lbl">📞 تاريخ التواصل</div><div class="info-val">${c.contactDate}</div></div>`:""}
        </div>
        ${c.notes?`<div class="info-box" style="margin-bottom:14px;background:#f5f8ff;border-radius:8px;padding:10px 12px;"><div class="info-lbl">💬 ملاحظات</div><div style="font-size:11px;color:#1e3a7a;line-height:1.7;margin-top:3px">${c.notes}</div></div>`:""}
        <div class="footer-bar">
          <div class="footer-txt">khalid-realestate.com | 🏛️ مرخصون من الهيئة العامة للعقار</div>
          <div class="footer-phone">📞 0568300022</div>
        </div>
      </div>
      <script>window.onload=()=>{window.print();window.close();}<\/script>
      </body></html>
    `);
    win.document.close();
  };

  const changeStatus = async (id, status) => {
    await updateDoc(doc(db,"clients",id),{clientStatus:status});
  };

  const addComment = async (id) => {
    if(!commentText.trim()) return;
    const client = clients.find(c=>c.id===id);
    const comments = [...(client?.comments||[]), { text:commentText, date:new Date().toLocaleDateString("en-GB"), time:new Date().toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}) }];
    await updateDoc(doc(db,"clients",id),{comments});
    setCommentText(""); setCommentId(null);
  };

  const exportClients = () => {
    const headers = ["#","الاسم","الجوال","التصنيف","الحي","الميزانية","نوع العقار (مالك)","نوع الاستثمار","طريقة الدفع","الحالة","تاريخ التواصل","تاريخ التسجيل","ملاحظات"];
    const rows = clients.map(c=>[c.clientNo||"",c.name||"",c.phone||"",c.clientType||"",c.area||"",c.budget||"",c.ownerPropertyType||"",c.investorType||"",c.paymentType||"",c.clientStatus||"معلق",c.contactDate||"",c.createdDate||"",c.notes||""]);
    const csv = [headers,...rows].map(r=>r.map(c=>`"${String(c).replace(/"/g,'""')}"`).join(",")).join("\n");
    const blob = new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"});
    const a = document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="عملاء.csv"; a.click();
  };

  const f = key => e => setForm(p=>({...p,[key]:e.target.value}));
  const IST = { width:"100%", boxSizing:"border-box", background:"#f0f4fc", border:"1px solid rgba(74,158,255,.2)", borderRadius:10, padding:"9px 12px", color:"#1e3a7a", fontFamily:"'Cairo',sans-serif", fontSize:13 };
  const Lbl = ({c}) => <div style={{fontSize:11,color:"#5a6a90",marginBottom:5,fontWeight:600}}>{c}</div>;

  const clientTypeColor = {"مستأجر":"#4ade80","مشتري":"#fbbf24","مالك":"#93c5fd","مستثمر":"#e879f9"};
  const requestColor = {"إيجار":"#4ade80","شراء":"#fbbf24","إيجار وشراء":"#e879f9"};
  const statusColor = {"معلق":"#fbbf24","مغلق":"#f87171"};

  const [search, setSearch] = useState("");
  const [typeFilter, setTypeFilter] = useState("الكل");
  const [statusFilter, setStatusFilter] = useState("الكل");

  const myClients = isManager ? clients : clients.filter(c=>c.assignedTo===currentUser?.username);

  const filtered = myClients.filter(c=>
    (typeFilter==="الكل"||c.clientType===typeFilter) &&
    (statusFilter==="الكل"||(c.clientStatus||"معلق")===statusFilter) &&
    (c.name?.includes(search)||c.phone?.includes(search)||c.area?.includes(search)||String(c.clientNo||"").includes(search)||search==="")
  );

  return (
    <div style={{paddingTop:96,minHeight:"100vh",background:T.bg}}>
      <div style={{background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",padding:"28px 24px 24px"}}>
        <div style={{maxWidth:1100,margin:"0 auto",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
          <div>
            <div style={{fontWeight:900,fontSize:22,color:"#fff",marginBottom:3}}>👥 سجل العملاء المحتملين</div>
            <div style={{fontSize:12,color:"rgba(255,255,255,.5)"}}>{isManager?`Khalid M. A. Ghafour Al-Shaikh Est. — إجمالي: ${clients.length} عميل`:`${currentUser?.displayName||""} — عملائي: ${myClients.length}`}</div>
          </div>
          <div style={{display:"flex",gap:8}}>
            {isManager&&<button onClick={exportClients} style={{background:"#16a34a22",border:"1px solid #16a34a44",color:"#4ade80",borderRadius:11,padding:"10px 18px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>📊 تصدير Excel</button>}
            <button onClick={openAdd} style={{background:"rgba(255,255,255,.15)",border:"1px solid rgba(255,255,255,.3)",color:"#fff",borderRadius:11,padding:"10px 22px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>+ إضافة عميل</button>
          </div>
        </div>
      </div>

      <div style={{maxWidth:1100,margin:"0 auto",padding:"20px 22px 0"}}>
        {/* Stats - clickable filters */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:9,marginBottom:16}}>
          {[
            {l:"الكل",v:myClients.length,i:"👥",c:"#93c5fd",key:"الكل"},
            {l:"مستأجرين",v:myClients.filter(c=>c.clientType==="مستأجر").length,i:"🏠",c:"#4ade80",key:"مستأجر"},
            {l:"مشترين",v:myClients.filter(c=>c.clientType==="مشتري").length,i:"💰",c:"#fbbf24",key:"مشتري"},
            {l:"ملاك",v:myClients.filter(c=>c.clientType==="مالك").length,i:"🔑",c:"#93c5fd",key:"مالك"},
            {l:"مستثمرين",v:myClients.filter(c=>c.clientType==="مستثمر").length,i:"📈",c:"#e879f9",key:"مستثمر"},
          ].map((s,i)=>(
            <div key={i} onClick={()=>setTypeFilter(s.key)} style={{background:typeFilter===s.key?`linear-gradient(135deg,${s.c}15,${s.c}08)`:T.bg2,border:`1px solid ${typeFilter===s.key?s.c+"44":s.c+"18"}`,borderRadius:12,padding:"12px 10px",cursor:"pointer",transition:"all .2s",position:"relative",overflow:"hidden",boxShadow:"0 2px 8px rgba(30,58,122,.06)"}}>
              <div style={{fontSize:20,marginBottom:4}}>{s.i}</div>
              <div style={{fontWeight:900,fontSize:18,color:s.c}}>{s.v}</div>
              <div style={{fontSize:10,color:typeFilter===s.key?s.c:"#4a6fa5"}}>{s.l}</div>
              {typeFilter===s.key&&<div style={{position:"absolute",top:6,left:6,width:6,height:6,borderRadius:"50%",background:s.c}}/>}
            </div>
          ))}
        </div>

        {/* Search + status filter */}
        <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 ابحث بالاسم أو الجوال أو الحي أو رقم التسلسل..." style={{flex:1,minWidth:200,background:"#f0f4fc",border:"1px solid rgba(74,158,255,.2)",borderRadius:10,padding:"9px 14px",color:"#1e3a7a",fontFamily:"'Cairo',sans-serif",fontSize:13}}/>
          <div style={{display:"flex",gap:6}}>
            {["الكل","معلق","مغلق"].map(s=>(
              <button key={s} onClick={()=>setStatusFilter(s)} style={{padding:"8px 14px",borderRadius:20,cursor:"pointer",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:12,background:statusFilter===s?(s==="مغلق"?"#f8717133":"#fbbf2433"):"#071840",color:statusFilter===s?(s==="مغلق"?"#f87171":"#fbbf24"):"#4a6fa5",border:statusFilter===s?`1px solid ${s==="مغلق"?"#f8717144":"#fbbf2444"}`:"1px solid #1e3a7a"}}>{s==="الكل"?"🗂️ الكل":s==="معلق"?"⏳ معلق":"✅ مغلق"}</button>
            ))}
          </div>
        </div>
        <div style={{fontSize:11,color:"#5a6a90",marginBottom:10}}>{filtered.length} عميل</div>

        {!loaded ? <div style={{textAlign:"center",padding:40,color:"#5a6a90"}}>جاري التحميل...</div> : filtered.length===0 ? (
          <div style={{textAlign:"center",padding:60,color:"#1e3a7a"}}><div style={{fontSize:46,marginBottom:10}}>👤</div><div style={{fontSize:13}}>{search?"لا توجد نتائج للبحث":"لا يوجد عملاء بعد"}</div></div>
        ) : (
          <div style={{display:"flex",flexDirection:"column",gap:10,paddingBottom:30}}>
            {filtered.map(c=>{
              const cStatus = c.clientStatus||"معلق";
              const sc = statusColor[cStatus]||"#fbbf24";
              return (
              <div key={c.id} style={{background:darkMode?"linear-gradient(160deg,#0f1f4a,#1a2d6b)":"white",border:`1px solid ${darkMode?"rgba(74,158,255,.15)":sc+"22"}`,borderRadius:16,padding:"14px 16px",boxShadow:"0 2px 10px rgba(30,58,122,.06)"}}>
                {/* Top row: name + badges */}
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:8,marginBottom:10}}>
                  <div style={{display:"flex",alignItems:"center",gap:12}}>
                    <div style={{position:"relative"}}>
                      <div style={{width:40,height:40,borderRadius:12,background:c.isDuplicate?"#fef2f2":"rgba(74,158,255,.1)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,color:c.isDuplicate?"#ef4444":"#2a4d9b",fontSize:13,flexShrink:0,border:c.isDuplicate?"2px solid #ef444440":"none"}}>#{c.clientNo||"—"}</div>
                      {c.isDuplicate&&<div style={{position:"absolute",top:-6,right:-6,background:"#ef4444",color:"#fff",fontSize:8,fontWeight:900,borderRadius:10,padding:"1px 5px",whiteSpace:"nowrap"}}>مكرر</div>}
                    </div>
                    <div>
                      <div style={{fontWeight:900,fontSize:15,color:darkMode?"#e8eeff":"#1e3a7a"}}>{c.name}</div>
                      {(isManager||c.clientType!=="مالك") ? (
                        <div style={{fontSize:12,color:darkMode?"#7ab8ff":"#5a6a90"}}>📞 {c.phone}</div>
                      ) : (
                        <div style={{fontSize:11,color:"#ef4444"}}>🔒 جوال المالك محجوب</div>
                      )}
                    </div>
                  </div>
                  <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
                    {c.clientType&&<span style={{background:clientTypeColor[c.clientType]+"20",color:clientTypeColor[c.clientType],border:`1px solid ${clientTypeColor[c.clientType]}40`,borderRadius:20,padding:"3px 11px",fontSize:11,fontWeight:700}}>{c.clientType}</span>}
                    <span style={{background:statusColor[c.clientStatus||"معلق"]+"20",color:statusColor[c.clientStatus||"معلق"],border:`1px solid ${statusColor[c.clientStatus||"معلق"]}40`,borderRadius:20,padding:"3px 11px",fontSize:11,fontWeight:700}}>{c.clientStatus||"معلق"}</span>
                    <select value={cStatus} onChange={e=>changeStatus(c.id,e.target.value)} style={{background:darkMode?"#0a1538":"#f0f4fc",border:"1px solid rgba(74,158,255,.2)",color:darkMode?"#7ab8ff":"#2a4d9b",borderRadius:8,padding:"3px 8px",fontSize:11,cursor:"pointer",fontFamily:"'Cairo',sans-serif"}}>
                      <option>معلق</option><option>مغلق</option>
                    </select>
                    {c.requestType&&<span style={{background:requestColor[c.requestType]+"20",color:requestColor[c.requestType],border:`1px solid ${requestColor[c.requestType]}40`,borderRadius:20,padding:"3px 11px",fontSize:11,fontWeight:700}}>{c.requestType}</span>}
                    {c.paymentType&&<span style={{background:"rgba(74,158,255,.1)",color:"#2a4d9b",border:"1px solid rgba(74,158,255,.25)",borderRadius:20,padding:"3px 11px",fontSize:11,fontWeight:700}}>{c.paymentType}</span>}
                  </div>
                </div>

                {/* Info grid */}
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:8,marginBottom:10}}>
                  {c.area&&<div style={{background:darkMode?"#0a1538":"#f5f8ff",borderRadius:8,padding:"6px 10px",fontSize:11,border:`1px solid ${darkMode?"rgba(74,158,255,.1)":"rgba(74,158,255,.08)"}`}}><span style={{color:darkMode?"#7ab8ff":"#5a6a90"}}>📍 الحي: </span><span style={{color:darkMode?"#e8eeff":"#2a4d9b",fontWeight:700}}>{c.area}</span></div>}
                  {c.budget&&<div style={{background:darkMode?"#0a1538":"#f5f8ff",borderRadius:8,padding:"6px 10px",fontSize:11,border:`1px solid ${darkMode?"rgba(74,158,255,.1)":"rgba(74,158,255,.08)"}`}}><span style={{color:darkMode?"#7ab8ff":"#5a6a90"}}>💰 الميزانية: </span><span style={{color:"#16a34a",fontWeight:700}}>{Number(c.budget).toLocaleString()} ﷼</span></div>}
                  {c.ownerPropertyType&&<div style={{background:darkMode?"#0a1538":"#f5f8ff",borderRadius:8,padding:"6px 10px",fontSize:11,border:`1px solid ${darkMode?"rgba(74,158,255,.1)":"rgba(74,158,255,.08)"}`}}><span style={{color:darkMode?"#7ab8ff":"#5a6a90"}}>🏠 العقار: </span><span style={{color:darkMode?"#e8eeff":"#2a4d9b",fontWeight:700}}>{c.ownerPropertyType}</span></div>}
                  {c.investorType&&<div style={{background:darkMode?"#0a1538":"#f5f8ff",borderRadius:8,padding:"6px 10px",fontSize:11,border:`1px solid ${darkMode?"rgba(74,158,255,.1)":"rgba(74,158,255,.08)"}`}}><span style={{color:darkMode?"#7ab8ff":"#5a6a90"}}>📈 الاستثمار: </span><span style={{color:"#7c3aed",fontWeight:700}}>{c.investorType}</span></div>}
                  {c.contactDate&&<div style={{background:darkMode?"#0a1538":"#f5f8ff",borderRadius:8,padding:"6px 10px",fontSize:11,border:`1px solid ${darkMode?"rgba(74,158,255,.1)":"rgba(74,158,255,.08)"}`}}><span style={{color:darkMode?"#7ab8ff":"#5a6a90"}}>📞 تواصل: </span><span style={{color:darkMode?"#e8eeff":"#2a4d9b",fontWeight:700}}>{c.contactDate}</span></div>}
                  <div style={{background:darkMode?"#0a1538":"#f5f8ff",borderRadius:8,padding:"6px 10px",fontSize:11,border:`1px solid ${darkMode?"rgba(74,158,255,.1)":"rgba(74,158,255,.08)"}`}}><span style={{color:darkMode?"#7ab8ff":"#5a6a90"}}>📅 تسجيل: </span><span style={{color:darkMode?"#e8eeff":"#2a4d9b",fontWeight:700}}>{c.createdDate}</span></div>
                </div>

                {c.notes&&<div style={{marginTop:0,marginBottom:10,background:darkMode?"#0a1538":"#f5f8ff",borderRadius:8,padding:"6px 10px",fontSize:11,color:darkMode?"#7ab8ff":"#5a6a90",whiteSpace:"pre-line",lineHeight:1.7,border:`1px solid ${darkMode?"rgba(74,158,255,.1)":"rgba(74,158,255,.08)"}`}}>💬 {c.notes}</div>}

                {/* Comments */}
                {(c.comments||[]).length>0&&(
                  <div style={{marginBottom:10,background:darkMode?"#0a1538":"#f5f8ff",borderRadius:8,padding:"8px 10px",border:`1px solid ${darkMode?"rgba(74,158,255,.1)":"rgba(74,158,255,.08)"}`}}>
                    {(c.comments||[]).map((cm,i)=>(
                      <div key={i} style={{borderBottom:i<c.comments.length-1?`1px solid ${darkMode?"rgba(74,158,255,.1)":"rgba(74,158,255,.08)"}`:  "none",paddingBottom:i<c.comments.length-1?6:0,marginBottom:i<c.comments.length-1?6:0}}>
                        <div style={{fontSize:10,color:darkMode?"#7ab8ff":"#5a6a90",marginBottom:2}}>{cm.date} {cm.time}</div>
                        <div style={{fontSize:12,color:darkMode?"#e8eeff":"#2a4d9b"}}>{cm.text}</div>
                      </div>
                    ))}
                  </div>
                )}

                {/* Action buttons row */}
                <div style={{display:"flex",gap:6,paddingTop:10,borderTop:`1px solid ${darkMode?"rgba(74,158,255,.1)":"rgba(74,158,255,.08)"}`,flexWrap:"wrap"}}>
                  <button onClick={()=>openEdit(c)} style={{background:"rgba(74,158,255,.1)",border:"1px solid rgba(74,158,255,.2)",color:darkMode?"#7ab8ff":"#2a4d9b",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:11,fontFamily:"'Cairo',sans-serif",fontWeight:600}}>✏️ تعديل</button>
                  <button onClick={()=>printClientCard(c)} style={{background:"rgba(74,158,255,.08)",border:"1px solid rgba(74,158,255,.15)",color:darkMode?"#7ab8ff":"#2a4d9b",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:11,fontFamily:"'Cairo',sans-serif",fontWeight:600}}>🖨️ طباعة</button>
                  {isManager&&<button onClick={()=>{setAssignModal(c.id);setAssignTo(c.assignedTo||"");setAssignMsg("");}} style={{background:"rgba(122,184,255,.1)",border:"1px solid rgba(74,158,255,.2)",color:darkMode?"#7ab8ff":"#2a4d9b",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:11,fontFamily:"'Cairo',sans-serif",fontWeight:600}}>📋 تحويل</button>}
                  {isManager&&<button onClick={()=>setDelId(c.id)} style={{background:"#ef444410",border:"1px solid #ef444425",color:"#ef4444",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:11,fontFamily:"'Cairo',sans-serif",fontWeight:600}}>🗑️ حذف</button>}
                  {/* Assigned badge */}
                  {c.assignedToName&&<span style={{marginRight:"auto",display:"inline-flex",alignItems:"center",gap:4,background:"rgba(74,158,255,.08)",border:"1px solid rgba(74,158,255,.15)",borderRadius:20,padding:"4px 10px",fontSize:10,color:darkMode?"#7ab8ff":"#2a4d9b",fontWeight:600}}>👤 {c.assignedToName}</span>}
                </div>

                {/* Comment input */}
                {commentId===c.id ? (
                  <div style={{marginTop:8,display:"flex",gap:6}}>
                    <input value={commentText} onChange={e=>setCommentText(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addComment(c.id)} placeholder="اكتب تعليقك..." style={{flex:1,background:darkMode?"#0a1538":"#f0f4fc",border:"1px solid rgba(74,158,255,.2)",borderRadius:8,padding:"6px 10px",color:darkMode?"#e8eeff":"#1e3a7a",fontFamily:"'Cairo',sans-serif",fontSize:12}}/>
                    <button onClick={()=>addComment(c.id)} style={{background:"#1e3a7a",border:"none",color:"#fff",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:12,fontFamily:"'Cairo',sans-serif",fontWeight:700}}>إرسال</button>
                    <button onClick={()=>{setCommentId(null);setCommentText("");}} style={{background:darkMode?"#0a1538":"#f0f4fc",border:"1px solid rgba(74,158,255,.2)",color:darkMode?"#7ab8ff":"#5a6a90",borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:12}}>إلغاء</button>
                  </div>
                ) : (
                  <button onClick={()=>setCommentId(c.id)} style={{marginTop:8,background:darkMode?"rgba(74,158,255,.05)":"rgba(74,158,255,.04)",border:`1px solid rgba(74,158,255,.12)`,color:darkMode?"#7ab8ff":"#5a6a90",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:11,fontFamily:"'Cairo',sans-serif",width:"100%"}}>+ إضافة تعليق</button>
                )}
              </div>
            );})}
          </div>
        )}
      </div>

      {/* Add Client Modal */}
      {showForm&&(
        <div style={{position:"fixed",inset:0,background:"#000b",zIndex:600,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={()=>{setShowForm(false);setEditClientId(null);}}>
          <div style={{background:"linear-gradient(160deg,#ffffff,#f5f8ff)",border:"1px solid #1a4faa",borderRadius:22,padding:24,maxWidth:540,width:"100%",maxHeight:"93vh",overflowY:"auto"}} onClick={e=>e.stopPropagation()}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,paddingBottom:14,borderBottom:"1px solid #1e3a7a"}}>
              <div style={{fontWeight:900,fontSize:16,color:"#1e3a7a"}}>{editClientId?"✏️ تعديل العميل":"👤 إضافة عميل جديد"}</div>
              <button onClick={()=>{setShowForm(false);setEditClientId(null);}} style={{background:"#1e3a7a",border:"none",color:"#aaa",width:30,height:30,borderRadius:"50%",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:11,marginBottom:14}}>
              <div style={{gridColumn:"1/-1"}}><Lbl c="الاسم *"/><input value={form.name} onChange={f("name")} style={IST} placeholder="اسم العميل"/></div>
              <div>
                <Lbl c="رقم الجوال *"/>
                {(!isManager && form.clientType==="مالك") ? (
                  <div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:10,padding:"9px 12px",fontSize:12,color:"#ef4444",fontWeight:600}}>🔒 جوال المالك محجوب للموظف</div>
                ) : (
                  <>
                    <input value={form.phone} onChange={f("phone")} style={{...IST, borderColor: clients.find(c=>c.phone===form.phone&&c.id!==editClientId)?"#ef4444":"rgba(74,158,255,.2)"}} placeholder="05xxxxxxxx"/>
                    {form.phone&&clients.find(c=>c.phone===form.phone&&c.id!==editClientId)&&(
                      <div style={{fontSize:11,color:"#ef4444",marginTop:4,fontWeight:700}}>
                        ⚠️ مكرر! مسجل للعميل: "{clients.find(c=>c.phone===form.phone&&c.id!==editClientId)?.name}" — #{clients.find(c=>c.phone===form.phone&&c.id!==editClientId)?.clientNo}
                      </div>
                    )}
                  </>
                )}
              </div>
              <div>
                <Lbl c="تاريخ التواصل"/>
                <input type="date" value={form.contactDate} onChange={f("contactDate")} style={{...IST, colorScheme: darkMode?"dark":"light"}}/>
              </div>

              {/* Client Type */}
              <div style={{gridColumn:"1/-1"}}>
                <Lbl c="تصنيف العميل"/>
                <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
                  {["مستأجر","مشتري","مالك","مستثمر"].map(t=>(
                    <button key={t} type="button" onClick={()=>setForm(p=>({...p,clientType:t}))} style={{flex:1,minWidth:80,padding:"9px",borderRadius:9,border:`2px solid ${form.clientType===t?clientTypeColor[t]:"rgba(74,158,255,.2)"}`,background:form.clientType===t?clientTypeColor[t]+"22":"transparent",color:form.clientType===t?clientTypeColor[t]:"#5a6a90",cursor:"pointer",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,transition:"all .15s"}}>{t}</button>
                  ))}
                </div>
              </div>

              {/* مستأجر - نوع الطلب */}
              {form.clientType==="مستأجر"&&<div style={{gridColumn:"1/-1"}}><Lbl c="نوع الإيجار المطلوب"/><select value={form.requestType} onChange={f("requestType")} style={IST}><option>سكني</option><option>تجاري</option><option>مفروش</option></select></div>}

              {/* مشتري - طريقة الدفع */}
              {form.clientType==="مشتري"&&(
                <div style={{gridColumn:"1/-1"}}>
                  <Lbl c="طريقة الدفع"/>
                  <div style={{display:"flex",gap:8}}>
                    {["كاش","تمويل بنكي","كاش وتمويل"].map(opt=>(
                      <button key={opt} type="button" onClick={()=>setForm(p=>({...p,paymentType:opt}))} style={{flex:1,padding:"8px",borderRadius:9,border:`1px solid ${form.paymentType===opt?"#2563c7":"#1e3a7a"}`,background:form.paymentType===opt?"#1a4faa":"#071840",color:form.paymentType===opt?"#fff":"#4a6fa5",cursor:"pointer",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:12}}>{opt}</button>
                    ))}
                  </div>
                </div>
              )}

              {/* مالك - نوع الملكية */}
              {form.clientType==="مالك"&&(
                <div style={{gridColumn:"1/-1"}}>
                  <Lbl c="نوع العقار المملوك"/>
                  <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
                    {["شقة","فيلا","عمارة","أرض","محل تجاري","مكتب"].map(t=>(
                      <button key={t} type="button" onClick={()=>setForm(p=>({...p,ownerPropertyType:t}))} style={{padding:"7px 12px",borderRadius:9,border:`1px solid ${form.ownerPropertyType===t?"#60a5fa":"#1e3a7a"}`,background:form.ownerPropertyType===t?"#1e40af":"#071840",color:form.ownerPropertyType===t?"#fff":"#4a6fa5",cursor:"pointer",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:12}}>{t}</button>
                    ))}
                  </div>
                </div>
              )}

              {/* مستثمر - نوع الاستثمار */}
              {form.clientType==="مستثمر"&&(
                <div style={{gridColumn:"1/-1"}}>
                  <Lbl c="نوع الاستثمار المطلوب"/>
                  <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
                    {["أراضي","عمائر","شقق","محطات وقود","محلات تجارية","مشاريع متعددة"].map(t=>(
                      <button key={t} type="button" onClick={()=>setForm(p=>({...p,investorType:t}))} style={{padding:"7px 12px",borderRadius:9,border:`1px solid ${form.investorType===t?"#e879f9":"#1e3a7a"}`,background:form.investorType===t?"#7e22ce":"#071840",color:form.investorType===t?"#fff":"#4a6fa5",cursor:"pointer",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:12}}>{t}</button>
                    ))}
                  </div>
                </div>
              )}

              <div><Lbl c="الحي / الموقع المطلوب"/><input value={form.area} onChange={f("area")} style={IST} placeholder="مثال: العزيزية، الخبر"/></div>
              <div style={{gridColumn:"1/-1"}}><Lbl c="الميزانية (﷼)"/><input type="number" value={form.budget} onChange={f("budget")} style={IST} placeholder="مثال: 500000"/></div>
              <div style={{gridColumn:"1/-1"}}><Lbl c="ملاحظات"/><textarea value={form.notes} onChange={f("notes")} rows={3} style={{...IST,resize:"none"}} placeholder="أي تفاصيل إضافية..."/></div>
            </div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={save} style={{flex:1,background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",color:"#fff",border:"none",borderRadius:11,padding:"12px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer"}}>{editClientId?"💾 حفظ التعديلات":"✅ حفظ العميل"}</button>
              <button onClick={()=>{setShowForm(false);setEditClientId(null);}} style={{background:"#f0f4fc",border:"1px solid rgba(74,158,255,.2)",color:"#5a6a90",borderRadius:11,padding:"12px 18px",fontFamily:"'Cairo',sans-serif",cursor:"pointer"}}>إلغاء</button>
            </div>
          </div>
        </div>
      )}

      {/* Assign Modal */}
      {assignModal&&(
        <div style={{position:"fixed",inset:0,background:"#000c",zIndex:700,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={()=>setAssignModal(null)}>
          <div style={{background:darkMode?"linear-gradient(160deg,#0f1f4a,#1a2d6b)":"white",border:"1px solid rgba(74,158,255,.25)",borderRadius:20,padding:24,maxWidth:420,width:"100%"}} onClick={e=>e.stopPropagation()}>
            <div style={{fontWeight:900,fontSize:16,color:darkMode?"#e8eeff":"#1e3a7a",marginBottom:4}}>📋 تحويل العميل</div>
            <div style={{fontSize:12,color:darkMode?"#7ab8ff":"#5a6a90",marginBottom:20,paddingBottom:14,borderBottom:"1px solid rgba(74,158,255,.1)"}}>{clients.find(c=>c.id===assignModal)?.name}</div>
            <div style={{marginBottom:14}}>
              <div style={{fontSize:11,color:darkMode?"#7ab8ff":"#5a6a90",marginBottom:8,fontWeight:600}}>اختر الموظف</div>
              <div style={{display:"flex",flexDirection:"column",gap:8}}>
                {employees.map(emp=>(
                  <button key={emp.username} onClick={()=>setAssignTo(emp.username)} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",borderRadius:10,border:`2px solid ${assignTo===emp.username?"#4a9eff":"rgba(74,158,255,.2)"}`,background:assignTo===emp.username?"rgba(74,158,255,.1)":"transparent",color:darkMode?"#e8eeff":"#1e3a7a",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer",textAlign:"right",width:"100%"}}>
                    <div style={{width:34,height:34,borderRadius:9,background:"rgba(74,158,255,.12)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>👤</div>
                    <div style={{flex:1}}>
                      <div>{emp.displayName}</div>
                      <div style={{fontSize:10,color:darkMode?"rgba(122,184,255,.6)":"#8899bb",fontWeight:400}}>{emp.username}</div>
                    </div>
                    {assignTo===emp.username&&<span style={{color:"#4a9eff",fontSize:18}}>✓</span>}
                  </button>
                ))}
              </div>
            </div>
            <div style={{marginBottom:16}}>
              <div style={{fontSize:11,color:darkMode?"#7ab8ff":"#5a6a90",marginBottom:6,fontWeight:600}}>رسالة للموظف (اختياري)</div>
              <textarea value={assignMsg} onChange={e=>setAssignMsg(e.target.value)} rows={3} placeholder="مثال: تواصل مع العميل بخصوص شقة في العليا..." style={{width:"100%",boxSizing:"border-box",background:darkMode?"#0a1538":"#f5f8ff",border:"1px solid rgba(74,158,255,.2)",borderRadius:10,padding:"9px 12px",color:darkMode?"#e8eeff":"#1e3a7a",fontFamily:"'Cairo',sans-serif",fontSize:13,resize:"none"}}/>
            </div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={assignClient} disabled={!assignTo} style={{flex:1,background:assignTo?"linear-gradient(135deg,#1e3a7a,#2a4d9b)":"#ccc",color:"white",border:"none",borderRadius:11,padding:"12px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:assignTo?"pointer":"not-allowed"}}>📋 تحويل العميل</button>
              <button onClick={()=>setAssignModal(null)} style={{background:darkMode?"#0a1538":"#f5f8ff",border:"1px solid rgba(74,158,255,.2)",color:darkMode?"#7ab8ff":"#5a6a90",borderRadius:11,padding:"12px 18px",fontFamily:"'Cairo',sans-serif",cursor:"pointer"}}>إلغاء</button>
            </div>
          </div>
        </div>
      )}

      {delId&&(
        <div style={{position:"fixed",inset:0,background:"#000c",zIndex:700,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:"linear-gradient(160deg,#ffffff,#f5f8ff)",border:"1px solid #ef444440",borderRadius:20,padding:28,maxWidth:300,textAlign:"center"}}>
            <div style={{fontSize:36,marginBottom:10}}>⚠️</div>
            <div style={{fontWeight:900,fontSize:15,color:"#1e3a7a",marginBottom:6}}>حذف العميل</div>
            <div style={{color:"#5a6a90",marginBottom:20,fontSize:13}}>سيتم الحذف نهائياً</div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>del(delId)} style={{flex:1,background:"#ef4444",color:"#fff",border:"none",borderRadius:11,padding:"10px",fontFamily:"'Cairo',sans-serif",fontWeight:700,cursor:"pointer"}}>احذف</button>
              <button onClick={()=>setDelId(null)} style={{flex:1,background:"#f0f4fc",border:"1px solid rgba(74,158,255,.2)",color:"#5a6a90",borderRadius:11,padding:"10px",fontFamily:"'Cairo',sans-serif",cursor:"pointer"}}>إلغاء</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Providers Page ─────────────────────────────────────────────────────────────
function ProvidersPage({ lang, T, darkMode }) {
  const [providers, setProviders] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [editId, setEditId] = useState(null);
  const [delId, setDelId] = useState(null);
  const [showWork, setShowWork] = useState(null);
  const [workText, setWorkText] = useState("");
  const emptyForm = { name:"", phone:"", specialty:"تكييف", rating:"ممتاز", notes:"" };
  const [form, setForm] = useState(emptyForm);

  const SPECIALTIES = ["تكييف","نجارة","سباكة","كهرباء","دهانات","تنظيف","حراسة","صيانة عامة","بناء","زجاج","ألمنيوم","أجهزة كهربائية"];
  const RATINGS = ["ممتاز","وسط","مقبول"];
  const ratingColor = {"ممتاز":"#4ade80","وسط":"#fbbf24","مقبول":"#f87171"};

  useEffect(()=>{
    const unsub = onSnapshot(collection(db,"providers"), snap=>{
      setProviders(snap.docs.map(d=>({id:d.id,...d.data()})));
      setLoaded(true);
    }, ()=>setLoaded(true));
    return ()=>unsub();
  },[]);

  const exportProviders = () => {
    const headers = ["الاسم","الجوال","التخصص","التقييم","ملاحظات","عدد الأعمال"];
    const rows = providers.map(p=>[p.name||"",p.phone||"",p.specialty||"",p.rating||"",(p.notes||""),(p.works||[]).length]);
    const csv = [headers,...rows].map(r=>r.map(c=>`"${String(c).replace(/"/g,'""')}"`).join(",")).join("\n");
    const blob = new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"});
    const a = document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="مزودو-الخدمات.csv"; a.click();
  };

  const save = async () => {
    if(!form.name||!form.phone) return alert("الاسم والجوال مطلوبان");
    if(editId) {
      await updateDoc(doc(db,"providers",editId),{...form});
    } else {
      await addDoc(collection(db,"providers"),{...form, works:[], createdAt:new Date().toISOString()});
    }
    setForm(emptyForm); setShowForm(false); setEditId(null);
  };

  const del = async (id) => { await deleteDoc(doc(db,"providers",id)); setDelId(null); };

  const printProviderCard = (p) => {
    const ratingColor = {"ممتاز":"#16a34a","وسط":"#b45309","مقبول":"#ef4444"};
    const win = window.open('','_blank','width=420,height=580');
    win.document.write(`
      <!DOCTYPE html><html dir="rtl" lang="ar">
      <head><meta charset="UTF-8">
      <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet">
      <style>
        *{box-sizing:border-box;margin:0;padding:0;}
        body{font-family:'Cairo',sans-serif;background:#f5f8ff;display:flex;justify-content:center;align-items:center;min-height:100vh;padding:20px;}
        .card{background:white;border-radius:16px;padding:28px;width:360px;box-shadow:0 4px 24px rgba(30,58,122,.12);border:1px solid rgba(74,158,255,.15);}
        .header{background:linear-gradient(135deg,#1e3a7a,#2a4d9b);border-radius:12px;padding:18px;margin-bottom:18px;display:flex;align-items:center;gap:14px;}
        .logo{width:50px;height:50px;border-radius:10px;background:white;overflow:hidden;flex-shrink:0;}
        .logo img{width:100%;height:100%;object-fit:contain;}
        .brand-n{font-size:11px;font-weight:900;color:white;line-height:1.3;}
        .brand-s{font-size:9px;color:rgba(255,255,255,.55);}
        .provider-info{display:flex;align-items:center;gap:14px;margin-bottom:16px;}
        .prov-icon{width:52px;height:52px;border-radius:14px;background:rgba(74,158,255,.08);border:1px solid rgba(74,158,255,.15);display:flex;align-items:center;justify-content:center;font-size:26px;flex-shrink:0;}
        .prov-name{font-size:20px;font-weight:900;color:#1e3a7a;}
        .prov-phone{font-size:13px;color:#5a6a90;margin-top:3px;}
        .badges{display:flex;gap:8px;margin-top:8px;flex-wrap:wrap;}
        .badge{border-radius:20px;padding:4px 14px;font-size:11px;font-weight:700;}
        .divider{height:1px;background:rgba(74,158,255,.1);margin:14px 0;}
        .info-row{display:flex;justify-content:space-between;background:#f5f8ff;border-radius:8px;padding:10px 12px;margin-bottom:10px;}
        .info-lbl{font-size:10px;color:#8899bb;}
        .info-val{font-size:12px;font-weight:700;color:#1e3a7a;}
        .works-title{font-size:11px;font-weight:700;color:#1e3a7a;margin-bottom:8px;}
        .work-item{font-size:11px;color:#5a6a90;padding:5px 0;border-bottom:1px solid #f0f4fc;display:flex;justify-content:space-between;}
        .footer-bar{background:#f0f4fc;border-radius:10px;padding:12px 14px;display:flex;justify-content:space-between;align-items:center;margin-top:14px;}
        .footer-txt{font-size:10px;color:#8899bb;}
        .footer-phone{font-size:12px;font-weight:700;color:#1e3a7a;}
        @media print{body{background:white;padding:0;}.card{box-shadow:none;}}
      </style>
      </head><body>
      <div class="card">
        <div class="header">
          <div class="logo"><img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1778958489/WhatsApp_Image_2026-05-16_at_9.59.47_PM_zhmw6y.jpg"/></div>
          <div>
            <div class="brand-n">مؤسسة خالد محمد عبدالغفور الشيخ</div>
            <div class="brand-s">Khalid M. A. Ghafour Al-Shaikh Est.</div>
          </div>
        </div>
        <div class="provider-info">
          <div class="prov-icon">🔧</div>
          <div>
            <div class="prov-name">${p.name||""}</div>
            <div class="prov-phone">📞 ${p.phone||""}</div>
            <div class="badges">
              <span class="badge" style="background:rgba(74,158,255,.1);color:#2a4d9b;border:1px solid rgba(74,158,255,.25)">${p.specialty||""}</span>
              <span class="badge" style="background:${ratingColor[p.rating]||"#2a4d9b"}18;color:${ratingColor[p.rating]||"#2a4d9b"};border:1px solid ${ratingColor[p.rating]||"#2a4d9b"}33">${p.rating||""}</span>
            </div>
          </div>
        </div>
        <div class="divider"></div>
        ${p.notes?`<div class="info-row"><div><div class="info-lbl">📝 ملاحظات</div><div class="info-val">${p.notes}</div></div></div>`:""}
        ${p.works&&p.works.length>0?`
          <div class="works-title">🛠️ الأعمال المنجزة (${p.works.length})</div>
          ${p.works.slice(0,5).map(w=>`<div class="work-item"><span>${w.text}</span><span style="color:#8899bb;font-size:10px">${w.date}</span></div>`).join("")}
          ${p.works.length>5?`<div style="font-size:10px;color:#8899bb;text-align:center;margin-top:6px">و ${p.works.length-5} أعمال أخرى...</div>`:""}
        `:""}
        <div class="footer-bar">
          <div class="footer-txt">khalid-realestate.com | 🏛️ مرخصون</div>
          <div class="footer-phone">📞 0568300022</div>
        </div>
      </div>
      <script>window.onload=()=>{window.print();window.close();}<\/script>
      </body></html>
    `);
    win.document.close();
  };

  const addWork = async (id) => {
    if(!workText.trim()) return;
    const p = providers.find(p=>p.id===id);
    const works = [...(p?.works||[]), { text:workText, date:new Date().toLocaleDateString("en-GB") }];
    await updateDoc(doc(db,"providers",id),{works});
    setWorkText(""); setShowWork(null);
  };

  const f = key => e => setForm(p=>({...p,[key]:e.target.value}));
  const IST = { width:"100%", boxSizing:"border-box", background:darkMode?"#0f1f4a":"#f5f8ff", border:`1px solid rgba(74,158,255,.25)`, borderRadius:10, padding:"9px 12px", color:T.text, fontFamily:"'Cairo',sans-serif", fontSize:13 };
  const Lbl = ({c}) => <div style={{fontSize:11,color:T.text3,marginBottom:5,fontWeight:600}}>{c}</div>;

  return (
    <div style={{paddingTop:96,minHeight:"100vh",background:T.bg}}>
      <div style={{background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",padding:"28px 24px 24px"}}>
        <div style={{maxWidth:1100,margin:"0 auto",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
          <div>
            <div style={{fontWeight:900,fontSize:22,color:"#fff",marginBottom:3}}>🔧 مزودو الخدمات</div>
            <div style={{fontSize:12,color:"rgba(255,255,255,.5)"}}>للمدير فقط</div>
          </div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={exportProviders} style={{background:"#16a34a22",border:"1px solid #16a34a44",color:"#4ade80",borderRadius:11,padding:"10px 18px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>📊 تصدير Excel</button>
            <button onClick={()=>{setForm(emptyForm);setEditId(null);setShowForm(true);}} style={{background:"rgba(255,255,255,.15)",border:"1px solid rgba(255,255,255,.3)",color:"#fff",borderRadius:11,padding:"10px 22px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>+ إضافة مزود</button>
          </div>
        </div>
      </div>

      {/* Stats */}
      <div style={{maxWidth:1100,margin:"0 auto",padding:"20px 22px 0"}}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:9,marginBottom:20}}>
          {[
            {l:"الإجمالي",v:providers.length,i:"🔧",c:"#7ab8ff"},
            {l:"ممتاز",v:providers.filter(p=>p.rating==="ممتاز").length,i:"⭐",c:"#4ade80"},
            {l:"وسط",v:providers.filter(p=>p.rating==="وسط").length,i:"👍",c:"#fbbf24"},
            {l:"مقبول",v:providers.filter(p=>p.rating==="مقبول").length,i:"⚠️",c:"#f87171"},
          ].map((s,i)=>(
            <div key={i} style={{background:darkMode?"linear-gradient(135deg,#0f1f4a,#1a2d6b)":"white",border:`1px solid ${s.c}22`,borderRadius:12,padding:"12px 10px",boxShadow:"0 2px 10px rgba(30,58,122,.07)"}}>
              <div style={{fontSize:20,marginBottom:4}}>{s.i}</div>
              <div style={{fontWeight:900,fontSize:18,color:s.c}}>{s.v}</div>
              <div style={{fontSize:10,color:T.text3}}>{s.l}</div>
            </div>
          ))}
        </div>

        {!loaded ? <div style={{textAlign:"center",padding:40,color:T.text3}}>جاري التحميل...</div> : providers.length===0 ? (
          <div style={{textAlign:"center",padding:60,color:T.text3}}><div style={{fontSize:46,marginBottom:10}}>🔧</div><div style={{fontSize:13}}>لا يوجد مزودو خدمات بعد</div></div>
        ) : (
          <div style={{display:"flex",flexDirection:"column",gap:10,paddingBottom:30}}>
            {providers.map(p=>(
              <div key={p.id} style={{background:darkMode?"linear-gradient(160deg,#0f1f4a,#1a2d6b)":"white",border:`1px solid rgba(74,158,255,.15)`,borderRadius:16,padding:"14px 16px",boxShadow:"0 2px 10px rgba(30,58,122,.06)"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:8,marginBottom:10}}>
                  <div style={{display:"flex",alignItems:"center",gap:12}}>
                    <div style={{width:44,height:44,borderRadius:12,background:"rgba(74,158,255,.1)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>🔧</div>
                    <div>
                      <div style={{fontWeight:900,fontSize:15,color:T.text}}>{p.name}</div>
                      <div style={{fontSize:12,color:T.text3}}>📞 {p.phone}</div>
                    </div>
                  </div>
                  <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
                    <span style={{background:"rgba(74,158,255,.1)",color:"#4a9eff",border:"1px solid rgba(74,158,255,.25)",borderRadius:20,padding:"3px 12px",fontSize:11,fontWeight:700}}>{p.specialty}</span>
                    <span style={{background:ratingColor[p.rating]+"20",color:ratingColor[p.rating],border:`1px solid ${ratingColor[p.rating]}40`,borderRadius:20,padding:"3px 12px",fontSize:11,fontWeight:700}}>{p.rating}</span>
                    <button onClick={()=>{setForm({name:p.name,phone:p.phone,specialty:p.specialty,rating:p.rating,notes:p.notes||""});setEditId(p.id);setShowForm(true);}} style={{background:"rgba(74,158,255,.1)",border:"1px solid rgba(74,158,255,.25)",color:"#4a9eff",borderRadius:8,padding:"4px 10px",cursor:"pointer",fontSize:11}}>✏️</button>
                    <button onClick={()=>printProviderCard(p)} style={{background:"rgba(74,158,255,.1)",border:"1px solid rgba(74,158,255,.25)",color:"#2a4d9b",borderRadius:8,padding:"4px 10px",cursor:"pointer",fontSize:11}}>🖨️</button>
                    <button onClick={()=>setDelId(p.id)} style={{background:"#ef444414",border:"1px solid #ef444428",color:"#f87171",borderRadius:8,padding:"4px 10px",cursor:"pointer",fontSize:11}}>🗑️</button>
                  </div>
                </div>

                {p.notes&&<div style={{fontSize:11,color:T.text3,background:darkMode?"#0a1538":"#f5f8ff",borderRadius:8,padding:"6px 10px",marginBottom:10}}>📝 {p.notes}</div>}

                {/* Works */}
                {(p.works||[]).length>0&&(
                  <div style={{background:darkMode?"#0a1538":"#f5f8ff",borderRadius:10,padding:"10px 12px",marginBottom:8}}>
                    <div style={{fontSize:11,color:T.text3,fontWeight:700,marginBottom:6}}>🛠️ الأعمال المنجزة:</div>
                    {(p.works||[]).map((w,i)=>(
                      <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:i<p.works.length-1?`1px solid rgba(74,158,255,.1)`:"none",paddingBottom:i<p.works.length-1?6:0,marginBottom:i<p.works.length-1?6:0}}>
                        <div style={{fontSize:12,color:T.text}}>• {w.text}</div>
                        <div style={{fontSize:10,color:T.text3,flexShrink:0,marginRight:8}}>{w.date}</div>
                      </div>
                    ))}
                  </div>
                )}

                {/* Add work */}
                {showWork===p.id ? (
                  <div style={{display:"flex",gap:6}}>
                    <input value={workText} onChange={e=>setWorkText(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addWork(p.id)} placeholder="اكتب العمل المنجز..." style={{flex:1,background:darkMode?"#0a1538":"#f5f8ff",border:"1px solid rgba(74,158,255,.25)",borderRadius:8,padding:"6px 10px",color:T.text,fontFamily:"'Cairo',sans-serif",fontSize:12}}/>
                    <button onClick={()=>addWork(p.id)} style={{background:"#1e3a7a",border:"none",color:"#fff",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:12,fontFamily:"'Cairo',sans-serif",fontWeight:700}}>إضافة</button>
                    <button onClick={()=>{setShowWork(null);setWorkText("");}} style={{background:darkMode?"#0a1538":"#f5f8ff",border:"1px solid rgba(74,158,255,.2)",color:T.text3,borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:12}}>إلغاء</button>
                  </div>
                ) : (
                  <button onClick={()=>setShowWork(p.id)} style={{background:darkMode?"rgba(74,158,255,.08)":"rgba(74,158,255,.06)",border:"1px solid rgba(74,158,255,.2)",color:"#4a9eff",borderRadius:8,padding:"5px 12px",cursor:"pointer",fontSize:11,fontFamily:"'Cairo',sans-serif",width:"100%"}}>+ إضافة عمل منجز</button>
                )}
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Add/Edit Modal */}
      {showForm&&(
        <div style={{position:"fixed",inset:0,background:"#000b",zIndex:600,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={()=>{setShowForm(false);setEditId(null);}}>
          <div style={{background:darkMode?"linear-gradient(160deg,#0f1f4a,#1a2d6b)":"white",border:"1px solid rgba(74,158,255,.25)",borderRadius:22,padding:24,maxWidth:480,width:"100%",maxHeight:"90vh",overflowY:"auto"}} onClick={e=>e.stopPropagation()}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,paddingBottom:14,borderBottom:"1px solid rgba(74,158,255,.15)"}}>
              <div style={{fontWeight:900,fontSize:16,color:T.text}}>{editId?"✏️ تعديل المزود":"🔧 إضافة مزود خدمات"}</div>
              <button onClick={()=>{setShowForm(false);setEditId(null);}} style={{background:"rgba(74,158,255,.1)",border:"none",color:T.text3,width:30,height:30,borderRadius:"50%",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:11,marginBottom:14}}>
              <div style={{gridColumn:"1/-1"}}><Lbl c="الاسم *"/><input value={form.name} onChange={f("name")} style={IST} placeholder="اسم المزود"/></div>
              <div><Lbl c="رقم الجوال *"/><input value={form.phone} onChange={f("phone")} style={IST} placeholder="05xxxxxxxx"/></div>
              <div>
                <Lbl c="التخصص"/>
                <select value={form.specialty} onChange={f("specialty")} style={IST}>
                  {SPECIALTIES.map(s=><option key={s}>{s}</option>)}
                </select>
              </div>
              <div style={{gridColumn:"1/-1"}}>
                <Lbl c="التقييم"/>
                <div style={{display:"flex",gap:8}}>
                  {RATINGS.map(r=>(
                    <button key={r} onClick={()=>setForm(p=>({...p,rating:r}))} style={{flex:1,padding:"8px",borderRadius:9,border:`1px solid ${form.rating===r?ratingColor[r]:"rgba(74,158,255,.2)"}`,background:form.rating===r?ratingColor[r]+"22":darkMode?"#0a1538":"#f5f8ff",color:form.rating===r?ratingColor[r]:T.text3,cursor:"pointer",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13}}>{r}</button>
                  ))}
                </div>
              </div>
              <div style={{gridColumn:"1/-1"}}><Lbl c="ملاحظات"/><textarea value={form.notes} onChange={f("notes")} rows={2} style={{...IST,resize:"none"}} placeholder="أي معلومات إضافية..."/></div>
            </div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={save} style={{flex:1,background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",color:"#fff",border:"none",borderRadius:11,padding:"12px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer"}}>{editId?"💾 حفظ التعديلات":"✅ حفظ المزود"}</button>
              <button onClick={()=>{setShowForm(false);setEditId(null);}} style={{background:darkMode?"#0a1538":"#f5f8ff",border:"1px solid rgba(74,158,255,.2)",color:T.text3,borderRadius:11,padding:"12px 18px",fontFamily:"'Cairo',sans-serif",cursor:"pointer"}}>إلغاء</button>
            </div>
          </div>
        </div>
      )}

      {delId&&(
        <div style={{position:"fixed",inset:0,background:"#000c",zIndex:700,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:darkMode?"linear-gradient(160deg,#0f1f4a,#1a2d6b)":"white",border:"1px solid #ef444440",borderRadius:20,padding:28,maxWidth:300,textAlign:"center"}}>
            <div style={{fontSize:36,marginBottom:10}}>⚠️</div>
            <div style={{fontWeight:900,fontSize:15,color:T.text,marginBottom:6}}>حذف المزود</div>
            <div style={{color:T.text3,marginBottom:20,fontSize:13}}>سيتم الحذف نهائياً</div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>del(delId)} style={{flex:1,background:"#ef4444",color:"#fff",border:"none",borderRadius:11,padding:"10px",fontFamily:"'Cairo',sans-serif",fontWeight:700,cursor:"pointer"}}>احذف</button>
              <button onClick={()=>setDelId(null)} style={{flex:1,background:darkMode?"#0a1538":"#f5f8ff",border:"1px solid rgba(74,158,255,.2)",color:T.text3,borderRadius:11,padding:"10px",fontFamily:"'Cairo',sans-serif",cursor:"pointer"}}>إلغاء</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function AboutPage({ lang, darkMode, T }) {
  const isEn=lang==="en";
  return (
    <div style={{paddingTop:96,minHeight:"100vh",background:T.bg}}>
      <div style={{background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",padding:"28px 24px 24px"}}><div style={{maxWidth:900,margin:"0 auto"}}><div style={{fontWeight:900,fontSize:22,color:"#fff",marginBottom:3}}>ℹ️ {isEn?"About Us":"عن المؤسسة"}</div></div></div>
      <div style={{maxWidth:900,margin:"0 auto",padding:"36px 24px"}}>

        {/* Main card */}
        <div style={{background:T.bg2,border:"1px solid rgba(74,158,255,.12)",borderRadius:20,padding:"32px",marginBottom:16,boxShadow:"0 4px 18px rgba(30,58,122,.07)"}}>
          <div style={{display:"flex",alignItems:"center",gap:18,marginBottom:24,flexWrap:"wrap"}}>
            <div style={{width:80,height:80,borderRadius:17,overflow:"hidden",flexShrink:0,boxShadow:"0 4px 20px rgba(74,158,255,.25)",border:"2px solid rgba(74,158,255,.2)"}}><img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1778958489/WhatsApp_Image_2026-05-16_at_9.59.47_PM_zhmw6y.jpg" alt="Logo" style={{width:"100%",height:"100%",objectFit:"cover"}}/></div>
            <div>
              <div style={{fontWeight:900,fontSize:20,color:T.text,marginBottom:4}}>مؤسسة خالد محمد عبدالغفور الشيخ</div>
              <div style={{fontSize:13,color:T.text3}}>Khalid M. A. Ghafour Al-Shaikh Est. | للخدمات العقارية</div>
              <div style={{marginTop:8,display:"inline-flex",alignItems:"center",gap:6,background:"rgba(22,163,74,.1)",border:"1px solid rgba(22,163,74,.25)",borderRadius:20,padding:"3px 12px"}}>
                <span style={{width:7,height:7,borderRadius:"50%",background:"#16a34a",display:"inline-block"}}/>
                <span style={{fontSize:11,color:"#16a34a",fontWeight:700}}>مرخصة من الهيئة العامة للعقار</span>
              </div>
            </div>
          </div>

          {/* About text */}
          <div style={{fontSize:14,color:T.text2,lineHeight:2,marginBottom:24}}>
            <p style={{marginBottom:14}}>
              {isEn
                ? "Founded in 1997 by Khalid Mohamed Abdulghafour Al-Shaikh, our establishment is among the earliest specialized real estate firms in the Eastern Province of Saudi Arabia. Over more than 27 years, we have built enduring trust with our clients based on transparency, professionalism, and integrity."
                : "تأسست المؤسسة عام 1997م على يد خالد محمد عبدالغفور الشيخ، لتكون من أوائل المؤسسات العقارية المتخصصة في المنطقة الشرقية بالمملكة العربية السعودية. على مدى أكثر من 27 عاماً، بنينا ثقة راسخة مع عملائنا قائمة على الشفافية والمهنية والنزاهة."
              }
            </p>
            <p style={{marginBottom:14}}>
              {isEn
                ? "We provide a comprehensive range of licensed real estate services including residential and commercial sales and leasing, property asset management, certified real estate valuation, and specialized investment consulting."
                : "نقدم طيفاً شاملاً من الخدمات العقارية المرخصة من الهيئة العامة للعقار، تشمل بيع وتأجير العقارات السكنية والتجارية، وإدارة الأصول العقارية، والتقييم العقاري المعتمد، والاستشارات الاستثمارية المتخصصة."
              }
            </p>
            <p>
              {isEn
                ? "Our vision is to be the most trusted real estate partner in the Eastern Province, by delivering innovative solutions that meet and exceed our clients' aspirations."
                : "رؤيتنا أن نكون الشريك العقاري الأول والأكثر موثوقية في المنطقة الشرقية، من خلال تقديم حلول عقارية مبتكرة تلبي تطلعات عملائنا وتتجاوز توقعاتهم."
              }
            </p>
          </div>

          {/* Stats */}
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:24}}>
            {[
              {n:"1997",l:isEn?"Est. Year":"سنة التأسيس",i:"🏛️"},
              {n:"27+",l:isEn?"Years Experience":"سنة خبرة",i:"⭐"},
              {n:"8+",l:isEn?"Services":"خدمة عقارية",i:"🏢"},
            ].map((s,i)=>(
              <div key={i} style={{background:T.bg3,borderRadius:12,padding:"14px",textAlign:"center",border:"1px solid rgba(74,158,255,.12)"}}>
                <div style={{fontSize:22,marginBottom:4}}>{s.i}</div>
                <div style={{fontWeight:900,fontSize:20,color:T.text,marginBottom:2}}>{s.n}</div>
                <div style={{fontSize:11,color:T.text3}}>{s.l}</div>
              </div>
            ))}
          </div>

          {/* Contact grid */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            {[["📞 "+PHONE,PHONE],["💬 WhatsApp","0568300022"],["🏛️ "+(isEn?"License":"الترخيص"),isEn?"Real Estate Gen. Authority":"هيئة العقار"],["📍 "+(isEn?"Location":"الموقع"),isEn?"Eastern Province, KSA":"المنطقة الشرقية، المملكة العربية السعودية"]].map(([l,v])=>(
              <div key={l} style={{background:T.bg3,borderRadius:11,padding:"13px 15px",border:"1px solid rgba(74,158,255,.12)"}}><div style={{fontSize:11,color:T.text3,marginBottom:4}}>{l}</div><div style={{fontSize:12,fontWeight:700,color:T.text2}}>{v}</div></div>
            ))}
          </div>
        </div>

        {/* CTA */}
        <div style={{background:"#1e3a7a",borderRadius:15,padding:"22px",textAlign:"center"}}>
          <div style={{fontWeight:900,fontSize:15,color:"#fff",marginBottom:6}}>{isEn?"Get in Touch":"تواصل معنا"}</div>
          <div style={{fontSize:12,color:"rgba(255,255,255,.5)",marginBottom:16}}>{isEn?"We're here to help 24/7":"فريقنا جاهز لمساعدتك على مدار الساعة"}</div>
          <div style={{display:"flex",gap:10,justifyContent:"center",flexWrap:"wrap"}}>
            <a href={`tel:${PHONE}`} style={{display:"inline-flex",alignItems:"center",gap:6,background:"#4a9eff",color:"#fff",borderRadius:11,padding:"10px 20px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14}}>📞 {PHONE}</a>
            <a href={`https://wa.me/${WA_NUMBER}`} target="_blank" rel="noopener noreferrer" style={{display:"inline-flex",alignItems:"center",gap:7,background:"rgba(37,211,102,.15)",border:"1px solid rgba(37,211,102,.3)",color:"#25d366",borderRadius:11,padding:"10px 18px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14}}><WaIcon size={15}/> WhatsApp</a>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── ROOT ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [page,setPage]           = useState("home");
  const [lang,setLang]           = useState("ar");
  const [darkMode,setDarkMode]   = useState(true);
  const [props,setProps]         = useState([]);
  const [loaded,setLoaded]       = useState(false);
  const [saving,setSaving]       = useState(false);
  const [scrolled,setScrolled]   = useState(false);
  const [isAdmin,setIsAdmin]     = useState(false);
  const [userRole,setUserRole]   = useState(null); // "admin" | "employee"
  const [currentUser,setCurrentUser] = useState(null); // {username, displayName, role}
  const [showLogin,setShowLogin] = useState(false);
  const [showForm,setShowForm]   = useState(false);
  const [editId,setEditId]       = useState(null);
  const [form,setForm]           = useState(emptyForm);
  const [lightbox,setLightbox]   = useState(null);
  const [delId,setDelId]         = useState(null);
  const [toast,setToast]         = useState(null);
  const [shareP,setShareP]       = useState(null);

  // ── Firebase real-time listener ──
  useEffect(()=>{
    const unsub = onSnapshot(collection(db,"properties"), (snapshot)=>{
      const data = snapshot.docs.map(d=>({id:d.id,...d.data()}));
      setProps(data);
      setLoaded(true);
    }, (err)=>{ console.error(err); setLoaded(true); });
    return ()=>unsub();
  },[]);

  useEffect(()=>{ const h=()=>setScrolled(window.scrollY>40); window.addEventListener("scroll",h); return ()=>window.removeEventListener("scroll",h); },[]);

  const showToast=(msg,type="ok")=>{ setToast({msg,type}); setTimeout(()=>setToast(null),3000); };
  const isEn=lang==="en";

  const openAdd=()=>{ setForm({...emptyForm,createdAt:today(),refNo:"REF-"+String(props.length+1).padStart(3,"0")}); setEditId(null); setShowForm(true); };
  const openEdit=p=>{ 
    const fd={...p,images:p.images||[],mapUrl:p.mapUrl||""};
    if(userRole==="employee") fd.ownerPhone="";
    setForm(fd); setEditId(p.id); setShowForm(true); 
  };

  const save=async()=>{
    if(!form.name) return showToast(isEn?"Enter property name":"أدخل اسم العقار","err");
    setSaving(true);
    try {
      const data={...form,updatedAt:today()};
      if(editId){
        await updateDoc(doc(db,"properties",editId),data);
        showToast(isEn?"Updated ✓":"تم التحديث ✓");
      } else {
        await addDoc(collection(db,"properties"),{...data,views:0,createdAt:today()});
        showToast(isEn?"Added ✓":"تم الإضافة ✓");
      }
    } catch(e){ showToast(isEn?"Error saving":"خطأ في الحفظ","err"); }
    setSaving(false);
    setShowForm(false);
  };

  const del=async(id)=>{
    try { await deleteDoc(doc(db,"properties",id)); showToast(isEn?"Deleted":"تم الحذف"); }
    catch { showToast(isEn?"Error":"خطأ","err"); }
    setDelId(null);
  };

  const changeStatus=async(id,status)=>{
    try { await updateDoc(doc(db,"properties",id),{status,updatedAt:today()}); showToast("✓"); }
    catch { showToast(isEn?"Error":"خطأ","err"); }
  };

  if(!loaded) return (
    <div style={{direction:"rtl",fontFamily:"'Cairo',sans-serif",minHeight:"100vh",background:"#1e3a7a",display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:14}}>
      <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet"/>
      <div style={{width:72,height:72,borderRadius:18,overflow:"hidden",border:"2px solid rgba(74,158,255,.3)",boxShadow:"0 8px 30px rgba(74,158,255,.2)"}}><img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1778958489/WhatsApp_Image_2026-05-16_at_9.59.47_PM_zhmw6y.jpg" alt="Logo" style={{width:"100%",height:"100%",objectFit:"cover"}}/></div>
      <div style={{color:"#7ab8ff",fontWeight:700,fontSize:14,fontFamily:"'Cairo',sans-serif"}}>جاري التحميل...</div>
    </div>
  );

  const T = darkMode ? {
    bg:"#0f1f4a", bg2:"#1a2d6b", bg3:"#0a1538", bg4:"#071030",
    text:"#e8eeff", text2:"#7ab8ff", text3:"#4a6aaa",
    border:"#2a4080", card:"linear-gradient(160deg,#1a2d6b,#0f1f4a)",
    navbar:"rgba(15,31,74,.98)", navbarBorder:"rgba(74,158,255,.25)",
    navText:"rgba(255,255,255,.55)", navActive:"rgba(74,158,255,.15)",
  } : {
    bg:"#f5f8ff", bg2:"#ffffff", bg3:"#edf1fb", bg4:"#e8eeff",
    text:"#1e3a7a", text2:"#2a4d9b", text3:"#5a6a90",
    border:"rgba(74,158,255,.2)", card:"linear-gradient(160deg,#ffffff,#f5f8ff)",
    navbar:"rgba(30,58,122,.98)", navbarBorder:"rgba(74,158,255,.3)",
    navText:"rgba(255,255,255,.55)", navActive:"rgba(74,158,255,.15)",
  };

  return (
    <div style={{direction:isEn?"ltr":"rtl",fontFamily:"'Cairo',sans-serif",minHeight:"100vh",background:T.bg,color:T.text,transition:"background .3s,color .3s"}}>
      <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet"/>

      {lightbox&&<Lightbox images={lightbox.images} startIndex={lightbox.idx} onClose={()=>setLightbox(null)}/>}
      {showLogin&&<LoginModal onSuccess={(user)=>{setIsAdmin(true);setUserRole(user.role);setCurrentUser(user);setShowLogin(false);showToast(user.role==="admin"?`مرحباً ${user.displayName}! 👑`:`مرحباً ${user.displayName}! 👤`);}} onClose={()=>setShowLogin(false)} lang={lang}/>}
      {showForm&&<PropForm form={form} setForm={setForm} onSave={save} onClose={()=>setShowForm(false)} editId={editId} T={T} userRole={userRole}/>}
      {shareP&&<ShareModal p={shareP} onClose={()=>setShareP(null)}/>}

      {toast&&(<div style={{position:"fixed",top:20,left:"50%",transform:"translateX(-50%)",background:toast.type==="err"?"#ef4444":"#1a4faa",border:`1px solid ${toast.type==="err"?"#ef4444":"#2563c7"}`,color:"#fff",padding:"10px 24px",borderRadius:12,zIndex:9999,fontWeight:700,fontSize:13,boxShadow:"0 8px 32px #000a",whiteSpace:"nowrap"}}>{toast.msg}</div>)}
      {saving&&(<div style={{position:"fixed",bottom:16,left:16,background:"#ddeeff",border:"1px solid rgba(74,158,255,.35)",color:"#2a4d9b",padding:"6px 13px",borderRadius:9,fontSize:11,fontWeight:600,zIndex:9998}}>💾 {isEn?"Saving...":"جاري الحفظ..."}</div>)}

      {delId&&(
        <div style={{position:"fixed",inset:0,background:"#000c",zIndex:700,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:T.card,border:"1px solid #ef444440",borderRadius:20,padding:28,maxWidth:320,textAlign:"center"}}>
            <div style={{fontSize:38,marginBottom:10}}>⚠️</div>
            <div style={{fontWeight:900,fontSize:16,color:T.text,marginBottom:6}}>{isEn?"Delete Property":"حذف العقار"}</div>
            <div style={{color:T.text3,marginBottom:20,fontSize:13}}>{isEn?"This cannot be undone.":"سيتم الحذف نهائياً."}</div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>del(delId)} style={{flex:1,background:"#ef4444",color:"#fff",border:"none",borderRadius:11,padding:"11px",fontFamily:"'Cairo',sans-serif",fontWeight:700,cursor:"pointer"}}>{isEn?"Delete":"احذف"}</button>
              <button onClick={()=>setDelId(null)} style={{flex:1,background:T.bg2,border:`1px solid ${T.border}`,color:T.text3,borderRadius:11,padding:"11px",fontFamily:"'Cairo',sans-serif",cursor:"pointer"}}>{isEn?"Cancel":"إلغاء"}</button>
            </div>
          </div>
        </div>
      )}

      <Navbar page={page} setPage={setPage} isAdmin={isAdmin} onLoginClick={()=>setShowLogin(true)} onLogout={()=>{setIsAdmin(false);setUserRole(null);setCurrentUser(null);showToast(isEn?"Logged out":"تم تسجيل الخروج");}} lang={lang} setLang={setLang} scrolled={scrolled} darkMode={darkMode} setDarkMode={setDarkMode} T={T} userRole={userRole}/>

      {page==="home"       && <HomePage setPage={setPage} lang={lang} darkMode={darkMode} T={T}/>}
      {page==="properties" && <PropertiesPage props={props} isAdmin={isAdmin} userRole={userRole} onEdit={openEdit} onDelete={id=>setDelId(id)} onChangeStatus={changeStatus} setLightbox={setLightbox} onShare={setShareP} onOpenAdd={openAdd} lang={lang} darkMode={darkMode} T={T}/>}
      {page==="services"   && <HomePage setPage={setPage} lang={lang} darkMode={darkMode} T={T}/>}
      {page==="about"      && <AboutPage lang={lang} darkMode={darkMode} T={T}/>}
      {page==="clients"    && <ClientsPage lang={lang} darkMode={darkMode} T={T} userRole={userRole} currentUser={currentUser}/>}
      {page==="providers"  && <ProvidersPage lang={lang} darkMode={darkMode} T={T}/>}

      {/* ── Floating WhatsApp Button ── */}
      <a href={`https://wa.me/${WA_NUMBER}`} target="_blank" rel="noopener noreferrer"
        style={{position:"fixed",bottom:24,left:24,zIndex:999,display:"flex",alignItems:"center",gap:10,background:"#25d366",borderRadius:50,padding:"14px 20px",boxShadow:"0 6px 24px #25d36655",textDecoration:"none",transition:"transform .2s"}}
        onMouseEnter={e=>e.currentTarget.style.transform="scale(1.08)"}
        onMouseLeave={e=>e.currentTarget.style.transform="scale(1)"}
      >
        <WaIcon size={26} color="#fff"/>
        <span style={{color:"#fff",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,whiteSpace:"nowrap"}}>{isEn?"Chat with us":"تواصل معنا"}</span>
      </a>

      <style>{`
        *{box-sizing:border-box;}
        input:focus,select:focus,textarea:focus{outline:none;border-color:#4a9eff!important;box-shadow:0 0 0 2px rgba(74,158,255,.12)!important;}
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-track{background:${darkMode?"#0a1538":"#f5f8ff"}}
        ::-webkit-scrollbar-thumb{background:${darkMode?"#2a4080":"rgba(74,158,255,.3)"};border-radius:4px}
        a{font-family:'Cairo',sans-serif;}
        @media(max-width:768px){
          .desktop-nav{display:none!important;}
          .desktop-actions{display:none!important;}
          .brand-text{display:none!important;}
        }
      `}</style>
    </div>
  );
}
