import { useState, useEffect, useCallback, useRef } from "react";
import { db } from "./firebase";
import { collection, getDocs, addDoc, updateDoc, deleteDoc, doc, onSnapshot } from "firebase/firestore";

const USERS = [{ username: "Khalil", password: "1234" }];
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

const today = () => new Date().toLocaleDateString("ar-SA");

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
  a.href=url; a.download=`عقارات-${new Date().toLocaleDateString("ar-SA").replace(/\//g,"-")}.csv`;
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
      <div style={{background:"linear-gradient(160deg,#071840,#0a1f54)",border:"1px solid #2563c7",borderRadius:20,padding:26,maxWidth:440,width:"100%"}} onClick={e=>e.stopPropagation()}>
        
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
            <img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1777727982/WhatsApp_Image_2026-04-30_at_1.38.52_AM_bx1fuy.jpg"
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
              <div style={{fontSize:15,fontWeight:900,color:"#93c5fd"}}>{p.area} م²</div>
            </div>}
            {p.rooms&&<div style={{background:"rgba(255,255,255,.07)",borderRadius:10,padding:"8px 12px"}}>
              <div style={{fontSize:10,color:"rgba(255,255,255,.4)"}}>الغرف</div>
              <div style={{fontSize:15,fontWeight:900,color:"#93c5fd"}}>{p.rooms} غرف</div>
            </div>}
          </div>

          {/* Tags */}
          <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:14,position:"relative"}}>
            <span style={{background:"#1a4faa44",color:"#93c5fd",fontSize:11,fontWeight:700,padding:"3px 10px",borderRadius:14}}>{p.type}</span>
            <span style={{background:"#1a4faa44",color:"#93c5fd",fontSize:11,fontWeight:700,padding:"3px 10px",borderRadius:14}}>{p.dealType}</span>
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
          <button onClick={downloadImage} disabled={generating} style={{display:"flex",alignItems:"center",gap:10,background:"linear-gradient(135deg,#1a4faa,#2563c7)",border:"none",color:"#fff",borderRadius:12,padding:"11px 18px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:generating?"not-allowed":"pointer",justifyContent:"center",opacity:generating?.7:1}}>
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
  useEffect(()=>{
    const h=e=>{ if(e.key==="Escape") onClose(); if(e.key==="ArrowLeft") setIdx(i=>(i+1)%images.length); if(e.key==="ArrowRight") setIdx(i=>(i-1+images.length)%images.length); };
    window.addEventListener("keydown",h); return ()=>window.removeEventListener("keydown",h);
  },[images,onClose]);
  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"#000e",zIndex:3000,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
      <div style={{position:"relative"}} onClick={e=>e.stopPropagation()}>
        <img src={images[idx]} alt="" style={{maxWidth:"88vw",maxHeight:"75vh",borderRadius:14,objectFit:"contain",display:"block"}}/>
        <button onClick={onClose} style={{position:"absolute",top:-12,left:-12,width:32,height:32,background:"#ef4444",border:"none",borderRadius:"50%",color:"#fff",fontSize:17,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
        {images.length>1&&<><button onClick={()=>setIdx(i=>(i-1+images.length)%images.length)} style={{position:"absolute",top:"50%",right:-50,transform:"translateY(-50%)",background:"#1a4faa",border:"1px solid #2563c7",color:"#fff",width:40,height:40,borderRadius:"50%",fontSize:22,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>‹</button><button onClick={()=>setIdx(i=>(i+1)%images.length)} style={{position:"absolute",top:"50%",left:-50,transform:"translateY(-50%)",background:"#1a4faa",border:"1px solid #2563c7",color:"#fff",width:40,height:40,borderRadius:"50%",fontSize:22,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>›</button></>}
        <div style={{textAlign:"center",marginTop:8,color:"#93c5fd",fontSize:12}}>{idx+1} / {images.length}</div>
      </div>
      {images.length>1&&(<div style={{display:"flex",gap:6,marginTop:10,overflowX:"auto",maxWidth:"88vw"}}>{images.map((img,i)=>(<img key={i} src={img} onClick={e=>{e.stopPropagation();setIdx(i);}} alt="" style={{width:54,height:54,objectFit:"cover",borderRadius:8,cursor:"pointer",flexShrink:0,border:i===idx?"2px solid #60a5fa":"2px solid transparent",opacity:i===idx?1:.5,transition:"all .2s"}}/>))}</div>)}
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
          <div style={{color:"#6b8cc4",fontSize:13}}>اسحب الصور أو <span style={{color:"#60a5fa",fontWeight:700}}>اضغط للاختيار</span><br/><span style={{fontSize:11,color:"#4a6fa5"}}>يتم ضغط الصور تلقائياً ⚡</span></div>
        )}
        <input ref={ref} type="file" accept="image/*" multiple style={{display:"none"}} onChange={e=>process(e.target.files)}/>
      </div>
      {images.length>0&&(<div style={{display:"flex",flexWrap:"wrap",gap:8}}>{images.map((img,i)=>(<div key={i} style={{position:"relative",width:72,height:72}}><img src={img} alt="" style={{width:"100%",height:"100%",objectFit:"cover",borderRadius:9,border:"2px solid #1e3a7a"}}/><button onClick={()=>onChange(images.filter((_,j)=>j!==i))} style={{position:"absolute",top:-5,right:-5,width:18,height:18,background:"#ef4444",border:"none",borderRadius:"50%",color:"#fff",fontSize:11,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>{i===0&&<div style={{position:"absolute",bottom:0,left:0,right:0,background:"#1a4faa",fontSize:9,color:"#fff",textAlign:"center",borderRadius:"0 0 7px 7px",fontWeight:900}}>رئيسية</div>}</div>))}</div>)}
    </div>
  );
}

function CodeInput({ label, value, onChange }) {
  const [show, setShow] = useState(false);
  return (<div><div style={{fontSize:11,color:"#6b8cc4",marginBottom:5,fontWeight:600}}>{label}</div><div style={{position:"relative"}}><input type={show?"text":"password"} value={value} onChange={onChange} placeholder="—" style={{width:"100%",boxSizing:"border-box",background:"#071840",border:"1px solid #fbbf2444",borderRadius:10,padding:"9px 36px 9px 12px",color:"#fbbf24",fontFamily:"inherit",fontSize:13,fontWeight:700}}/><button onClick={()=>setShow(s=>!s)} style={{position:"absolute",top:"50%",left:9,transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:14,padding:0}}>{show?"🙈":"👁️"}</button></div></div>);
}

// ── NAVBAR ────────────────────────────────────────────────────────────────────
function Navbar({ page, setPage, isAdmin, onLoginClick, onLogout, lang, setLang, scrolled, darkMode, setDarkMode, T }) {
  const isEn = lang==="en";
  const [menuOpen, setMenuOpen] = useState(false);
  const navItems = isEn
    ? [["home","🏠 Home"],["properties","🏘️ Properties"],["services","✦ Services"],["about","ℹ️ About"]]
    : [["home","🏠 الرئيسية"],["properties","🏘️ العقارات"],["services","✦ خدماتنا"],["about","ℹ️ عن المؤسسة"]];

  const go = (p) => { setPage(p); setMenuOpen(false); };

  return (
    <div style={{position:"fixed",top:0,left:0,right:0,zIndex:200,background:T.navbar,backdropFilter:"blur(18px)",borderBottom:`1px solid ${T.navbarBorder}`,transition:"background .3s"}}>
      <div style={{maxWidth:1200,margin:"0 auto",padding:"0 16px",display:"flex",alignItems:"center",justifyContent:"space-between",height:64}}>

        {/* Brand */}
        <button onClick={()=>go("home")} style={{display:"flex",alignItems:"center",gap:10,background:"none",border:"none",cursor:"pointer",padding:0,flexShrink:0}}>
          <div style={{width:44,height:44,borderRadius:10,overflow:"hidden",flexShrink:0,boxShadow:"0 2px 10px #1a4faa55"}}><img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1777727982/WhatsApp_Image_2026-04-30_at_1.38.52_AM_bx1fuy.jpg" alt="Logo" style={{width:"100%",height:"100%",objectFit:"cover"}}/></div>
          <div style={{textAlign:"right"}}>
            <div style={{fontWeight:900,fontSize:11,color:T.text,fontFamily:"'Cairo',sans-serif",lineHeight:1.2,whiteSpace:"nowrap"}}>{isEn?"Khalid Al-Shaikh Est.":"مؤسسة خالد محمد عبدالغفور الشيخ"}</div>
            <div style={{fontSize:9,color:T.text3,fontFamily:"'Cairo',sans-serif"}}>{isEn?"Real Estate Services":"للخدمات العقارية"}</div>
          </div>
        </button>

        {/* Desktop nav - hidden on mobile via CSS class */}
        <div className="desktop-nav" style={{display:"flex",gap:2}}>
          {navItems.map(([p,label])=>(<button key={p} onClick={()=>go(p)} style={{background:page===p?darkMode?"rgba(255,255,255,.1)":"rgba(26,79,170,.1)":"none",border:"none",color:page===p?T.text:T.text3,borderRadius:8,padding:"6px 10px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,cursor:"pointer",whiteSpace:"nowrap"}}>{label}</button>))}
        </div>

        {/* Right actions */}
        <div style={{display:"flex",alignItems:"center",gap:6,flexShrink:0}}>
          <div className="desktop-actions" style={{display:"flex",alignItems:"center",gap:6}}>
            <button onClick={()=>setDarkMode(d=>!d)} style={{background:darkMode?"rgba(255,255,255,.07)":"rgba(0,0,0,.06)",border:`1px solid ${darkMode?"rgba(255,255,255,.13)":"rgba(0,0,0,.12)"}`,color:T.text,borderRadius:8,padding:"5px 9px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>{darkMode?"☀️":"🌙"}</button>
            <button onClick={()=>setLang(isEn?"ar":"en")} style={{background:darkMode?"rgba(255,255,255,.07)":"rgba(0,0,0,.06)",border:`1px solid ${darkMode?"rgba(255,255,255,.13)":"rgba(0,0,0,.12)"}`,color:T.text,borderRadius:8,padding:"5px 9px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,cursor:"pointer"}}>{isEn?"🇸🇦":"🇬🇧 EN"}</button>
            <a href={`tel:${PHONE}`} style={{background:"#1a4faa",color:"#fff",borderRadius:8,padding:"6px 10px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,cursor:"pointer",textDecoration:"none",whiteSpace:"nowrap"}}>📞 {PHONE}</a>
            <a href={`https://wa.me/${WA_NUMBER}`} target="_blank" rel="noopener noreferrer" style={{background:"#25d36622",border:"1px solid #25d36640",color:"#25d366",borderRadius:8,padding:"6px 9px",textDecoration:"none",display:"flex",alignItems:"center"}}><WaIcon size={12}/></a>
            {isAdmin
              ? <button onClick={onLogout} style={{background:"#ef444422",border:"1px solid #ef444444",color:"#f87171",borderRadius:8,padding:"6px 10px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,cursor:"pointer"}}>🔓</button>
              : <button onClick={onLoginClick} style={{background:"linear-gradient(135deg,#1a4faa,#2563c7)",border:"none",color:"#fff",borderRadius:8,padding:"7px 12px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,cursor:"pointer",whiteSpace:"nowrap"}}>🔐 {isEn?"Admin":"الإدارة"}</button>
            }
          </div>
          {/* Hamburger - always visible */}
          <button onClick={()=>setMenuOpen(s=>!s)} style={{background:darkMode?"rgba(255,255,255,.08)":"rgba(0,0,0,.06)",border:`1px solid ${darkMode?"rgba(255,255,255,.15)":"rgba(0,0,0,.12)"}`,color:T.text,borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>
            {menuOpen?"✕":"☰"}
          </button>
        </div>
      </div>

      {/* Mobile dropdown menu */}
      {menuOpen && (
        <div style={{background:"rgba(7,16,58,.98)",borderTop:"1px solid rgba(255,255,255,.08)",padding:"12px 16px",display:"flex",flexDirection:"column",gap:8}}>
          {navItems.map(([p,label])=>(
            <button key={p} onClick={()=>go(p)} style={{background:page===p?"rgba(255,255,255,.1)":"transparent",border:"none",color:page===p?"#fff":"rgba(255,255,255,.6)",borderRadius:10,padding:"13px 16px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",textAlign:"right",width:"100%"}}>
              {label}
            </button>
          ))}
          <div style={{height:1,background:"rgba(255,255,255,.08)",margin:"4px 0"}}/>
          <button onClick={()=>{setLang(isEn?"ar":"en");setMenuOpen(false);}} style={{background:"rgba(255,255,255,.07)",border:"1px solid rgba(255,255,255,.13)",color:"#fff",borderRadius:10,padding:"13px 16px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",textAlign:"right",width:"100%"}}>{isEn?"🇸🇦 عربي":"🇬🇧 English"}</button>
          <button onClick={()=>{setDarkMode(d=>!d);setMenuOpen(false);}} style={{background:"rgba(255,255,255,.07)",border:"1px solid rgba(255,255,255,.13)",color:"#fff",borderRadius:10,padding:"13px 16px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",textAlign:"right",width:"100%"}}>{darkMode?"☀️ الوضع الفاتح":"🌙 الوضع الداكن"}</button>
          <a href={`tel:${PHONE}`} style={{display:"block",background:"#1a4faa",color:"#fff",borderRadius:10,padding:"13px 16px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,textAlign:"right"}}>📞 {PHONE}</a>
          <a href={`https://wa.me/${WA_NUMBER}`} target="_blank" rel="noopener noreferrer" style={{display:"flex",alignItems:"center",gap:8,background:"#25d36620",border:"1px solid #25d36640",color:"#25d366",borderRadius:10,padding:"13px 16px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,justifyContent:"flex-end"}}><WaIcon size={16}/> WhatsApp</a>
          <div style={{height:1,background:"rgba(255,255,255,.08)",margin:"4px 0"}}/>
          {isAdmin
            ? <button onClick={()=>{onLogout();setMenuOpen(false);}} style={{background:"#ef444420",border:"1px solid #ef444440",color:"#f87171",borderRadius:10,padding:"13px 16px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",textAlign:"right",width:"100%"}}>{isEn?"🔓 Logout":"🔓 تسجيل خروج"}</button>
            : <button onClick={()=>{onLoginClick();setMenuOpen(false);}} style={{background:"linear-gradient(135deg,#1a4faa,#2563c7)",border:"none",color:"#fff",borderRadius:10,padding:"13px 16px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",textAlign:"right",width:"100%"}}>🔐 {isEn?"Admin Login":"دخول الإدارة"}</button>
          }
        </div>
      )}

      {isAdmin&&(<div style={{background:"#1a4faa18",borderTop:"1px solid #2563c733",padding:"5px 20px",display:"flex",alignItems:"center",justifyContent:"center",gap:10}}><span style={{fontSize:11,color:"#93c5fd",fontWeight:600}}>🔒 {isEn?"Admin Mode":"وضع الإدارة"}</span></div>)}
    </div>
  );
}

// ── Login Modal ───────────────────────────────────────────────────────────────
function LoginModal({ onSuccess, onClose, lang }) {
  const isEn=lang==="en";
  const [un,setUn]=useState(""); const [pw,setPw]=useState(""); const [show,setShow]=useState(false); const [err,setErr]=useState(""); const [loading,setLoading]=useState(false);
  const login=()=>{ setLoading(true); setTimeout(()=>{ const ok=USERS.find(u=>u.username===un&&u.password===pw); if(ok) onSuccess(); else { setErr(isEn?"Incorrect credentials":"بيانات الدخول غير صحيحة"); setLoading(false); } },600); };
  const IST={width:"100%",boxSizing:"border-box",background:"#071840",border:"1px solid #1e3a7a",borderRadius:10,padding:"10px 13px",color:"#e8eef8",fontFamily:"'Cairo',sans-serif",fontSize:14};
  return (
    <div style={{position:"fixed",inset:0,background:"#000c",zIndex:1500,display:"flex",alignItems:"center",justifyContent:"center",padding:20}} onClick={onClose}>
      <div style={{background:"linear-gradient(160deg,#071840,#0a1f54)",border:"2px solid #2563c7",borderRadius:22,padding:32,maxWidth:380,width:"100%",boxShadow:"0 24px 60px #1a4faa33"}} onClick={e=>e.stopPropagation()}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{width:60,height:60,borderRadius:16,background:"linear-gradient(135deg,#1a4faa,#2563c7)",margin:"0 auto 14px",display:"flex",alignItems:"center",justifyContent:"center",fontSize:28}}>🏛️</div>
          <div style={{fontWeight:900,fontSize:18,color:"#fff",marginBottom:4}}>{isEn?"Admin Login":"دخول الإدارة"}</div>
          <div style={{fontSize:11,color:"#4a6fa5"}}>Khalid M. A. Ghafour Al-Shaikh Est.</div>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:14,marginBottom:16}}>
          <div><div style={{fontSize:12,color:"#6b8cc4",marginBottom:6,fontWeight:600}}>{isEn?"Username":"اسم المستخدم"}</div><input value={un} onChange={e=>setUn(e.target.value)} onKeyDown={e=>e.key==="Enter"&&login()} style={IST}/></div>
          <div><div style={{fontSize:12,color:"#6b8cc4",marginBottom:6,fontWeight:600}}>{isEn?"Password":"كلمة السر"}</div><div style={{position:"relative"}}><input type={show?"text":"password"} value={pw} onChange={e=>setPw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&login()} style={{...IST,paddingLeft:42}}/><button onClick={()=>setShow(s=>!s)} style={{position:"absolute",top:"50%",left:12,transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:16,padding:0}}>{show?"🙈":"👁️"}</button></div></div>
        </div>
        {err&&<div style={{background:"#ef444418",border:"1px solid #ef444430",color:"#f87171",borderRadius:10,padding:"9px 14px",fontSize:13,fontWeight:600,marginBottom:14,textAlign:"center"}}>⚠️ {err}</div>}
        <button onClick={login} disabled={loading} style={{width:"100%",background:"linear-gradient(135deg,#1a4faa,#2563c7)",color:"#fff",border:"none",borderRadius:12,padding:"13px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:15,cursor:"pointer",opacity:loading?.7:1}}>{loading?(isEn?"Verifying...":"⏳ جاري التحقق..."):(isEn?"Login 🔓":"دخول 🔓")}</button>
      </div>
    </div>
  );
}

// ── Prop Form ─────────────────────────────────────────────────────────────────
function PropForm({ form, setForm, onSave, onClose, editId }) {
  const f=key=>e=>setForm(p=>({...p,[key]:e.target.value}));
  const IST={width:"100%",boxSizing:"border-box",background:"#071840",border:"1px solid #1e3a7a",borderRadius:10,padding:"9px 12px",color:"#e8eef8",fontFamily:"'Cairo',sans-serif",fontSize:13};
  const Lbl=({c})=><div style={{fontSize:11,color:"#6b8cc4",marginBottom:5,fontWeight:600}}>{c}</div>;
  const Sec=({c})=><div style={{fontSize:11,color:"#60a5fa",fontWeight:700,marginBottom:10,padding:"6px 12px",background:"#1a4faa18",borderRadius:9,border:"1px solid #1a4faa33"}}>{c}</div>;
  return (
    <div style={{position:"fixed",inset:0,background:"#000b",zIndex:600,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={onClose}>
      <div style={{background:"linear-gradient(160deg,#071840,#0a1f54)",border:"1px solid #1a4faa",borderRadius:22,padding:24,maxWidth:600,width:"100%",maxHeight:"93vh",overflowY:"auto"}} onClick={e=>e.stopPropagation()}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20,paddingBottom:14,borderBottom:"1px solid #1e3a7a"}}>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <div style={{width:36,height:36,borderRadius:10,background:"linear-gradient(135deg,#1a4faa,#2563c7)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>🏢</div>
            <div><div style={{fontWeight:900,fontSize:14,color:"#e8eef8"}}>{editId?"تعديل العقار":"إضافة عقار جديد"}</div><div style={{fontSize:10,color:"#4a6fa5"}}>مؤسسة خالد محمد عبدالغفور الشيخ</div></div>
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
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:11,marginBottom:16}}>
          {[["اسم المالك","ownerName"],["رقم الجوال","ownerPhone"]].map(([lbl,key])=>(<div key={key}><Lbl c={lbl}/><input value={form[key]} onChange={f(key)} style={IST}/></div>))}
        </div>

        <Sec c="🔐 رموز الدخول"/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:11,marginBottom:16}}>
          <CodeInput label="رمز دخول الشقة" value={form.aptCode} onChange={f("aptCode")}/>
          <CodeInput label="رمز دخول المبنى" value={form.buildingCode} onChange={f("buildingCode")}/>
        </div>

        <Sec c="📍 موقع العقار"/>
        <div style={{marginBottom:14}}>
          <Lbl c="رابط قوقل ماب"/>
          <input value={form.mapUrl} onChange={f("mapUrl")} placeholder="https://maps.google.com/?q=..." style={{...IST,borderColor:"#16a34a33",color:"#4ade80"}}/>
          <div style={{fontSize:10,color:"#4a6fa5",marginTop:5}}>💡 قوقل ماب → ابحث عن الموقع → مشاركة → نسخ الرابط</div>
        </div>

        <div style={{marginBottom:14}}><Lbl c="📝 ملاحظات"/><textarea value={form.notes} onChange={f("notes")} rows={2} style={{...IST,resize:"none"}}/></div>

        <Sec c="🖼️ صور العقار"/>
        <div style={{marginBottom:20}}><ImageUploader images={form.images||[]} onChange={imgs=>setForm(p=>({...p,images:typeof imgs==="function"?imgs(p.images||[]):imgs}))}/></div>

        <div style={{display:"flex",gap:10}}>
          <button onClick={onSave} style={{flex:1,background:"linear-gradient(135deg,#1a4faa,#2563c7)",color:"#fff",border:"none",borderRadius:11,padding:"12px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer"}}>{editId?"حفظ التعديلات":"إضافة العقار"}</button>
          <button onClick={onClose} style={{background:"#071840",border:"1px solid #1e3a7a",color:"#6b8cc4",borderRadius:11,padding:"12px 18px",fontFamily:"'Cairo',sans-serif",cursor:"pointer",fontWeight:600}}>إلغاء</button>
        </div>
      </div>
    </div>
  );
}

// ── Public Card ───────────────────────────────────────────────────────────────
function PublicCard({ p, setLightbox, onShare, lang }) {
  const [hov,setHov]=useState(false);
  const isEn=lang==="en";
  const imgs=p.images||[]; const sc=SC[p.status]||SC["متوفر"];

  const trackView = async () => {
    try { await updateDoc(doc(db,"properties",p.id),{views:(p.views||0)+1}); } catch(e) {}
  };
  const waMsg=encodeURIComponent(`${isEn?"Hello, I'd like to inquire about":"مرحباً، أود الاستفسار عن"} ${p.name} - ${p.address}${p.rentPrice?" | "+Number(p.rentPrice).toLocaleString()+" ﷼/سنة":""}${p.salePrice?" | "+Number(p.salePrice).toLocaleString()+" ﷼":""}`);
  const statusEn={"متوفر":"Available","مؤجر":"Rented","مباع":"Sold","قريب الانتهاء":"Expiring","صيانة":"Maintenance"};
  return (
    <div style={{background:"linear-gradient(160deg,#071840,#0a1f54)",border:`1px solid ${sc.c}25`,borderRadius:18,overflow:"hidden",transition:"all .15s",transform:hov?"translateY(-3px)":"none",boxShadow:hov?`0 12px 36px ${sc.c}18`:"none"}}
      onMouseEnter={()=>setHov(true)} onMouseLeave={()=>setHov(false)}>
      <div style={{position:"relative",height:185,background:"#03102e",cursor:imgs.length>0?"pointer":"default"}} onClick={()=>{if(imgs.length>0){setLightbox({images:imgs,idx:0});trackView();}}}>
        {imgs.length>0?(<><img src={imgs[0]} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/><div style={{position:"absolute",inset:0,background:"linear-gradient(to top,#07184088,transparent 55%)",pointerEvents:"none"}}/>{imgs.length>1&&<div style={{position:"absolute",bottom:10,left:10,background:"#000a",color:"#fff",fontSize:11,fontWeight:700,padding:"3px 10px",borderRadius:18}}>📷 {imgs.length}</div>}</>):(
          <div style={{width:"100%",height:"100%",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",color:"#1e3a7a",gap:8}}><span style={{fontSize:38}}>🏠</span></div>
        )}
        <div style={{position:"absolute",top:10,right:10}}><span style={{display:"inline-flex",alignItems:"center",gap:5,padding:"4px 11px",borderRadius:20,background:sc.bg,color:sc.c,fontSize:11,fontWeight:700,border:`1px solid ${sc.c}33`}}><span style={{width:6,height:6,borderRadius:"50%",background:sc.c}}/>{isEn?statusEn[p.status]:p.status}</span></div>
        {p.furnished&&<div style={{position:"absolute",top:10,left:10,background:"#fbbf2420",color:"#fbbf24",fontSize:10,fontWeight:700,padding:"3px 9px",borderRadius:18,border:"1px solid #fbbf2440"}}>🛋️ {isEn?"Furnished":"مفروش"}</div>}
        <button onClick={e=>{e.stopPropagation();onShare(p);}} style={{position:"absolute",bottom:10,left:10,background:"#1a4faa",border:"none",color:"#fff",borderRadius:20,padding:"4px 10px",fontSize:11,fontWeight:700,cursor:"pointer"}}>📤 {isEn?"Share":"مشاركة"}</button>
      </div>
      <div style={{padding:"14px 16px"}}>
        <div style={{fontWeight:900,fontSize:15,color:"#e8eef8",marginBottom:3}}>{p.name}</div>
        <div style={{fontSize:11,color:"#4a6fa5",marginBottom:8}}>📍 {p.address}</div>
        <div style={{marginBottom:9}}>{p.adLicenseNo?<span style={{display:"inline-flex",alignItems:"center",gap:5,background:"#1a4faa22",border:"1px solid #2563c744",color:"#93c5fd",borderRadius:8,padding:"3px 10px",fontSize:10,fontWeight:700}}>🏛️ {isEn?"Ad License:":"رخصة إعلانية:"} {p.adLicenseNo}</span>:<span style={{display:"inline-flex",alignItems:"center",gap:5,background:"#ef444418",border:"1px solid #ef444430",color:"#f87171",borderRadius:8,padding:"3px 10px",fontSize:10,fontWeight:700}}>⚠️ {isEn?"Pending License":"قيد الترخيص"}</span>}</div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:9}}>
          {p.rentPrice&&<div style={{background:"#4ade8015",border:"1px solid #4ade8028",borderRadius:9,padding:"5px 10px",fontSize:11,color:"#4ade80",fontWeight:700}}>🏠 {Number(p.rentPrice).toLocaleString()} {isEn?"SAR/yr":"﷼/سنة"}</div>}
          {p.salePrice&&<div style={{background:"#fbbf2415",border:"1px solid #fbbf2428",borderRadius:9,padding:"5px 10px",fontSize:11,color:"#fbbf24",fontWeight:700}}>💰 {Number(p.salePrice).toLocaleString()} {isEn?"SAR":"﷼"}</div>}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:10}}>
          {p.area&&<div style={{background:"#071840",borderRadius:8,padding:"5px 9px",fontSize:11,color:"#6b8cc4"}}>📐 {p.area} م²</div>}
          {p.builtArea&&<div style={{background:"#071840",borderRadius:8,padding:"5px 9px",fontSize:11,color:"#6b8cc4"}}>🏗️ {p.builtArea} م²</div>}
        </div>
        {p.notes&&<div style={{fontSize:11,color:"#4a6fa5",background:"#071840",borderRadius:8,padding:"6px 10px",marginBottom:10}}>💬 {p.notes}</div>}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:7,marginBottom:p.mapUrl?8:0}}>
          <a href={`tel:${PHONE}`} style={{display:"flex",alignItems:"center",justifyContent:"center",gap:4,background:"linear-gradient(135deg,#1a4faa,#2563c7)",color:"#fff",borderRadius:10,padding:"9px 6px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11}}>📞 {isEn?"Call":"اتصال"}</a>
          <a href={`https://wa.me/${WA_NUMBER}?text=${waMsg}`} target="_blank" rel="noopener noreferrer" style={{display:"flex",alignItems:"center",justifyContent:"center",gap:4,background:"#25d36618",border:"1px solid #25d36640",color:"#25d366",borderRadius:10,padding:"9px 6px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11}}><WaIcon size={12}/> {isEn?"WhatsApp":"واتساب"}</a>
          <button onClick={()=>onShare(p)} style={{display:"flex",alignItems:"center",justifyContent:"center",gap:4,background:"#6366f118",border:"1px solid #6366f140",color:"#a5b4fc",borderRadius:10,padding:"9px 6px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,cursor:"pointer"}}>📤 {isEn?"Share":"شارك"}</button>
        </div>
        {p.mapUrl&&<a href={p.mapUrl} target="_blank" rel="noopener noreferrer" style={{display:"flex",alignItems:"center",justifyContent:"center",gap:6,background:"#16a34a18",border:"1px solid #16a34a40",color:"#4ade80",borderRadius:10,padding:"8px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:12,width:"100%",boxSizing:"border-box"}}>📍 {isEn?"View on Map":"عرض على الخريطة"}</a>}
      </div>
    </div>
  );
}

// ── Admin Card ────────────────────────────────────────────────────────────────
function AdminCard({ p, onEdit, onDelete, onChangeStatus, setLightbox, onShare }) {
  const [hov,setHov]=useState(false);
  const imgs=p.images||[]; const sc=SC[p.status]||SC["متوفر"];
  return (
    <div style={{background:"linear-gradient(160deg,#071840,#0a1f54)",border:`1px solid ${!p.adLicenseNo?"#ef444435":sc.c+"25"}`,borderRadius:18,overflow:"hidden",transition:"all .15s",transform:hov?"translateY(-3px)":"none"}}
      onMouseEnter={()=>setHov(true)} onMouseLeave={()=>setHov(false)}>
      {!p.adLicenseNo&&<div style={{background:"#ef444415",borderBottom:"1px solid #ef444428",padding:"5px 12px",fontSize:10,color:"#f87171",fontWeight:700}}>⚠️ يجب إضافة رقم الترخيص الإعلاني</div>}
      <div style={{position:"relative",height:148,background:"#03102e",cursor:imgs.length>0?"pointer":"default"}} onClick={()=>imgs.length>0&&setLightbox({images:imgs,idx:0})}>
        {imgs.length>0?(<><img src={imgs[0]} alt="" style={{width:"100%",height:"100%",objectFit:"cover"}}/><div style={{position:"absolute",inset:0,background:"linear-gradient(to top,#07184088,transparent 55%)",pointerEvents:"none"}}/></>):(<div style={{width:"100%",height:"100%",display:"flex",alignItems:"center",justifyContent:"center",color:"#1e3a7a",fontSize:32}}>🏠</div>)}
        <div style={{position:"absolute",top:8,right:8}}><span style={{display:"inline-flex",alignItems:"center",gap:4,padding:"3px 10px",borderRadius:20,background:sc.bg,color:sc.c,fontSize:10,fontWeight:700,border:`1px solid ${sc.c}33`}}><span style={{width:5,height:5,borderRadius:"50%",background:sc.c}}/>{p.status}</span></div>
        {p.refNo&&<div style={{position:"absolute",bottom:8,left:8,background:"#1a4faa",color:"#fff",fontSize:9,padding:"2px 7px",borderRadius:12,fontWeight:700}}>{p.refNo}</div>}
      </div>
      <div style={{padding:"11px 13px"}}>
        <div style={{fontWeight:900,fontSize:13,color:"#e8eef8",marginBottom:2}}>{p.name}</div>
        <div style={{fontSize:10,color:"#4a6fa5",marginBottom:7}}>📍 {p.address}</div>
        <div style={{background:"#03102e",borderRadius:9,padding:"8px 10px",marginBottom:7,border:"1px solid #0e2050"}}>
          <div style={{fontSize:9,color:p.adLicenseNo?"#4ade80":"#f87171",marginBottom:2}}>🏛️ رخصة: {p.adLicenseNo||"غير مُدخل ⚠️"}</div>
          <div style={{fontSize:9,color:"#6b8cc4",marginBottom:2}}>📋 عقد: {p.marketingContractNo||"—"}</div>
          {p.ownerName&&<div style={{fontSize:9,color:"#93c5fd",marginBottom:2}}>👤 {p.ownerName} {p.ownerPhone&&"— "+p.ownerPhone}</div>}
          <div style={{fontSize:9,color:"#a5b4fc",fontWeight:700}}>👁️ {p.views||0} مشاهدة</div>
        </div>
        <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:7}}>
          {p.rentPrice&&<div style={{background:"#4ade8015",border:"1px solid #4ade8025",borderRadius:7,padding:"3px 8px",fontSize:9,color:"#4ade80",fontWeight:700}}>إيجار: {Number(p.rentPrice).toLocaleString()}</div>}
          {p.salePrice&&<div style={{background:"#fbbf2415",border:"1px solid #fbbf2425",borderRadius:7,padding:"3px 8px",fontSize:9,color:"#fbbf24",fontWeight:700}}>بيع: {Number(p.salePrice).toLocaleString()}</div>}
          {p.minPrice&&<div style={{background:"#f8717115",border:"1px solid #f8717125",borderRadius:7,padding:"3px 8px",fontSize:9,color:"#f87171",fontWeight:700}}>أدنى: {Number(p.minPrice).toLocaleString()}</div>}
        </div>
        <div style={{display:"flex",gap:4,flexWrap:"wrap",marginBottom:8}}>{STATUS_OPTIONS.map(s=>(<button key={s} onClick={()=>onChangeStatus(p.id,s)} style={{padding:"3px 8px",borderRadius:6,border:"none",cursor:"pointer",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:9,background:p.status===s?"linear-gradient(135deg,#1a4faa,#2563c7)":"#0e2050",color:p.status===s?"#fff":"#4a6fa5"}}>{s}</button>))}</div>
        <div style={{display:"flex",gap:6}}>
          <button onClick={()=>onEdit(p)} style={{flex:1,background:"#0e2563",border:"1px solid #2563c7",color:"#93c5fd",borderRadius:8,padding:"7px",fontFamily:"'Cairo',sans-serif",fontWeight:600,fontSize:11,cursor:"pointer"}}>✏️ تعديل</button>
          <button onClick={()=>onShare(p)} style={{background:"#6366f118",border:"1px solid #6366f130",color:"#a5b4fc",borderRadius:8,padding:"7px 10px",cursor:"pointer",fontSize:11}}>📤</button>
          {imgs.length>0&&<button onClick={()=>setLightbox({images:imgs,idx:0})} style={{background:"#1a4faa22",border:"1px solid #2563c740",color:"#93c5fd",borderRadius:8,padding:"7px 10px",cursor:"pointer",fontSize:11}}>🖼️</button>}
          {p.mapUrl&&<a href={p.mapUrl} target="_blank" rel="noopener noreferrer" style={{background:"#16a34a18",border:"1px solid #16a34a40",color:"#4ade80",borderRadius:8,padding:"7px 10px",fontSize:11,textDecoration:"none",display:"flex",alignItems:"center"}}>📍</a>}
          <button onClick={()=>onDelete(p.id)} style={{background:"#ef444414",border:"1px solid #ef444428",color:"#f87171",borderRadius:8,padding:"7px 10px",cursor:"pointer",fontSize:11}}>🗑️</button>
        </div>
      </div>
    </div>
  );
}

// ── Pages ─────────────────────────────────────────────────────────────────────
function HomePage({ setPage, lang, darkMode, T }) {
  const isEn=lang==="en";
  const services=isEn?[
    {icon:"🏢",t:"Buy & Rent",d:"Best real estate opportunities for sale and rent at competitive prices"},
    {icon:"🗝️",t:"Property Management",d:"Full management — rent collection, contracts, tenant relations"},
    {icon:"📊",t:"Valuation",d:"Professional valuation per Real Estate General Authority standards"},
    {icon:"💡",t:"Consulting",d:"Specialized consulting including market analysis and investment guidance"},
    {icon:"🔧",t:"Maintenance",d:"Comprehensive maintenance services — routine and emergency"},
    {icon:"🛋️",t:"Furnished Rental",d:"Fully furnished units ready to move in immediately"},
    {icon:"📋",t:"Registry Services",d:"Property subdivisions, deed updates, and official registry procedures"},
    {icon:"🏛️",t:"Ejar Platform",d:"Lease contract documentation and legal contract preparation"},
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
    <div style={{paddingTop:64,background:T.bg}}>
      <div style={{minHeight:"calc(100vh - 64px)",background:darkMode?"linear-gradient(160deg,#07103a 0%,#0e2563 55%,#1a4faa 100%)":"linear-gradient(160deg,#eff6ff 0%,#dbeafe 55%,#bfdbfe 100%)",display:"flex",alignItems:"center",justifyContent:"center",position:"relative",overflow:"hidden",padding:"60px 24px"}}>
        <div style={{position:"absolute",inset:0,backgroundImage:"linear-gradient(rgba(255,255,255,.015) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.015) 1px,transparent 1px)",backgroundSize:"48px 48px"}}/>
        <div style={{position:"relative",zIndex:2,textAlign:"center",maxWidth:780}}>
          <div style={{display:"inline-flex",alignItems:"center",gap:8,background:"rgba(255,255,255,.07)",border:"1px solid rgba(255,255,255,.13)",borderRadius:30,padding:"6px 18px",fontSize:12,color:"rgba(255,255,255,.65)",marginBottom:28}}>
            <span style={{width:7,height:7,borderRadius:"50%",background:"#4ade80",boxShadow:"0 0 8px #4ade80",display:"inline-block"}}/>
            {isEn?"Licensed by Real Estate General Authority":"مرخصون من الهيئة العامة للعقار"}
          </div>
          {/* Hero Logo */}
          <div style={{marginBottom:24,display:"flex",justifyContent:"center"}}>
            <img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1777727982/WhatsApp_Image_2026-04-30_at_1.38.52_AM_bx1fuy.jpg"
              alt="Logo" style={{width:160,height:160,objectFit:"cover",borderRadius:24,boxShadow:"0 8px 40px #00000055"}}/>
          </div>
          <h1 style={{fontSize:"clamp(28px,5vw,54px)",fontWeight:900,color:darkMode?"#fff":"#1e3a6e",lineHeight:1.15,marginBottom:10}}>
            {isEn?"Khalid M. A. Ghafour":"مؤسسة خالد محمد"}<br/>
            <span style={{background:"linear-gradient(135deg,#60a5fa,#93c5fd)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",backgroundClip:"text"}}>{isEn?"Al-Shaikh Est.":"عبدالغفور الشيخ"}</span>
          </h1>
          <p style={{fontSize:13,color:darkMode?"rgba(255,255,255,.45)":"rgba(30,58,110,.5)",marginBottom:14}}>Khalid M. A. Ghafour Al-Shaikh Est. | Real Estate Services</p>
          <p style={{fontSize:15,color:darkMode?"rgba(255,255,255,.62)":"rgba(30,58,110,.7)",lineHeight:1.8,maxWidth:580,margin:"0 auto 36px"}}>{isEn?"Your trusted partner for all real estate services":"شريكك الموثوق في جميع الخدمات العقارية"}</p>
          <div style={{display:"flex",gap:12,justifyContent:"center",flexWrap:"wrap"}}>
            <button onClick={()=>setPage("properties")} style={{background:"linear-gradient(135deg,#1a4faa,#2563c7)",color:"#fff",border:"none",borderRadius:13,padding:"13px 28px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:15,cursor:"pointer"}}>🏘️ {isEn?"Browse Properties":"تصفح العقارات"}</button>
            <a href={`tel:${PHONE}`} style={{background:"rgba(255,255,255,.08)",border:"1px solid rgba(255,255,255,.18)",color:"#fff",borderRadius:13,padding:"13px 24px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:15,display:"flex",alignItems:"center",gap:7}}>📞 {isEn?"Contact Us":"تواصل الآن"}</a>
          </div>
          <div style={{display:"flex",justifyContent:"center",gap:40,marginTop:52,paddingTop:36,borderTop:"1px solid rgba(255,255,255,.07)",flexWrap:"wrap"}}>
            {[[isEn?"8+":"٨+",isEn?"Services":"خدمة عقارية"],[isEn?"100%":"١٠٠٪",isEn?"Authority Compliant":"امتثال للهيئة"],[isEn?"24/7":"٢٤/٧",isEn?"Available":"تواصل مستمر"]].map(([n,l])=>(
              <div key={l} style={{textAlign:"center"}}><div style={{fontSize:28,fontWeight:900,color:darkMode?"#fff":"#1e3a6e"}}>{n}</div><div style={{fontSize:11,color:darkMode?"rgba(255,255,255,.4)":"rgba(30,58,110,.5)",marginTop:3}}>{l}</div></div>
            ))}
          </div>
        </div>
      </div>
      <div style={{background:T.bg,padding:"70px 24px"}}>
        <div style={{maxWidth:1100,margin:"0 auto"}}>
          <div style={{textAlign:"center",marginBottom:46}}>
            <div style={{display:"inline-block",background:"#1a4faa22",border:"1px solid #1a4faa44",color:"#93c5fd",borderRadius:20,padding:"5px 16px",fontSize:11,fontWeight:700,marginBottom:12}}>{isEn?"Our Services":"خدماتنا"}</div>
            <h2 style={{fontSize:"clamp(20px,3.5vw,32px)",fontWeight:900,color:T.text}}>{isEn?"Complete Real Estate Services":"خدمات عقارية متكاملة تحت سقف واحد"}</h2>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:14}}>
            {services.map((s,i)=>(<div key={i} style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:17,padding:"24px 20px",transition:"all .25s"}} onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-4px)";}} onMouseLeave={e=>{e.currentTarget.style.transform="translateY(0)";}}>
              <div style={{width:46,height:46,borderRadius:13,background:darkMode?"#1a4faa22":"#dbeafe",border:`1px solid ${T.border}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:21,marginBottom:14}}>{s.icon}</div>
              <div style={{fontSize:14,fontWeight:800,color:T.text,marginBottom:7}}>{s.t}</div>
              <div style={{fontSize:12,color:T.text3,lineHeight:1.7}}>{s.d}</div>
            </div>))}
          </div>
        </div>
      </div>
      <div style={{background:"linear-gradient(135deg,#0e2563,#1a4faa)",padding:"66px 24px",textAlign:"center"}}>
        <h2 style={{fontSize:"clamp(20px,3.5vw,34px)",fontWeight:900,color:"#fff",marginBottom:10}}>{isEn?"Looking for a property?":"هل تبحث عن عقار؟"}</h2>
        <p style={{color:"rgba(255,255,255,.6)",fontSize:14,marginBottom:28}}>{isEn?"Our team is ready 24/7":"فريقنا جاهز لمساعدتك على مدار الساعة"}</p>
        <div style={{display:"flex",gap:12,justifyContent:"center",flexWrap:"wrap"}}>
          <a href={`tel:${PHONE}`} style={{display:"inline-flex",alignItems:"center",gap:8,background:"rgba(255,255,255,.14)",border:"1px solid rgba(255,255,255,.28)",color:"#fff",borderRadius:13,padding:"12px 26px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:900,fontSize:16}}>📞 {PHONE}</a>
          <a href={`https://wa.me/${WA_NUMBER}`} target="_blank" rel="noopener noreferrer" style={{display:"inline-flex",alignItems:"center",gap:8,background:"#25d36620",border:"1px solid #25d36650",color:"#25d366",borderRadius:13,padding:"12px 24px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:15}}><WaIcon size={18}/> WhatsApp</a>
        </div>
      </div>
      <div style={{background:T.bg4,borderTop:`1px solid ${T.border}`,padding:"28px 24px"}}>
        <div style={{maxWidth:1100,margin:"0 auto",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:14}}>
          <div style={{display:"flex",alignItems:"center",gap:11}}>
            <div style={{width:36,height:36,borderRadius:10,overflow:"hidden"}}><img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1777727982/WhatsApp_Image_2026-04-30_at_1.38.52_AM_bx1fuy.jpg" alt="Logo" style={{width:"100%",height:"100%",objectFit:"cover"}}/></div>
            <div><div style={{fontWeight:800,fontSize:12,color:"#fff"}}>مؤسسة خالد محمد عبدالغفور الشيخ</div><div style={{fontSize:10,color:"#4a6fa5"}}>Khalid M. A. Ghafour Al-Shaikh Est.</div></div>
          </div>
          <div style={{fontSize:11,color:"#2a3a6a"}}>© 2025 | {isEn?"Licensed by Real Estate General Authority":"مرخصة من الهيئة العامة للعقار"}</div>
        </div>
      </div>
    </div>
  );
}

function PropertiesPage({ props, isAdmin, onEdit, onDelete, onChangeStatus, setLightbox, onShare, onOpenAdd, lang, darkMode, T }) {
  const isEn=lang==="en";

  // ترجمة الفلاتر
  const statusLabels = isEn
    ? {all:"All", "متوفر":"Available","مؤجر":"Rented","مباع":"Sold","قريب الانتهاء":"Expiring","صيانة":"Maintenance"}
    : {all:"الكل", "متوفر":"متوفر","مؤجر":"مؤجر","مباع":"مباع","قريب الانتهاء":"قريب الانتهاء","صيانة":"صيانة"};
  const dealLabels = isEn
    ? {all:"All","إيجار":"Rent","بيع":"Sale","إيجار وبيع":"Rent & Sale"}
    : {all:"الكل","إيجار":"إيجار","بيع":"بيع","إيجار وبيع":"إيجار وبيع"};

  const [filter,setFilter]=useState("all"); const [search,setSearch]=useState(""); const [dealF,setDealF]=useState("all");

  const filtered=props.filter(p=>
    (filter==="all"||p.status===filter)&&
    (dealF==="all"||p.dealType===dealF)&&
    (p.name.includes(search)||p.address.includes(search)||(p.ownerName||"").includes(search))
  );
  return (
    <div style={{paddingTop:isAdmin?90:64,minHeight:"100vh",background:T.bg}}>
      <div style={{background:"linear-gradient(135deg,#0e2563,#1a4faa)",padding:"28px 24px 24px"}}>
        <div style={{maxWidth:1200,margin:"0 auto",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
          <div><div style={{fontWeight:900,fontSize:22,color:"#fff",marginBottom:3}}>{isAdmin?(isEn?"🔒 Admin Panel":"🔒 لوحة الإدارة"):(isEn?"🏘️ Available Properties":"🏘️ العقارات المتاحة")}</div><div style={{fontSize:12,color:"rgba(255,255,255,.5)"}}>Khalid M. A. Ghafour Al-Shaikh Est.</div></div>
          {isAdmin&&<div style={{display:"flex",gap:8}}>
            <button onClick={()=>exportToExcel(props)} style={{background:"#16a34a22",border:"1px solid #16a34a44",color:"#4ade80",borderRadius:11,padding:"10px 18px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>📊 {isEn?"Export Excel":"تصدير Excel"}</button>
            <button onClick={onOpenAdd} style={{background:"rgba(255,255,255,.15)",border:"1px solid rgba(255,255,255,.3)",color:"#fff",borderRadius:11,padding:"10px 22px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>+ {isEn?"Add Property":"إضافة عقار"}</button>
          </div>}
        </div>
      </div>
      <div style={{maxWidth:1200,margin:"0 auto",padding:"22px"}}>
        {isAdmin&&(()=>{const st={total:props.length,available:props.filter(p=>p.status==="متوفر").length,rented:props.filter(p=>p.status==="مؤجر").length,income:props.filter(p=>p.rentPrice).reduce((s,p)=>s+Number(p.rentPrice),0),noLicense:props.filter(p=>!p.adLicenseNo).length};return(<div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:9,marginBottom:20}}>{[{l:isEn?"Total":"الإجمالي",v:st.total,i:"🏢",c:"#93c5fd"},{l:isEn?"Available":"متوفر",v:st.available,i:"✅",c:"#4ade80"},{l:isEn?"Rented":"مؤجر",v:st.rented,i:"🔑",c:"#a5b4fc"},{l:isEn?"Monthly Income":"الدخل السنوي",v:st.income.toLocaleString()+" ﷼",i:"💰",c:"#fbbf24"},{l:isEn?"No License":"بدون ترخيص",v:st.noLicense,i:"⚠️",c:"#f87171"}].map((s,i)=>(<div key={i} style={{background:"linear-gradient(135deg,#071840,#0e2563)",border:`1px solid ${s.c}22`,borderRadius:12,padding:"12px 10px",position:"relative",overflow:"hidden"}}><div style={{position:"absolute",top:-8,left:-8,width:32,height:32,background:s.c+"15",borderRadius:"50%"}}/><div style={{fontSize:17,marginBottom:5}}>{s.i}</div><div style={{fontWeight:900,fontSize:15,color:s.c}}>{s.v}</div><div style={{fontSize:10,color:"#4a6fa5",marginTop:1}}>{s.l}</div></div>))}</div>);})()}
        <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder={isEn?"🔍 Search...":"🔍 ابحث..."} style={{flex:1,minWidth:180,background:T.bg2,border:`1px solid ${T.border}`,borderRadius:10,padding:"8px 13px",color:T.text,fontFamily:"'Cairo',sans-serif",fontSize:13}}/>
          <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
            {["all",...STATUS_OPTIONS].map(f=>(<button key={f} onClick={()=>setFilter(f)} style={{padding:"7px 12px",borderRadius:8,cursor:"pointer",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,background:filter===f?"linear-gradient(135deg,#1a4faa,#2563c7)":T.bg2,color:filter===f?"#fff":T.text3,border:filter===f?"none":`1px solid ${T.border}`}}>{f==="all"?statusLabels.all:statusLabels[f]}</button>))}
          </div>
          <div style={{display:"flex",gap:5}}>
            {["all","إيجار","بيع","إيجار وبيع"].map(f=>(<button key={f} onClick={()=>setDealF(f)} style={{padding:"7px 12px",borderRadius:8,cursor:"pointer",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:11,background:dealF===f?"#fbbf2433":T.bg2,color:dealF===f?"#fbbf24":T.text3,border:dealF===f?"1px solid #fbbf2444":`1px solid ${T.border}`}}>{f==="all"?dealLabels.all:dealLabels[f]}</button>))}
          </div>
        </div>
        <div style={{fontSize:11,color:T.text3,marginBottom:12}}>{filtered.length} {isEn?"properties":"عقار"}</div>
        {filtered.length===0?(<div style={{textAlign:"center",padding:"60px 0",color:"#1e3a7a"}}><div style={{fontSize:46,marginBottom:10}}>🏚️</div><div style={{fontSize:13,fontWeight:600}}>{isEn?"No results":"لا توجد نتائج"}</div></div>):(
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(295px,1fr))",gap:15}}>
            {filtered.map(p=>isAdmin
              ? <AdminCard key={p.id} p={p} onEdit={onEdit} onDelete={onDelete} onChangeStatus={onChangeStatus} setLightbox={setLightbox} onShare={onShare}/>
              : <PublicCard key={p.id} p={p} setLightbox={setLightbox} onShare={onShare} lang={lang}/>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

function AboutPage({ lang, darkMode, T }) {
  const isEn=lang==="en";
  return (
    <div style={{paddingTop:64,minHeight:"100vh",background:T.bg}}>
      <div style={{background:"linear-gradient(135deg,#0e2563,#1a4faa)",padding:"28px 24px 24px"}}><div style={{maxWidth:900,margin:"0 auto"}}><div style={{fontWeight:900,fontSize:22,color:"#fff",marginBottom:3}}>ℹ️ {isEn?"About Us":"عن المؤسسة"}</div></div></div>
      <div style={{maxWidth:900,margin:"0 auto",padding:"36px 24px"}}>

        {/* Main card */}
        <div style={{background:"linear-gradient(140deg,#0e2050,#12286a)",border:"1px solid rgba(255,255,255,.07)",borderRadius:20,padding:"32px",marginBottom:16}}>
          <div style={{display:"flex",alignItems:"center",gap:18,marginBottom:24,flexWrap:"wrap"}}>
            <div style={{width:80,height:80,borderRadius:17,overflow:"hidden",flexShrink:0,boxShadow:"0 4px 20px #1a4faa55"}}><img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1777727982/WhatsApp_Image_2026-04-30_at_1.38.52_AM_bx1fuy.jpg" alt="Logo" style={{width:"100%",height:"100%",objectFit:"cover"}}/></div>
            <div>
              <div style={{fontWeight:900,fontSize:20,color:"#fff",marginBottom:4}}>مؤسسة خالد محمد عبدالغفور الشيخ</div>
              <div style={{fontSize:13,color:"#6b8cc4"}}>Khalid M. A. Ghafour Al-Shaikh Est. | للخدمات العقارية</div>
              <div style={{marginTop:8,display:"inline-flex",alignItems:"center",gap:6,background:"#4ade8015",border:"1px solid #4ade8030",borderRadius:20,padding:"3px 12px"}}>
                <span style={{width:7,height:7,borderRadius:"50%",background:"#4ade80",display:"inline-block"}}/>
                <span style={{fontSize:11,color:"#4ade80",fontWeight:700}}>مرخصة من الهيئة العامة للعقار</span>
              </div>
            </div>
          </div>

          {/* About text */}
          <div style={{fontSize:14,color:"#93c5fd",lineHeight:2,marginBottom:24}}>
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
              <div key={i} style={{background:"#03102e",borderRadius:12,padding:"14px",textAlign:"center",border:"1px solid #1e3a7a"}}>
                <div style={{fontSize:22,marginBottom:4}}>{s.i}</div>
                <div style={{fontWeight:900,fontSize:20,color:"#fff",marginBottom:2}}>{s.n}</div>
                <div style={{fontSize:11,color:"#4a6fa5"}}>{s.l}</div>
              </div>
            ))}
          </div>

          {/* Contact grid */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            {[["📞 "+PHONE,PHONE],["💬 WhatsApp","0568300022"],["🏛️ "+(isEn?"License":"الترخيص"),isEn?"Real Estate Gen. Authority":"هيئة العقار"],["📍 "+(isEn?"Location":"الموقع"),isEn?"Eastern Province, KSA":"المنطقة الشرقية، المملكة العربية السعودية"]].map(([l,v])=>(
              <div key={l} style={{background:"#03102e",borderRadius:11,padding:"13px 15px",border:"1px solid #1e3a7a"}}><div style={{fontSize:11,color:"#4a6fa5",marginBottom:4}}>{l}</div><div style={{fontSize:12,fontWeight:700,color:"#93c5fd"}}>{v}</div></div>
            ))}
          </div>
        </div>

        {/* CTA */}
        <div style={{background:"#1a4faa18",border:"1px solid #2563c740",borderRadius:15,padding:"22px",textAlign:"center"}}>
          <div style={{fontWeight:900,fontSize:15,color:"#fff",marginBottom:6}}>{isEn?"Get in Touch":"تواصل معنا"}</div>
          <div style={{fontSize:12,color:"#4a6fa5",marginBottom:16}}>{isEn?"We're here to help 24/7":"فريقنا جاهز لمساعدتك على مدار الساعة"}</div>
          <div style={{display:"flex",gap:10,justifyContent:"center",flexWrap:"wrap"}}>
            <a href={`tel:${PHONE}`} style={{display:"inline-flex",alignItems:"center",gap:6,background:"linear-gradient(135deg,#1a4faa,#2563c7)",color:"#fff",borderRadius:11,padding:"10px 20px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14}}>📞 {PHONE}</a>
            <a href={`https://wa.me/${WA_NUMBER}`} target="_blank" rel="noopener noreferrer" style={{display:"inline-flex",alignItems:"center",gap:7,background:"#25d36620",border:"1px solid #25d36644",color:"#25d366",borderRadius:11,padding:"10px 18px",textDecoration:"none",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14}}><WaIcon size={15}/> WhatsApp</a>
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
  const openEdit=p=>{ setForm({...p,images:p.images||[],mapUrl:p.mapUrl||""}); setEditId(p.id); setShowForm(true); };

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
    <div style={{direction:"rtl",fontFamily:"'Cairo',sans-serif",minHeight:"100vh",background:"#07103a",display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:14}}>
      <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet"/>
      <div style={{width:56,height:56,borderRadius:16,overflow:"hidden"}}><img src="https://res.cloudinary.com/dumtp0krl/image/upload/v1777727982/WhatsApp_Image_2026-04-30_at_1.38.52_AM_bx1fuy.jpg" alt="Logo" style={{width:"100%",height:"100%",objectFit:"cover"}}/></div>
      <div style={{color:"#93c5fd",fontWeight:700,fontSize:14,fontFamily:"'Cairo',sans-serif"}}>جاري التحميل...</div>
    </div>
  );

  const T = darkMode ? {
    bg:"#07103a", bg2:"#071840", bg3:"#0a1f54", bg4:"#03102e",
    text:"#e8eef8", text2:"#93c5fd", text3:"#4a6fa5",
    border:"#1e3a7a", card:"linear-gradient(160deg,#071840,#0a1f54)",
    navbar:"rgba(7,16,58,.98)", navbarBorder:"rgba(255,255,255,.08)",
    navText:"rgba(255,255,255,.6)", navActive:"rgba(255,255,255,.1)",
  } : {
    bg:"#f0f4ff", bg2:"#fff", bg3:"#e8eeff", bg4:"#f8faff",
    text:"#1e3a6e", text2:"#1a4faa", text3:"#64748b",
    border:"#dbeafe", card:"linear-gradient(160deg,#fff,#f0f4ff)",
    navbar:"rgba(255,255,255,.98)", navbarBorder:"rgba(0,0,0,.08)",
    navText:"rgba(30,58,110,.6)", navActive:"rgba(26,79,170,.1)",
  };

  return (
    <div style={{direction:isEn?"ltr":"rtl",fontFamily:"'Cairo',sans-serif",minHeight:"100vh",background:T.bg,color:T.text,transition:"background .3s,color .3s"}}>
      <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet"/>

      {lightbox&&<Lightbox images={lightbox.images} startIndex={lightbox.idx} onClose={()=>setLightbox(null)}/>}
      {showLogin&&<LoginModal onSuccess={()=>{setIsAdmin(true);setShowLogin(false);showToast(isEn?"Welcome! Admin mode active 🔓":"مرحباً! وضع الإدارة مفعّل 🔓");}} onClose={()=>setShowLogin(false)} lang={lang}/>}
      {showForm&&<PropForm form={form} setForm={setForm} onSave={save} onClose={()=>setShowForm(false)} editId={editId} T={T}/>}
      {shareP&&<ShareModal p={shareP} onClose={()=>setShareP(null)}/>}

      {toast&&(<div style={{position:"fixed",top:20,left:"50%",transform:"translateX(-50%)",background:toast.type==="err"?"#ef4444":"#1a4faa",border:`1px solid ${toast.type==="err"?"#ef4444":"#2563c7"}`,color:"#fff",padding:"10px 24px",borderRadius:12,zIndex:9999,fontWeight:700,fontSize:13,boxShadow:"0 8px 32px #000a",whiteSpace:"nowrap"}}>{toast.msg}</div>)}
      {saving&&(<div style={{position:"fixed",bottom:16,left:16,background:"#0e2563",border:"1px solid #2563c7",color:"#93c5fd",padding:"6px 13px",borderRadius:9,fontSize:11,fontWeight:600,zIndex:9998}}>💾 {isEn?"Saving...":"جاري الحفظ..."}</div>)}

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

      <Navbar page={page} setPage={setPage} isAdmin={isAdmin} onLoginClick={()=>setShowLogin(true)} onLogout={()=>{setIsAdmin(false);showToast(isEn?"Logged out":"تم تسجيل الخروج");}} lang={lang} setLang={setLang} scrolled={scrolled} darkMode={darkMode} setDarkMode={setDarkMode} T={T}/>

      {page==="home"       && <HomePage setPage={setPage} lang={lang} darkMode={darkMode} T={T}/>}
      {page==="properties" && <PropertiesPage props={props} isAdmin={isAdmin} onEdit={openEdit} onDelete={id=>setDelId(id)} onChangeStatus={changeStatus} setLightbox={setLightbox} onShare={setShareP} onOpenAdd={openAdd} lang={lang} darkMode={darkMode} T={T}/>}
      {page==="services"   && <HomePage setPage={setPage} lang={lang} darkMode={darkMode} T={T}/>}
      {page==="about"      && <AboutPage lang={lang} darkMode={darkMode} T={T}/>}

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
        input:focus,select:focus,textarea:focus{outline:none;border-color:#2563c7!important;}
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-track{background:${T.bg}}
        ::-webkit-scrollbar-thumb{background:${darkMode?"#1e3a7a":"#bfdbfe"};border-radius:4px}
        a{font-family:'Cairo',sans-serif;}
        @media(max-width:768px){
          .desktop-nav{display:none!important;}
          .desktop-actions{display:none!important;}
        }
      `}</style>
    </div>
  );
}
