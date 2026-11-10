import { useState, useEffect, useCallback, useRef, useMemo } from "react";
import { db } from "./firebase";
import { initializeApp, deleteApp } from "firebase/app";
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut, sendPasswordResetEmail, createUserWithEmailAndPassword, updatePassword, reauthenticateWithCredential, EmailAuthProvider } from "firebase/auth";
import { collection, getDocs, getDoc, setDoc, addDoc, updateDoc, deleteDoc, deleteField, doc, onSnapshot, runTransaction, query, where, orderBy, limit } from "firebase/firestore";

// Firebase Authentication — يستخدم نفس تطبيق Firebase الذي أنشأه ./firebase
const auth = getAuth(db.app);
// سجل التدقيق: يُكتب للمدير فقط ولا يُعدَّل ولا يُحذف (القواعد تمنع ذلك)
const logAudit = (action, target, detail="") => {
  try{
    const u = auth.currentUser; if(!u) return;
    addDoc(collection(db,"audit_logs"),{ ts:new Date().toISOString(), uid:u.uid, email:u.email||"", action, target:String(target||""), detail:String(detail||"").slice(0,300) }).catch(()=>{});
  }catch(e){}
};
const hesc = s => String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

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
              {userRole==="admin"&&<button onClick={()=>go("employees")} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 12px",borderRadius:12,border:"none",background:page==="employees"?(darkMode?"rgba(74,158,255,.15)":"rgba(74,158,255,.08)"):"none",color:page==="employees"?(darkMode?"#7ab8ff":"#1e3a7a"):(darkMode?"rgba(255,255,255,.6)":"#5a6a90"),fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",width:"100%",textAlign:"right",marginBottom:2}}>
                <div style={{width:36,height:36,borderRadius:10,background:darkMode?"rgba(255,255,255,.06)":"#edf1fb",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>👨‍💼</div>
                <span>{isEn?"Employees":"إدارة الموظفين"}</span>
              </button>}
              {userRole==="admin"&&<button onClick={()=>go("pm")} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 12px",borderRadius:12,border:"none",background:page==="pm"?(darkMode?"rgba(74,158,255,.15)":"rgba(74,158,255,.08)"):"none",color:page==="pm"?(darkMode?"#7ab8ff":"#1e3a7a"):(darkMode?"rgba(255,255,255,.6)":"#5a6a90"),fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",width:"100%",textAlign:"right",marginBottom:2}}>
                <div style={{width:36,height:36,borderRadius:10,background:darkMode?"rgba(255,255,255,.06)":"#edf1fb",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>🏛️</div>
                <span>{isEn?"Property Mgmt":"إدارة الأملاك"}</span>
              </button>}
              {userRole==="admin"&&<button onClick={()=>go("security")} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 12px",borderRadius:12,border:"none",background:page==="security"?(darkMode?"rgba(74,158,255,.15)":"rgba(74,158,255,.08)"):"none",color:page==="security"?(darkMode?"#7ab8ff":"#1e3a7a"):(darkMode?"rgba(255,255,255,.6)":"#5a6a90"),fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",width:"100%",textAlign:"right",marginBottom:2}}>
                <div style={{width:36,height:36,borderRadius:10,background:darkMode?"rgba(255,255,255,.06)":"#edf1fb",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>🛡️</div>
                <span>{isEn?"Security":"الأمان"}</span>
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
function LoginModal({ onAttempt, onClose, lang }) {
  const isEn=lang==="en";
  const [em,setEm]=useState(""); const [pw,setPw]=useState(""); const [show,setShow]=useState(false);
  const [err,setErr]=useState(""); const [info,setInfo]=useState(""); const [loading,setLoading]=useState(false);
  const login=async()=>{
    if(loading) return;
    setErr(""); setInfo("");
    if(!em.trim()||!pw) return setErr(isEn?"Enter email and password":"اكتب البريد وكلمة السر");
    setLoading(true);
    try{
      onAttempt();
      await signInWithEmailAndPassword(auth,em.trim(),pw);
      onClose();
    }catch(e){
      setErr(e?.code==="auth/too-many-requests"
        ? (isEn?"Too many attempts — try again later":"محاولات كثيرة — انتظر قليلاً ثم أعد المحاولة")
        : (isEn?"Incorrect credentials":"بيانات الدخول غير صحيحة"));
      setLoading(false);
    }
  };
  const forgot=async()=>{
    setErr(""); setInfo("");
    if(!em.trim()) return setErr(isEn?"Enter your email first":"اكتب بريدك أولاً ثم اضغط «نسيت كلمة السر»");
    try{ await sendPasswordResetEmail(auth,em.trim()); }catch(e){}
    setInfo(isEn?"If the email is registered, a reset link was sent.":"إذا كان البريد مسجلاً فسيصلك رابط لإعادة تعيين كلمة السر.");
  };
  const IST={width:"100%",boxSizing:"border-box",background:"#f0f4fc",border:"1px solid rgba(74,158,255,.2)",borderRadius:10,padding:"10px 13px",color:"#1e3a7a",fontFamily:"'Cairo',sans-serif",fontSize:14};
  return (
    <div style={{position:"fixed",inset:0,background:"#000c",zIndex:1500,display:"flex",alignItems:"center",justifyContent:"center",padding:20}} onClick={onClose}>
      <div style={{background:"linear-gradient(160deg,#ffffff,#f5f8ff)",border:"2px solid #2563c7",borderRadius:22,padding:32,maxWidth:380,width:"100%",boxShadow:"0 24px 60px #1a4faa33"}} onClick={e=>e.stopPropagation()}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{width:60,height:60,borderRadius:16,background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",margin:"0 auto 14px",display:"flex",alignItems:"center",justifyContent:"center",fontSize:28}}>🏛️</div>
          <div style={{fontWeight:900,fontSize:18,color:"#1e3a7a",marginBottom:4}}>{isEn?"Staff Login":"دخول الإدارة والموظفين"}</div>
          <div style={{fontSize:11,color:"#5a6a90"}}>Khalid M. A. Ghafour Al-Shaikh Est.</div>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:14,marginBottom:16}}>
          <div><div style={{fontSize:12,color:"#5a6a90",marginBottom:6,fontWeight:600}}>{isEn?"Email":"البريد الإلكتروني"}</div>
            <input type="email" autoComplete="username" dir="ltr" value={em} onChange={e=>setEm(e.target.value)} onKeyDown={e=>e.key==="Enter"&&login()} style={IST}/></div>
          <div><div style={{fontSize:12,color:"#5a6a90",marginBottom:6,fontWeight:600}}>{isEn?"Password":"كلمة السر"}</div>
            <div style={{position:"relative"}}>
              <input type={show?"text":"password"} autoComplete="current-password" dir="ltr" value={pw} onChange={e=>setPw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&login()} style={{...IST,paddingLeft:40}}/>
              <button type="button" onClick={()=>setShow(s=>!s)} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:16}}>{show?"🙈":"👁️"}</button>
            </div></div>
        </div>
        {err&&<div style={{background:"#ef444418",border:"1px solid #ef444430",color:"#ef4444",borderRadius:10,padding:"9px 14px",fontSize:13,fontWeight:600,marginBottom:14,textAlign:"center"}}>⚠️ {err}</div>}
        {info&&<div style={{background:"#16a34a14",border:"1px solid #16a34a30",color:"#16a34a",borderRadius:10,padding:"9px 14px",fontSize:12,fontWeight:600,marginBottom:14,textAlign:"center"}}>✔ {info}</div>}
        <button onClick={login} disabled={loading} style={{width:"100%",background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",color:"#fff",border:"none",borderRadius:12,padding:"13px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:15,cursor:"pointer",opacity:loading?.7:1}}>{loading?(isEn?"Verifying...":"⏳ جاري التحقق..."):(isEn?"Login":"🔐 دخول")}</button>
        <button onClick={forgot} style={{width:"100%",background:"none",border:"none",color:"#2a4d9b",fontFamily:"'Cairo',sans-serif",fontWeight:600,fontSize:12,cursor:"pointer",marginTop:12}}>{isEn?"Forgot password?":"نسيت كلمة السر؟"}</button>
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
  const [dbEmployees, setDbEmployees] = useState([]);
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
    // الموظف يقرأ فقط العملاء المحوّلين له — والقواعد في Firestore تفرض ذلك أيضاً
    const base = collection(db,"clients");
    const q = isManager ? base : query(base, where("assignedTo","==",currentUser?.uid||"__none__"));
    const unsub = onSnapshot(q, snap=>{
      const data = snap.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>b.createdAt?.localeCompare?.(a.createdAt)||0);
      setClients(data); setLoaded(true);
    }, ()=>setLoaded(true));
    return ()=>unsub();
  },[isManager,currentUser?.uid]);

  useEffect(()=>{
    if(!isManager) return;
    const unsub = onSnapshot(query(collection(db,"users"),where("role","==","employee")), snap=>{
      setDbEmployees(snap.docs.map(d=>({id:d.id,...d.data()})));
    }, ()=>{});
    return ()=>unsub();
  },[isManager]);

  const openAdd = () => { setForm(emptyClient); setEditClientId(null); setShowForm(true); };
  const openEdit = (c) => { 
    const fd = {...emptyClient,...c};
    if(isManager===false && c.clientType==="مالك") fd.phone="";
    setForm(fd); setEditClientId(c.id); setShowForm(true); 
  };

  const save = async () => {
    if(!form.name||!form.phone) return alert("الاسم ورقم الجوال مطلوبان");
    const pk = normPhone(form.phone);
    let duplicate = null;
    try{
      if(pk){ const ps = await getDoc(doc(db,"client_phones",pk)); if(ps.exists()&&ps.data().clientId!==editClientId) duplicate = ps.data(); }
    }catch(e){}
    if(duplicate) { const ok=window.confirm(`⚠️ رقم الجوال مكرر!\nمسجل للعميل: "${duplicate.name}" — #${duplicate.clientNo}\nهل تريد الحفظ رغم التكرار؟`); if(!ok) return; }
    try{
      if(editClientId) {
        const existing = clients.find(c=>c.id===editClientId);
        const saveData = {...form}; delete saveData.id;
        // الموظف لا يعدل جوال المالك — يُحفظ الرقم الأصلي
        if(!isManager && form.clientType==="مالك" && existing?.phone) saveData.phone = existing.phone;
        await updateDoc(doc(db,"clients",editClientId),{...saveData, updatedAt:new Date().toISOString(), isDuplicate:!!duplicate});
        const oldK = normPhone(existing?.phone);
        if(isManager && pk!==oldK){
          if(oldK){ const o=await getDoc(doc(db,"client_phones",oldK)); if(o.exists()&&o.data().clientId===editClientId) await deleteDoc(doc(db,"client_phones",oldK)); }
          if(pk&&!duplicate) await setDoc(doc(db,"client_phones",pk),{clientId:editClientId,clientNo:existing?.clientNo||0,name:form.name});
        }
      } else {
        const ref = doc(collection(db,"clients"));
        const localMax = clients.reduce((max,c)=>Math.max(max,c.clientNo||0),0);
        await runTransaction(db, async tx=>{
          const cref = doc(db,"meta","counters"), pref = pk?doc(db,"client_phones",pk):null;
          const cs = await tx.get(cref);
          const ps = pref ? await tx.get(pref) : null;
          const base = cs.exists() ? (cs.data().clientNo||0) : (isManager ? localMax : null);
          if(base===null) throw new Error("NO_COUNTER");
          const n = base+1;
          tx.set(cref,{clientNo:n},{merge:true});
          tx.set(ref,{...form, clientNo:n, createdAt:new Date().toISOString(), createdDate:new Date().toLocaleDateString("en-GB"), comments:[], isDuplicate:!!duplicate, assignedTo:currentUser?.uid||"", assignedToName:currentUser?.displayName||""});
          if(pref && !ps.exists()) tx.set(pref,{clientId:ref.id,clientNo:n,name:form.name});
        });
      }
    }catch(e){
      if(e?.message==="NO_COUNTER") return alert("يجب على المدير تشغيل «تجهيز الأمان» من صفحة الأمان قبل إضافة العملاء.");
      console.error(e); return alert("تعذّر الحفظ — تحقق من الاتصال أو الصلاحيات");
    }
    setForm(emptyClient); setShowForm(false); setEditClientId(null);
  };

  const assignClient = async () => {
    if(!assignTo||!assignModal) return;
    const emp = dbEmployees.find(u=>u.id===assignTo);
    const comment = { text:`📋 تم تحويل العميل إلى ${emp?.displayName||assignTo}${assignMsg?` — رسالة: ${assignMsg}`:""}`, date:new Date().toLocaleDateString("en-GB"), time:new Date().toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}), system:true };
    const existing = clients.find(c=>c.id===assignModal);
    const comments = [...(existing?.comments||[]), comment];
    await updateDoc(doc(db,"clients",assignModal),{ assignedTo:assignTo, assignedToName:emp?.displayName||assignTo, comments });
    setAssignModal(null); setAssignTo(""); setAssignMsg("");
  };

  const del = async (id) => {
    const c = clients.find(x=>x.id===id);
    await deleteDoc(doc(db,"clients",id));
    logAudit("حذف عميل",id,`${c?.name||""} ${c?.phone||""}`);
    try{ const k=normPhone(c?.phone); if(k){ const p=await getDoc(doc(db,"client_phones",k)); if(p.exists()&&p.data().clientId===id) await deleteDoc(doc(db,"client_phones",k)); } }catch(e){}
    setDelId(null);
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
            <div class="client-name">${hesc(c.name||"")}</div>
            <div class="client-phone">📞 ${hesc(c.phone||"")}</div>
            <div class="type-badge" style="background:${(ratingColor[c.clientType]||"#2a4d9b")}18;color:${ratingColor[c.clientType]||"#2a4d9b"};border:1px solid ${ratingColor[c.clientType]||"#2a4d9b"}33">${hesc(c.clientType||"")}</div>
          </div>
          <div class="num-badge">#${hesc(c.clientNo||"—")}</div>
        </div>
        <div class="divider"></div>
        <div class="info-grid">
          <div class="info-box"><div class="info-lbl">📍 الحي / الموقع</div><div class="info-val">${hesc(c.area||"—")}</div></div>
          <div class="info-box"><div class="info-lbl">💰 الميزانية</div><div class="info-val">${c.budget?Number(c.budget).toLocaleString()+" ﷼":"—"}</div></div>
          <div class="info-box"><div class="info-lbl">📋 الحالة</div><div class="info-val">${hesc(c.clientStatus||"معلق")}</div></div>
          <div class="info-box"><div class="info-lbl">📅 تاريخ التسجيل</div><div class="info-val">${hesc(c.createdDate||"—")}</div></div>
          ${c.ownerPropertyType?`<div class="info-box"><div class="info-lbl">🏠 نوع العقار</div><div class="info-val">${hesc(c.ownerPropertyType)}</div></div>`:""}
          ${c.investorType?`<div class="info-box"><div class="info-lbl">📈 الاستثمار</div><div class="info-val">${hesc(c.investorType)}</div></div>`:""}
          ${c.paymentType?`<div class="info-box"><div class="info-lbl">💳 طريقة الدفع</div><div class="info-val">${hesc(c.paymentType)}</div></div>`:""}
          ${c.contactDate?`<div class="info-box"><div class="info-lbl">📞 تاريخ التواصل</div><div class="info-val">${hesc(c.contactDate)}</div></div>`:""}
        </div>
        ${c.notes?`<div class="info-box" style="margin-bottom:14px;background:#f5f8ff;border-radius:8px;padding:10px 12px;"><div class="info-lbl">💬 ملاحظات</div><div style="font-size:11px;color:#1e3a7a;line-height:1.7;margin-top:3px">${hesc(c.notes)}</div></div>`:""}
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

  // الموظف يشوف فقط العملاء المحوّلين له صراحةً — المدير يشوف الكل
  const myClients = isManager 
    ? clients 
    : (currentUser ? clients.filter(c => c.assignedTo === currentUser.uid) : []);

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
                      {(isManager || c.clientType!=="مالك" || c.assignedTo===currentUser?.uid) ? (
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
                  {isManager&&<button onClick={()=>openEdit(c)} style={{background:"rgba(74,158,255,.1)",border:"1px solid rgba(74,158,255,.2)",color:darkMode?"#7ab8ff":"#2a4d9b",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:11,fontFamily:"'Cairo',sans-serif",fontWeight:600}}>✏️ تعديل</button>}
                  {isManager&&<button onClick={()=>printClientCard(c)} style={{background:"rgba(74,158,255,.08)",border:"1px solid rgba(74,158,255,.15)",color:darkMode?"#7ab8ff":"#2a4d9b",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:11,fontFamily:"'Cairo',sans-serif",fontWeight:600}}>🖨️ طباعة</button>}
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
                {(!isManager && editClientId && form.clientType==="مالك") ? (
                  // موظف يعدل مالك موجود
                  (() => {
                    const existingClient = clients.find(c=>c.id===editClientId);
                    const addedByMe = existingClient?.assignedTo === currentUser?.uid;
                    return addedByMe ? (
                      // هو اللي أضافه — يشوف الرقم بس ما يعدله
                      <div style={{background:darkMode?"#0a1538":"#f0f4fc",border:"1px solid rgba(74,158,255,.2)",borderRadius:10,padding:"9px 12px",fontSize:13,color:darkMode?"#e8eeff":"#1e3a7a",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                        <span>{existingClient?.phone||"—"}</span>
                        <span style={{fontSize:10,color:"#5a6a90",fontWeight:400}}>🔒 لا يمكن التعديل</span>
                      </div>
                    ) : (
                      // محوّل له من شخص آخر — محجوب
                      <div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:10,padding:"9px 12px",fontSize:12,color:"#ef4444",fontWeight:600}}>🔒 جوال المالك محجوب</div>
                    );
                  })()
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
                {dbEmployees.length===0 ? (
                  <div style={{textAlign:"center",padding:"20px",color:darkMode?"#7ab8ff":"#5a6a90",fontSize:13}}>لا يوجد موظفين — أضف موظفين أولاً من صفحة إدارة الموظفين</div>
                ) : dbEmployees.map(emp=>(
                  <button key={emp.id} onClick={()=>setAssignTo(emp.id)} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",borderRadius:10,border:`2px solid ${assignTo===emp.id?"#4a9eff":"rgba(74,158,255,.2)"}`,background:assignTo===emp.id?"rgba(74,158,255,.1)":"transparent",color:darkMode?"#e8eeff":"#1e3a7a",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer",textAlign:"right",width:"100%"}}>
                    <div style={{width:34,height:34,borderRadius:9,background:"rgba(74,158,255,.12)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>👤</div>
                    <div style={{flex:1}}>
                      <div>{emp.displayName}</div>
                      <div style={{fontSize:10,color:darkMode?"rgba(122,184,255,.6)":"#8899bb",fontWeight:400}}>{emp.email||emp.username}</div>
                    </div>
                    {assignTo===emp.id&&<span style={{color:"#4a9eff",fontSize:18}}>✓</span>}
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

  const del = async (id) => { const pv=providers.find(x=>x.id===id); await deleteDoc(doc(db,"providers",id)); logAudit("حذف مزود خدمة",id,pv?.name||""); setDelId(null); };

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
            <div class="prov-name">${hesc(p.name||"")}</div>
            <div class="prov-phone">📞 ${hesc(p.phone||"")}</div>
            <div class="badges">
              <span class="badge" style="background:rgba(74,158,255,.1);color:#2a4d9b;border:1px solid rgba(74,158,255,.25)">${hesc(p.specialty||"")}</span>
              <span class="badge" style="background:${ratingColor[p.rating]||"#2a4d9b"}18;color:${ratingColor[p.rating]||"#2a4d9b"};border:1px solid ${ratingColor[p.rating]||"#2a4d9b"}33">${hesc(p.rating||"")}</span>
            </div>
          </div>
        </div>
        <div class="divider"></div>
        ${p.notes?`<div class="info-row"><div><div class="info-lbl">📝 ملاحظات</div><div class="info-val">${hesc(p.notes)}</div></div></div>`:""}
        ${p.works&&p.works.length>0?`
          <div class="works-title">🛠️ الأعمال المنجزة (${p.works.length})</div>
          ${p.works.slice(0,5).map(w=>`<div class="work-item"><span>${hesc(w.text)}</span><span style="color:#8899bb;font-size:10px">${hesc(w.date)}</span></div>`).join("")}
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

// ── Employees Page (Firebase Auth accounts) ───────────────────────────────────
const PW_MIN = 10;
function EmployeesPage({ lang, T, darkMode, currentUser }) {
  const [users, setUsers] = useState([]);
  const [legacy, setLegacy] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const emptyForm = { displayName:"", email:"", password:"", username:"" };
  const [form, setForm] = useState(emptyForm);
  const [err, setErr] = useState("");
  const [msg, setMsg] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(()=>{
    const unsub = onSnapshot(query(collection(db,"users"),where("role","==","employee")), snap=>{
      setUsers(snap.docs.map(d=>({id:d.id,...d.data()})));
      setLoaded(true);
    }, ()=>setLoaded(true));
    return ()=>unsub();
  },[]);
  const loadLegacy = ()=>getDocs(collection(db,"employees")).then(s=>setLegacy(s.docs.map(d=>{ const x=d.data(); return {id:d.id,displayName:x.displayName||"",username:x.username||""}; }))).catch(()=>setLegacy([]));
  useEffect(()=>{ loadLegacy(); },[]);

  const flash = m => { setMsg(m); setTimeout(()=>setMsg(""),4000); };
  const create = async () => {
    if(busy) return;
    setErr("");
    const displayName=form.displayName.trim(), email=form.email.trim().toLowerCase(), password=form.password;
    if(!displayName||!email||!password) return setErr("الاسم والبريد وكلمة السر مطلوبة");
    if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return setErr("البريد الإلكتروني غير صحيح");
    if(password.length<PW_MIN) return setErr(`كلمة السر لازم ${PW_MIN} أحرف على الأقل`);
    if(/^\d+$/.test(password)) return setErr("كلمة السر لا يجوز أن تكون أرقاماً فقط");
    setBusy(true);
    let sec=null;
    try{
      // تطبيق ثانوي حتى لا يخرج المدير من حسابه أثناء إنشاء حساب الموظف
      sec = initializeApp(db.app.options, "sec-"+Date.now());
      const sa = getAuth(sec);
      const cred = await createUserWithEmailAndPassword(sa,email,password);
      const uid = cred.user.uid;
      await signOut(sa);
      await setDoc(doc(db,"users",uid),{role:"employee",active:true,displayName,email,username:form.username.trim(),createdAt:new Date().toISOString()});
      logAudit("إنشاء حساب موظف",uid,email);
      setForm(emptyForm); setShowForm(false);
      flash("✔ تم إنشاء حساب الموظف — سلّمه البريد وكلمة السر المؤقتة، ويقدر يغيّرها بعد الدخول");
    }catch(e){
      setErr(e?.code==="auth/email-already-in-use"?"هذا البريد مسجل مسبقاً. إذا كان موظفاً سابقاً فعّل حسابه من القائمة أو أرسل له رابط إعادة تعيين."
        :e?.code==="auth/weak-password"?"كلمة السر ضعيفة":"تعذّر إنشاء الحساب — تحقق من الاتصال وصلاحياتك");
    }
    try{ if(sec) await deleteApp(sec); }catch(e){}
    setBusy(false);
  };
  const toggle = async u => { try{ await updateDoc(doc(db,"users",u.id),{active:u.active===false}); logAudit(u.active===false?"تفعيل موظف":"تعطيل موظف",u.id,u.email||""); }catch(e){ flash("تعذّر التحديث"); } };
  const reset = async u => { try{ await sendPasswordResetEmail(auth,u.email); flash("✔ أُرسل رابط إعادة تعيين كلمة السر إلى "+u.email); }catch(e){ flash("تعذّر الإرسال"); } };
  const delLegacy = async l => {
    if(!window.confirm(`حذف الحساب القديم «${l.displayName}»؟ (هذا الحساب القديم كلمة سره محفوظة نصاً)`)) return;
    try{ await deleteDoc(doc(db,"employees",l.id)); loadLegacy(); }catch(e){ flash("تعذّر الحذف"); }
  };
  const f = key => e => setForm(p=>({...p,[key]:e.target.value}));
  const IST = {width:"100%",boxSizing:"border-box",background:darkMode?"#0a1538":"#f5f8ff",border:"1px solid rgba(74,158,255,.2)",borderRadius:10,padding:"9px 12px",color:darkMode?"#e8eeff":"#1e3a7a",fontFamily:"'Cairo',sans-serif",fontSize:13};
  const Lbl = ({c})=><div style={{fontSize:11,color:darkMode?"#7ab8ff":"#5a6a90",marginBottom:5,fontWeight:600}}>{c}</div>;
  const rowBtn = (c)=>({background:`${c}14`,border:`1px solid ${c}33`,color:c,borderRadius:9,padding:"7px 12px",cursor:"pointer",fontSize:12,fontFamily:"'Cairo',sans-serif",fontWeight:600});

  return (
    <div style={{paddingTop:96,minHeight:"100vh",background:T.bg}}>
      <div style={{background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",padding:"28px 24px 24px"}}>
        <div style={{maxWidth:900,margin:"0 auto",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
          <div>
            <div style={{fontWeight:900,fontSize:22,color:"#fff",marginBottom:3}}>👨‍💼 إدارة الموظفين</div>
            <div style={{fontSize:12,color:"rgba(255,255,255,.5)"}}>للمدير فقط — حسابات الموظفين بتسجيل دخول مشفّر (Firebase Authentication)</div>
          </div>
          <button onClick={()=>{setForm(emptyForm);setErr("");setShowForm(true);}} style={{background:"rgba(255,255,255,.15)",border:"1px solid rgba(255,255,255,.3)",color:"#fff",borderRadius:11,padding:"10px 22px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>+ إضافة موظف</button>
        </div>
      </div>

      <div style={{maxWidth:900,margin:"0 auto",padding:"24px 22px"}}>
        {msg&&<div style={{background:"#16a34a14",border:"1px solid #16a34a30",color:"#16a34a",borderRadius:10,padding:"10px 14px",fontSize:12,fontWeight:700,marginBottom:14}}>{msg}</div>}
        <div style={{background:darkMode?"rgba(74,158,255,.08)":"rgba(74,158,255,.06)",border:"1px solid rgba(74,158,255,.2)",borderRadius:14,padding:"16px 18px",marginBottom:20,display:"flex",alignItems:"center",gap:14}}>
          <div style={{width:44,height:44,borderRadius:12,background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>👑</div>
          <div>
            <div style={{fontWeight:900,fontSize:14,color:darkMode?"#e8eeff":"#1e3a7a"}}>{currentUser?.displayName} — المدير</div>
            <div style={{fontSize:11,color:darkMode?"#7ab8ff":"#5a6a90"}} dir="ltr">{currentUser?.email}</div>
          </div>
        </div>

        {legacy.length>0&&(
          <div style={{background:"#ef444410",border:"1px solid #ef444440",borderRadius:14,padding:"14px 16px",marginBottom:20}}>
            <div style={{fontWeight:900,fontSize:13,color:"#ef4444",marginBottom:4}}>⚠️ حسابات قديمة غير آمنة ({legacy.length})</div>
            <div style={{fontSize:12,color:darkMode?"#fca5a5":"#b91c1c",marginBottom:10,lineHeight:1.8}}>هذه الحسابات كلمات سرها محفوظة نصاً في قاعدة البيانات. أنشئ لكل موظف حساباً جديداً ثم احذف القديم.</div>
            {legacy.map(l=>(
              <div key={l.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,padding:"8px 0",borderTop:"1px dashed #ef444430",flexWrap:"wrap"}}>
                <div style={{fontSize:13,color:darkMode?"#e8eeff":"#1e3a7a",fontWeight:700}}>{l.displayName} <span style={{fontWeight:400,fontSize:11,color:darkMode?"#7ab8ff":"#5a6a90"}}>({l.username})</span></div>
                <div style={{display:"flex",gap:6}}>
                  <button onClick={()=>{setForm({displayName:l.displayName,email:"",password:"",username:l.username});setErr("");setShowForm(true);}} style={rowBtn("#2a4d9b")}>إنشاء حساب جديد</button>
                  <button onClick={()=>delLegacy(l)} style={rowBtn("#ef4444")}>🗑️ حذف القديم</button>
                </div>
              </div>
            ))}
          </div>
        )}

        {!loaded ? <div style={{textAlign:"center",padding:40,color:T.text3}}>جاري التحميل...</div> : users.length===0 ? (
          <div style={{textAlign:"center",padding:60,color:T.text3}}><div style={{fontSize:46,marginBottom:10}}>👤</div><div style={{fontSize:13}}>لا يوجد موظفون بعد</div></div>
        ) : (
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            {users.map(u=>(
              <div key={u.id} style={{background:darkMode?"linear-gradient(160deg,#0f1f4a,#1a2d6b)":"white",border:"1px solid rgba(74,158,255,.12)",borderRadius:14,padding:"16px 18px",boxShadow:"0 2px 8px rgba(30,58,122,.06)",display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,flexWrap:"wrap",opacity:u.active===false?.6:1}}>
                <div style={{display:"flex",alignItems:"center",gap:14}}>
                  <div style={{width:46,height:46,borderRadius:12,background:"rgba(74,158,255,.1)",border:"1px solid rgba(74,158,255,.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,flexShrink:0}}>👤</div>
                  <div>
                    <div style={{fontWeight:900,fontSize:15,color:darkMode?"#e8eeff":"#1e3a7a",marginBottom:2}}>{u.displayName} {u.active===false&&<span style={{fontSize:10,color:"#ef4444",fontWeight:700}}>• موقوف</span>}</div>
                    <div style={{fontSize:11,color:darkMode?"#7ab8ff":"#5a6a90"}} dir="ltr">{u.email}</div>
                  </div>
                </div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                  <button onClick={()=>reset(u)} style={rowBtn("#2a4d9b")}>🔑 رابط تغيير كلمة السر</button>
                  <button onClick={()=>toggle(u)} style={rowBtn(u.active===false?"#16a34a":"#ef4444")}>{u.active===false?"✔ تفعيل":"⏸ إيقاف"}</button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {showForm&&(
        <div style={{position:"fixed",inset:0,background:"#000b",zIndex:600,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={()=>setShowForm(false)}>
          <div style={{background:darkMode?"linear-gradient(160deg,#0f1f4a,#1a2d6b)":"white",border:"1px solid rgba(74,158,255,.25)",borderRadius:22,padding:24,maxWidth:420,width:"100%"}} onClick={e=>e.stopPropagation()}>
            <div style={{fontWeight:900,fontSize:16,color:darkMode?"#e8eeff":"#1e3a7a",marginBottom:20,paddingBottom:14,borderBottom:"1px solid rgba(74,158,255,.1)"}}>👤 إضافة موظف جديد</div>
            <div style={{display:"flex",flexDirection:"column",gap:12,marginBottom:16}}>
              <div><Lbl c="الاسم الظاهر (مثال: محمد)"/><input value={form.displayName} onChange={f("displayName")} style={IST} placeholder="اسم الموظف"/></div>
              <div><Lbl c="البريد الإلكتروني (يستخدمه للدخول)"/><input type="email" value={form.email} onChange={f("email")} style={IST} placeholder="name@example.com" dir="ltr" autoComplete="off"/></div>
              <div><Lbl c={`كلمة سر مؤقتة (${PW_MIN} أحرف على الأقل، ليست أرقاماً فقط)`}/><input type="text" value={form.password} onChange={f("password")} style={IST} dir="ltr" autoComplete="new-password"/></div>
              <div><Lbl c="اليوزر القديم (اختياري — لربط عملائه السابقين به)"/><input value={form.username} onChange={f("username")} style={IST} placeholder="مثال: emp2" dir="ltr"/></div>
            </div>
            {err&&<div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#ef4444",marginBottom:12}}>⚠️ {err}</div>}
            <div style={{display:"flex",gap:10}}>
              <button onClick={create} disabled={busy} style={{flex:1,background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",color:"#fff",border:"none",borderRadius:11,padding:"12px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:14,cursor:"pointer",opacity:busy?.7:1}}>{busy?"جاري الإنشاء...":"✅ إنشاء الحساب"}</button>
              <button onClick={()=>setShowForm(false)} style={{background:darkMode?"#0a1538":"#f5f8ff",border:"1px solid rgba(74,158,255,.2)",color:darkMode?"#7ab8ff":"#5a6a90",borderRadius:11,padding:"12px 18px",fontFamily:"'Cairo',sans-serif",cursor:"pointer"}}>إلغاء</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Security Page (admin) ─────────────────────────────────────────────────────
const normPhone = p => String(p||"").replace(/\D/g,"");
const BACKUP_COLS = ["properties","properties_private","clients","client_phones","providers","users","pm_props","pm_tenants","pm_contracts","pm_payments","pm_maint","pm_transfers","pm_vouchers","pm_meta","meta"];
async function downloadBackup(){
  const out = { exportedAt:new Date().toISOString(), collections:{} };
  for(const name of BACKUP_COLS){
    try{
      const snap = await getDocs(collection(db,name));
      out.collections[name] = Object.fromEntries(snap.docs.map(d=>[d.id,d.data()]));
    }catch(e){ out.collections[name] = { __error: e?.code||"failed" }; }
  }
  const blob = new Blob([JSON.stringify(out)],{type:"application/json"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "backup-"+new Date().toISOString().slice(0,10)+".json";
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(a.href),5000);
  await setDoc(doc(db,"meta","backup"),{lastAt:new Date().toISOString()},{merge:true});
  logAudit("تنزيل نسخة احتياطية","backup","");
  return Object.values(out.collections).reduce((n,c)=>n+Object.keys(c).length,0);
}
function SecurityPage({ lang, T, darkMode, currentUser, rawProps }) {
  const [cur,setCur]=useState(""); const [np,setNp]=useState(""); const [np2,setNp2]=useState(""); const [pwMsg,setPwMsg]=useState(null);
  const [stat,setStat]=useState(null); const [log,setLog]=useState([]); const [busy,setBusy]=useState(false);
  const [lastBk,setLastBk]=useState(null); const [bkMsg,setBkMsg]=useState("");
  useEffect(()=>{ getDoc(doc(db,"meta","backup")).then(d=>setLastBk(d.exists()?d.data().lastAt:null)).catch(()=>{}); },[]);
  const [audit,setAudit]=useState(null);
  const loadAudit = async()=>{ try{ const s=await getDocs(query(collection(db,"audit_logs"),orderBy("ts","desc"),limit(50))); setAudit(s.docs.map(d=>({id:d.id,...d.data()}))); }catch(e){ setAudit([]); } };
  const doBackup = async()=>{ setBusy(true); setBkMsg(""); try{ const n=await downloadBackup(); setLastBk(new Date().toISOString()); setBkMsg(`✔ تم تنزيل النسخة (${n} سجل). احفظ الملف في مكان آمن خارج الجهاز.`); }catch(e){ setBkMsg("تعذّر إنشاء النسخة"); } setBusy(false); };
  const txt=darkMode?"#e8eeff":"#1e3a7a", sub=darkMode?"#7ab8ff":"#5a6a90";
  const IST = {width:"100%",boxSizing:"border-box",background:darkMode?"#0a1538":"#f5f8ff",border:"1px solid rgba(74,158,255,.2)",borderRadius:10,padding:"9px 12px",color:txt,fontFamily:"'Cairo',sans-serif",fontSize:13};
  const card = {background:darkMode?"linear-gradient(160deg,#0f1f4a,#1a2d6b)":"white",border:"1px solid rgba(74,158,255,.15)",borderRadius:16,padding:"18px 20px",marginBottom:16};
  const btn = {background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",color:"#fff",border:"none",borderRadius:11,padding:"11px 20px",fontFamily:"'Cairo',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"};

  const changePw = async () => {
    setPwMsg(null);
    if(np.length<PW_MIN) return setPwMsg({ok:false,t:`كلمة السر الجديدة لازم ${PW_MIN} أحرف على الأقل`});
    if(/^\d+$/.test(np)) return setPwMsg({ok:false,t:"لا يجوز أن تكون أرقاماً فقط"});
    if(np!==np2) return setPwMsg({ok:false,t:"التأكيد لا يطابق كلمة السر الجديدة"});
    try{
      const u=auth.currentUser;
      await reauthenticateWithCredential(u,EmailAuthProvider.credential(u.email,cur));
      await updatePassword(u,np);
      setCur(""); setNp(""); setNp2(""); setPwMsg({ok:true,t:"✔ تم تغيير كلمة السر"});
    }catch(e){ setPwMsg({ok:false,t:e?.code==="auth/invalid-credential"||e?.code==="auth/wrong-password"?"كلمة السر الحالية غير صحيحة":"تعذّر التغيير — أعد المحاولة"}); }
  };

  const scan = async () => {
    setBusy(true);
    try{
      const [cs,ps,us,ctr] = await Promise.all([
        getDocs(collection(db,"clients")), getDocs(collection(db,"client_phones")), getDocs(collection(db,"users")), getDoc(doc(db,"meta","counters")),
      ]);
      let legacy=0; try{ legacy=(await getDocs(collection(db,"employees"))).size; }catch(e){}
      const clients=cs.docs.map(d=>({id:d.id,...d.data()})), phones=new Set(ps.docs.map(d=>d.id)), uids=new Set(us.docs.map(d=>d.id));
      const maxNo=clients.reduce((m,c)=>Math.max(m,c.clientNo||0),0);
      setStat({
        legacy,
        ownerExposed:(rawProps||[]).filter(p=>p.ownerName||p.ownerPhone).length,
        noIndex:clients.filter(c=>normPhone(c.phone)&&!phones.has(normPhone(c.phone))).length,
        counterOk:ctr.exists()&&(ctr.data().clientNo||0)>=maxNo,
        unmapped:clients.filter(c=>c.assignedTo&&!uids.has(c.assignedTo)).length,
        clients:clients.length,
      });
    }catch(e){ setLog(l=>[...l,"تعذّر الفحص: "+(e?.code||e?.message||"")]); }
    setBusy(false);
  };
  useEffect(()=>{ scan(); /* eslint-disable-next-line */ },[]);

  const migrate = async () => {
    if(busy) return;
    if(!window.confirm("تشغيل تجهيز الأمان؟ العملية آمنة ويمكن تكرارها.")) return;
    setBusy(true); setLog([]);
    const L = m=>setLog(l=>[...l,m]);
    try{
      // 1) بيانات الملاك من الوثيقة العامة إلى مجموعة خاصة
      let n1=0;
      for(const p of (rawProps||[])){
        if(!(p.ownerName||p.ownerPhone)) continue;
        await setDoc(doc(db,"properties_private",p.id),{ownerName:p.ownerName||"",ownerPhone:p.ownerPhone||""},{merge:true});
        await updateDoc(doc(db,"properties",p.id),{ownerName:deleteField(),ownerPhone:deleteField()});
        n1++;
      }
      L(`✔ نُقلت بيانات ${n1} عقار إلى مكان خاص`);
      // 2) فهرس الجوالات وعدّاد العملاء + ربط العملاء بالموظفين
      const [cs,us]=await Promise.all([getDocs(collection(db,"clients")),getDocs(collection(db,"users"))]);
      const clients=cs.docs.map(d=>({id:d.id,...d.data()})), users=us.docs.map(d=>({id:d.id,...d.data()}));
      let n2=0, dup=0;
      for(const c of clients){
        const k=normPhone(c.phone); if(!k) continue;
        const ref=doc(db,"client_phones",k), ex=await getDoc(ref);
        if(ex.exists()){ if(ex.data().clientId!==c.id) dup++; continue; }
        await setDoc(ref,{clientId:c.id,clientNo:c.clientNo||0,name:c.name||""}); n2++;
      }
      const maxNo=clients.reduce((m,c)=>Math.max(m,c.clientNo||0),0);
      const ctrRef=doc(db,"meta","counters"), ctr=await getDoc(ctrRef);
      if(!ctr.exists()||(ctr.data().clientNo||0)<maxNo) await setDoc(ctrRef,{clientNo:maxNo},{merge:true});
      L(`✔ فهرس الجوالات: أُضيف ${n2} (مكرر مسبقاً: ${dup}) — عدّاد العملاء: ${maxNo}`);
      let n3=0;
      for(const c of clients){
        if(!c.assignedTo||users.some(u=>u.id===c.assignedTo)) continue;
        const target = c.assignedTo==="Khalil" ? {id:currentUser.uid,displayName:currentUser.displayName}
          : users.find(u=>u.username&&u.username===c.assignedTo);
        if(!target) continue;
        await updateDoc(doc(db,"clients",c.id),{assignedTo:target.id,assignedToName:target.displayName||""}); n3++;
      }
      L(`✔ رُبط ${n3} عميل بحسابات الموظفين الجديدة`);
    }catch(e){ L("⚠️ توقفت العملية: "+(e?.code||e?.message||"خطأ")); }
    setBusy(false); scan();
  };

  const Row = ({ok,t,hint}) => (
    <div style={{display:"flex",gap:10,alignItems:"flex-start",padding:"9px 0",borderBottom:"1px dashed rgba(74,158,255,.15)"}}>
      <span style={{fontSize:16}}>{ok?"✅":"⚠️"}</span>
      <div style={{flex:1}}><div style={{fontSize:13,color:txt,fontWeight:700}}>{t}</div>{!ok&&hint&&<div style={{fontSize:11,color:sub,marginTop:2}}>{hint}</div>}</div>
    </div>
  );

  return (
    <div style={{paddingTop:96,minHeight:"100vh",background:T.bg}}>
      <div style={{background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",padding:"28px 24px 24px"}}>
        <div style={{maxWidth:900,margin:"0 auto"}}>
          <div style={{fontWeight:900,fontSize:22,color:"#fff",marginBottom:3}}>🛡️ الأمان</div>
          <div style={{fontSize:12,color:"rgba(255,255,255,.5)"}}>للمدير فقط — كلمة السر وتجهيز البيانات للحماية</div>
        </div>
      </div>
      <div style={{maxWidth:900,margin:"0 auto",padding:"24px 22px"}}>
        <div style={card}>
          <div style={{fontWeight:900,fontSize:15,color:txt,marginBottom:4}}>🔑 تغيير كلمة سر حسابي</div>
          <div style={{fontSize:11,color:sub,marginBottom:12}} dir="ltr">{currentUser?.email}</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:10,marginBottom:12}}>
            <input type="password" autoComplete="current-password" placeholder="كلمة السر الحالية" value={cur} onChange={e=>setCur(e.target.value)} style={IST} dir="ltr"/>
            <input type="password" autoComplete="new-password" placeholder={`الجديدة (${PW_MIN}+ أحرف)`} value={np} onChange={e=>setNp(e.target.value)} style={IST} dir="ltr"/>
            <input type="password" autoComplete="new-password" placeholder="تأكيد الجديدة" value={np2} onChange={e=>setNp2(e.target.value)} style={IST} dir="ltr"/>
          </div>
          {pwMsg&&<div style={{fontSize:12,fontWeight:700,color:pwMsg.ok?"#16a34a":"#ef4444",marginBottom:10}}>{pwMsg.t}</div>}
          <button onClick={changePw} style={btn}>تغيير كلمة السر</button>
        </div>

        <div style={card}>
          <div style={{fontWeight:900,fontSize:15,color:txt,marginBottom:10}}>🧰 حالة تجهيز الأمان</div>
          {!stat?<div style={{fontSize:12,color:sub}}>جاري الفحص...</div>:(
            <>
              <Row ok={stat.legacy===0} t={stat.legacy===0?"لا توجد حسابات موظفين قديمة (كلمات سر نصية)":`${stat.legacy} حساب موظف قديم بكلمة سر نصية`} hint="من صفحة «إدارة الموظفين» أنشئ حسابات جديدة ثم احذف القديمة."/>
              <Row ok={stat.ownerExposed===0} t={stat.ownerExposed===0?"بيانات الملاك غير ظاهرة في الوثائق العامة":`${stat.ownerExposed} عقار ببيانات مالك ظاهرة للعامة`} hint="اضغط «تشغيل تجهيز الأمان»."/>
              <Row ok={stat.noIndex===0} t={stat.noIndex===0?"فهرس منع تكرار الجوالات جاهز":`${stat.noIndex} عميل غير مفهرس لمنع التكرار`} hint="اضغط «تشغيل تجهيز الأمان»."/>
              <Row ok={stat.counterOk} t={stat.counterOk?"عدّاد أرقام العملاء جاهز":"عدّاد أرقام العملاء غير مهيأ"} hint="اضغط «تشغيل تجهيز الأمان» — بدونه قد تتكرر أرقام العملاء."/>
              <Row ok={stat.unmapped===0} t={stat.unmapped===0?"كل العملاء مرتبطون بحسابات حالية":`${stat.unmapped} عميل مرتبط بموظف غير موجود`} hint="أنشئ حساب الموظف باليوزر القديم نفسه ثم شغّل التجهيز، أو أعد تحويل العميل يدوياً."/>
            </>
          )}
          <div style={{display:"flex",gap:10,marginTop:14,flexWrap:"wrap"}}>
            <button onClick={migrate} disabled={busy} style={{...btn,opacity:busy?.7:1}}>{busy?"جاري التنفيذ...":"🛠️ تشغيل تجهيز الأمان"}</button>
            <button onClick={scan} disabled={busy} style={{...btn,background:darkMode?"#0a1538":"#f5f8ff",color:sub,border:"1px solid rgba(74,158,255,.2)"}}>🔄 إعادة الفحص</button>
          </div>
          {log.length>0&&<div style={{marginTop:12,fontSize:12,color:txt,lineHeight:2}}>{log.map((l,i)=><div key={i}>{l}</div>)}</div>}
        </div>
        <div style={{marginTop:16,padding:18,borderRadius:16,border:"1px solid rgba(74,158,255,.2)",background:darkMode?"#0f1d4a":"#fff",textAlign:"center"}}>
          <div style={{fontWeight:800,fontSize:15,color:txt,marginBottom:6}}>💾 النسخ الاحتياطي</div>
          <div style={{fontSize:12,color:sub,marginBottom:12}}>{lastBk?`آخر نسخة: ${new Date(lastBk).toLocaleDateString("en-GB")}`:"لم تؤخذ نسخة بعد"} — يُنصح بنسخة كل شهر</div>
          <button onClick={doBackup} disabled={busy} style={{...btn,opacity:busy?.7:1}}>⬇️ تنزيل نسخة احتياطية الآن</button>
          {bkMsg&&<div style={{marginTop:10,fontSize:12,color:txt}}>{bkMsg}</div>}
        </div>
        <div style={{marginTop:16,padding:18,borderRadius:16,border:"1px solid rgba(74,158,255,.2)",background:darkMode?"#0f1d4a":"#fff"}}>
          <div style={{fontWeight:800,fontSize:15,color:txt,marginBottom:10,textAlign:"center"}}>📜 سجل التدقيق (آخر 50 عملية)</div>
          {audit===null?<div style={{textAlign:"center"}}><button onClick={loadAudit} style={btn}>عرض السجل</button></div>
            :audit.length===0?<div style={{textAlign:"center",fontSize:12,color:sub}}>لا توجد عمليات مسجلة بعد</div>
            :<div style={{fontSize:12,color:txt,lineHeight:1.9}}>{audit.map(a=>(<div key={a.id} style={{padding:"6px 0",borderBottom:"1px dashed rgba(74,158,255,.15)"}}><strong>{a.action}</strong> — {a.detail||a.target}<div style={{fontSize:10,color:sub}}>{new Date(a.ts).toLocaleString("en-GB")} · {a.email}</div></div>))}</div>}
        </div>
      </div>
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
// ═══════════════════════════════════════════════════════════════════════════════
// 🏛️ إدارة الأملاك — Property Management (للمدير فقط)
// Firestore: pm_props, pm_tenants, pm_contracts, pm_payments, pm_maint,
//            pm_transfers, pm_vouchers, pm_meta/counters (ترقيم السندات والعقود)
// ═══════════════════════════════════════════════════════════════════════════════
const PM_FIRM = "مؤسسة خالد محمد عبدالغفور الشيخ";
const PM_LOGO = "https://res.cloudinary.com/dumtp0krl/image/upload/v1778958489/WhatsApp_Image_2026-05-16_at_9.59.47_PM_zhmw6y.jpg";

const PM_CSS = `
.pm{--navy:#1e3a7a; --navy2:#2a4d9b; --navy3:#0f1f4a; --blue:#4a9eff; --blue2:#7ab8ff; --bg:#f5f8ff; --card:#ffffff; --soft:#edf1fb; --text:#1e3a7a; --muted:#5a6a90; --line:rgba(74,158,255,.18); --green:#16a34a; --red:#ef4444; --amber:#b45309; --shadow:0 2px 12px rgba(30,58,122,.07);}
.pm.dark{--bg:#0f1f4a; --card:#1a2d6b; --soft:#0a1538; --text:#e8eeff; --muted:#7ab8ff; --line:rgba(74,158,255,.2); --green:#4ade80; --red:#f87171; --amber:#fbbf24; --shadow:none;}
.pm *{box-sizing:border-box;margin:0;padding:0;font-family:'Cairo',sans-serif;}
.pm{background:var(--bg);color:var(--text);transition:background .25s,color .25s;}
.pm .wrap{max-width:1150px;margin:0 auto;padding:22px 18px 60px;}
.pm h2.title{font-size:20px;font-weight:900;margin-bottom:4px;}
.pm p.sub{font-size:12px;color:var(--muted);margin-bottom:18px;}
.pm .row{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;flex-wrap:wrap;margin-bottom:6px;}
.pm .card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:16px 18px;box-shadow:var(--shadow);}
.pm .gridS{display:grid;grid-template-columns:repeat(auto-fit,minmax(165px,1fr));gap:12px;margin-bottom:16px;}
.pm .stat .ic{font-size:22px;margin-bottom:6px;}
.pm .stat .v{font-size:20px;font-weight:900;}
.pm .stat .l{font-size:11px;color:var(--muted);margin-top:2px;}
.pm .stat.red .v{color:var(--red);}
.pm .stat.green .v{color:var(--green);}
.pm .stat.amber .v{color:var(--amber);}
.pm .cols{display:grid;grid-template-columns:1.25fr 1fr;gap:14px;}
.pm .ch{font-weight:900;font-size:14px;margin-bottom:12px;display:flex;align-items:center;justify-content:space-between;}
.pm .alerts{max-height:430px;overflow-y:auto;}
.pm .alert{display:flex;align-items:center;gap:11px;padding:10px 12px;border-radius:12px;margin-bottom:8px;border:1px solid var(--line);background:var(--soft);}
.pm .alert .d{width:34px;height:34px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;}
.pm .alert.late .d{background:rgba(239,68,68,.12);}
.pm .alert.end .d{background:rgba(180,83,9,.12);}
.pm .alert.soon .d{background:rgba(74,158,255,.14);}
.pm .alert.mnt .d{background:rgba(122,184,255,.18);}
.pm .alert .t{flex:1;font-size:12px;font-weight:700;line-height:1.5;}
.pm .alert .t small{display:block;color:var(--muted);font-weight:400;font-size:11px;}
.pm .alert button{background:var(--navy);color:#fff;border:none;border-radius:8px;padding:5px 11px;font-size:11px;font-weight:700;cursor:pointer;white-space:nowrap;}
.pm .empty{text-align:center;color:var(--muted);font-size:12px;padding:20px;}
.pm .bars{display:flex;align-items:flex-end;gap:6px;height:130px;margin-top:6px;}
.pm .bar{flex:1;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;gap:3px;height:100%;}
.pm .bar i{display:block;width:100%;background:linear-gradient(180deg,var(--blue),var(--navy2));border-radius:6px 6px 0 0;min-height:3px;}
.pm .bar span{font-size:9px;color:var(--muted);}
.pm .bar em{font-size:9px;font-style:normal;color:var(--muted);font-weight:700;height:12px;}
.pm .occ{margin-top:16px;}
.pmprog{height:9px;background:var(--soft);border-radius:6px;overflow:hidden;margin-top:6px;}
.pmprog.sm{height:6px;margin:4px 0;}
.pmprog i{display:block;height:100%;background:linear-gradient(90deg,var(--blue),var(--navy2));border-radius:6px;}
.pm .pgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(310px,1fr));gap:14px;}
.pm .pc .head{display:flex;gap:12px;align-items:center;margin-bottom:12px;}
.pm .pc .pi{width:50px;height:50px;border-radius:14px;background:var(--soft);border:1px solid var(--line);display:flex;align-items:center;justify-content:center;font-size:24px;}
.pm .pc .pn{font-weight:900;font-size:15px;}
.pm .pc .pa{font-size:11px;color:var(--muted);}
.pm .units{display:flex;gap:5px;flex-wrap:wrap;margin:10px 0 4px;}
.pm .uc{font-size:10px;font-weight:700;border-radius:8px;padding:3px 9px;border:1px solid;}
.pm .uc.occ{color:var(--green);background:rgba(22,163,74,.1);border-color:rgba(22,163,74,.3);}
.pm .uc.vac{color:var(--muted);background:var(--soft);border-color:var(--line);}
.pm .kv{display:flex;justify-content:space-between;font-size:12px;padding:6px 0;border-bottom:1px dashed var(--line);gap:10px;}
.pm .kv span{color:var(--muted);}
.pm .kv b{font-weight:700;text-align:left;}
.pm .btn{background:var(--navy);color:#fff;border:none;border-radius:10px;padding:9px 16px;font-weight:700;font-size:12px;cursor:pointer;}
.pm .btn.blue{background:var(--blue);}
.pm .btn.ghost{background:var(--soft);color:var(--text);border:1px solid var(--line);}
.pm .btn.sm{padding:6px 11px;font-size:11px;border-radius:8px;}
.pm .btn.wa{background:rgba(37,211,102,.12);color:#16a34a;border:1px solid rgba(37,211,102,.35);}
.pm body.dark .btn.wa{color:#4ade80;}
.pm .pc .btns{display:flex;gap:8px;margin-top:12px;}
.pm .pc .btns .btn{flex:1;}
.pm .tw{overflow-x:auto;}
.pm table{width:100%;border-collapse:collapse;font-size:12px;min-width:820px;}
.pm th{text-align:right;color:var(--muted);font-weight:700;font-size:11px;padding:10px 10px;border-bottom:2px solid var(--line);}
.pm td{padding:11px 10px;border-bottom:1px solid var(--line);}
.pm tr:hover td{background:var(--soft);}
.pm .badge{display:inline-block;border-radius:20px;padding:3px 11px;font-size:11px;font-weight:700;border:1px solid;white-space:nowrap;}
.pm .b-green{color:var(--green);background:rgba(22,163,74,.1);border-color:rgba(22,163,74,.3);}
.pm .b-red{color:var(--red);background:rgba(239,68,68,.1);border-color:rgba(239,68,68,.3);}
.pm .b-amber{color:var(--amber);background:rgba(180,83,9,.1);border-color:rgba(180,83,9,.3);}
.pm .b-blue{color:var(--navy2);background:rgba(74,158,255,.12);border-color:rgba(74,158,255,.3);}
.pm .b-gray{color:var(--muted);background:var(--soft);border-color:var(--line);}
.pm body.dark .b-blue{color:var(--blue2);}
.pm .chips{display:flex;gap:7px;flex-wrap:wrap;margin-bottom:14px;}
.pm .chip{border:1px solid var(--line);background:var(--card);color:var(--muted);border-radius:20px;padding:6px 14px;font-size:12px;font-weight:700;cursor:pointer;}
.pm .chip.on{background:var(--navy);border-color:var(--navy);color:#fff;}
.pm .crow{display:flex;align-items:center;gap:14px;padding:12px 14px;border-radius:14px;margin-bottom:8px;background:var(--card);border:1px solid var(--line);flex-wrap:wrap;}
.pm .crow .who{flex:1;min-width:170px;}
.pm .crow .who b{font-size:13px;display:block;}
.pm .crow .who small{font-size:11px;color:var(--muted);}
.pm .crow .due{font-size:11px;color:var(--muted);min-width:100px;}
.pm .crow .amtbox{min-width:190px;}
.pm .crow .amt{font-weight:900;font-size:14px;}
.pm .crow .amtbox small{font-size:10px;color:var(--muted);}
.pm .crow .acts{display:flex;gap:6px;flex-wrap:wrap;}
.pm .tgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:14px;}
.pm .tc .tn{font-weight:900;font-size:15px;margin-bottom:2px;}
.pm .tc .tp{font-size:11px;color:var(--muted);margin-bottom:10px;}
.pm .tc .btns{display:flex;gap:8px;margin-top:12px;flex-wrap:wrap;}
.pm .modal{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:100;display:none;align-items:center;justify-content:center;padding:16px;}
.pm .modal.on{display:flex;}
.pm .mbox{background:var(--card);border:1px solid var(--line);border-radius:20px;padding:22px;width:100%;max-width:480px;max-height:92vh;overflow-y:auto;}
.pm .mbox h3{font-size:16px;font-weight:900;margin-bottom:14px;padding-bottom:12px;border-bottom:1px solid var(--line);}
.pm .fg{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px;}
.pm .fg .full{grid-column:1/-1;}
.pm label{display:block;font-size:11px;color:var(--muted);font-weight:700;margin-bottom:4px;}
.pm input,.pm select,.pm textarea{width:100%;background:var(--soft);border:1px solid var(--line);border-radius:10px;padding:9px 11px;color:var(--text);font-size:13px;}
.pm input:focus,.pm select:focus,.pm textarea:focus{outline:none;border-color:var(--blue);}
.pm .info{background:var(--soft);border:1px solid var(--line);border-radius:12px;padding:10px 12px;margin-bottom:12px;font-size:12px;line-height:1.9;}
.pm .pnote{font-size:12px;font-weight:700;color:var(--muted);background:var(--soft);border-radius:10px;padding:8px 12px;margin-bottom:14px;}
.pm .hrow{display:flex;align-items:center;gap:10px;padding:10px 0;border-bottom:1px dashed var(--line);font-size:12px;}
.pm .hrow .hm{flex:1;}
.pm .hrow small{display:block;color:var(--muted);font-size:10px;}
.pm .mgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:12px;}
.pm .mc .mt{font-weight:900;font-size:14px;margin-bottom:4px;}
.pm .mc .ms{font-size:11px;color:var(--muted);margin-bottom:10px;}
.pm .mc .btns{display:flex;gap:8px;margin-top:10px;}
.pmpr{background:#fff;color:#1e3a7a;direction:rtl;padding:22px;border:1px solid #d5def0;border-radius:12px;max-width:780px;font-family:'Cairo',sans-serif;}
.pmpr *{font-family:'Cairo',sans-serif;}
.pmpr .ph{display:flex;align-items:center;gap:12px;background:linear-gradient(135deg,#1e3a7a,#2a4d9b);color:#fff;border-radius:12px;padding:14px;margin-bottom:16px;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
.pmpr .ph img{width:48px;height:48px;border-radius:9px;background:#fff;object-fit:contain;}
.pmpr .ph b{font-size:14px;display:block;}
.pmpr .ph small{font-size:11px;opacity:.8;}
.pmpr .ph .rt{margin-right:auto;text-align:left;}
.pmpr h4{font-size:12px;color:#5a6a90;margin:14px 0 6px;}
.pmpr .kv{border-bottom:1px dashed #d5def0;}
.pmpr .kv span{color:#5a6a90;}
.pmpr .kv b{color:#1e3a7a;}
.pmpr table{min-width:0;font-size:11px;margin-bottom:6px;}
.pmpr th{background:#edf1fb;color:#5a6a90;padding:6px 8px;border-bottom:1px solid #d5def0;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
.pmpr td{padding:6px 8px;border-bottom:1px solid #e6ecf8;color:#1e3a7a;}
.pmpr tr:hover td{background:none;}
.pmpr .sum{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:14px 0;}
.pmpr .sum div{background:#edf1fb;border:1px solid #d5def0;border-radius:10px;padding:9px;text-align:center;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
.pmpr .sum small{display:block;font-size:10px;color:#5a6a90;}
.pmpr .sum b{font-size:14px;font-weight:900;}
.pmpr .sum .hl{background:#1e3a7a;}
.pmpr .sum .hl small{color:#7ab8ff;}
.pmpr .sum .hl b{color:#fff;}
.pmpr .big{text-align:center;margin:16px 0;font-size:22px;font-weight:900;border:2px dashed #4a9eff;border-radius:12px;padding:12px;}
.pmpr .foot{margin-top:18px;font-size:10px;color:#8899bb;text-align:center;border-top:1px solid #e6ecf8;padding-top:10px;}
.pmpr .sign{display:flex;justify-content:space-between;margin-top:26px;font-size:11px;color:#5a6a90;}
.pmpr .red{color:#ef4444;}
.pmpr .grn{color:#16a34a;}
.pm .stmtWrap{overflow-x:auto;}
.pm .pmtabs{background:#0a1538;display:flex;gap:4px;padding:8px 14px;overflow-x:auto;border-bottom:1px solid rgba(74,158,255,.2);}
.pm .pmtab{white-space:nowrap;border:1px solid transparent;background:none;color:rgba(255,255,255,.55);font-weight:700;font-size:12px;border-radius:18px;padding:7px 15px;cursor:pointer;transition:.2s;}
.pm .pmtab:hover{color:#fff;}
.pm .pmtab.on{background:rgba(74,158,255,.18);border-color:rgba(74,158,255,.4);color:#7ab8ff;}
.pm .modal{z-index:650;}
.pm .pmtoast{position:fixed;top:104px;left:50%;transform:translateX(-50%);background:#16a34a;color:#fff;font-weight:700;font-size:13px;padding:10px 22px;border-radius:30px;z-index:9999;box-shadow:0 6px 24px rgba(0,0,0,.25);max-width:92vw;text-align:center;}
.pm .void-row td{opacity:.55;text-decoration:line-through;}
.pm .void-row td:last-child,.pm .void-row td:nth-child(2){text-decoration:none;opacity:1;}
@media(max-width:900px){.pm .cols{grid-template-columns:1fr;}}
@media(max-width:700px){.pm .wrap{padding:16px 12px 60px;}}
#pm-print-area{display:none;}
@media print{
  @page{size:A4;margin:12mm;}
  body.pm-printing{background:#fff!important;}
  body.pm-printing > *:not(#pm-print-area){display:none!important;}
  body.pm-printing > #pm-print-area{display:block!important;background:#fff!important;}
  body.pm-printing #pm-print-area .pmpr{border:none;max-width:none;padding:0;}
}
`;

// ── أدوات ───────────────────────────────────────────────────────────────────
const pmEsc = s => String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const pmD = s => new Date(s+"T00:00:00");
const pmFmt = d => new Date(d).toLocaleDateString("en-GB");
const pmIso = d => { const x=new Date(d); return x.getFullYear()+"-"+String(x.getMonth()+1).padStart(2,"0")+"-"+String(x.getDate()).padStart(2,"0"); };
const pmMoney = n => Math.round(n).toLocaleString("en-US")+" ﷼";
const pmDays = (a,b) => Math.round((b-a)/86400000);
const pmAddMonths = (d,m) => {
  const x=new Date(d), day=x.getDate(); x.setDate(1); x.setMonth(x.getMonth()+m);
  const last=new Date(x.getFullYear(),x.getMonth()+1,0).getDate(); x.setDate(Math.min(day,last)); return x;
};
const pmSum = (a,f) => a.reduce((s,x)=>s+f(x),0);
const pmToday = () => { const d=new Date(); d.setHours(0,0,0,0); return d; };

// الهوية: 10 أرقام تبدأ بـ 1 (مواطن) أو 2 (مقيم) + فحص التحقق (تنبيه فقط)
const pmIdCheck = v => {
  if(!v) return "";
  if(!/^[12]\d{9}$/.test(v)) return "format";
  let s=0; for(let i=0;i<10;i++){ const d=+v[i]; if(i%2===0){ const z=d*2; s+=z>9?z-9:z; } else s+=d; }
  return s%10===0?"":"checksum";
};
const pmOkId = (label,v) => {
  const r=pmIdCheck(v);
  if(r==="format"){ alert(label+": يجب أن يكون 10 أرقام ويبدأ بـ 1 (مواطن) أو 2 (مقيم)"); return false; }
  if(r==="checksum") return confirm(label+": الرقم لا يجتاز فحص التحقق — تأكد منه. هل تريد الحفظ؟");
  return true;
};

const PM_ST_META = { paid:["b-green","✔ مدفوع"], part:["b-blue","◐ جزئي"], late:["b-red","⚠ متأخر"], soon:["b-amber","مستحق قريباً"], up:["b-gray","قادم"] };
const PM_GCATS = {
  in:["عمولة وساطة","أتعاب تسويق","رسوم إدارة أملاك","تقييم عقاري","استشارات","خدمات إيجار وتوثيق","عربون","تأمين","أخرى"],
  out:["إيجار المكتب","رواتب","كهرباء وإنترنت","قرطاسية ومطبوعات","تسويق وإعلان","صيانة المكتب","رسوم حكومية","رد أمانة","أخرى"],
};
const PM_GNAT = {
  in:[["office","دخل للمكتب"],["trust","أمانة (عربون / تأمين) — لا تدخل في الدخل"]],
  out:[["office","مصروف على المكتب"],["trust","رد أمانة — لا يدخل في المصروفات"]],
};
const pmNatureText = v => v.nature==="trust"
  ? (v.kind==="in"?"أمانة (عربون / تأمين) — لا تُحتسب دخلاً للمكتب":"رد أمانة — لا يُحتسب مصروفاً على المكتب")
  : (v.kind==="in"?"دخل للمكتب":"مصروف على المكتب");

// ترقيم متسلسل آمن (معاملة Firestore): seq → سندات القبض، vseq → سندات الصرف، cno → العقود
async function pmTx(field, build){
  return runTransaction(db, async tx=>{
    const ref=doc(db,"pm_meta","counters");
    const snap=await tx.get(ref);
    const n=((snap.exists()?snap.data()[field]:0)||0)+1;
    tx.set(ref,{[field]:n},{merge:true});
    build(n,tx);
    return n;
  });
}

function pmPrint(html){
  let el=document.getElementById("pm-print-area");
  if(!el){ el=document.createElement("div"); el.id="pm-print-area"; el.className="pm"; document.body.appendChild(el); }
  el.innerHTML=html;
  document.body.classList.add("pm-printing");
  const done=()=>{ document.body.classList.remove("pm-printing"); window.removeEventListener("afterprint",done); };
  window.addEventListener("afterprint",done);
  setTimeout(()=>window.print(),60);
}

// ── النموذج الحسابي (نفس منطق الديمو المعتمد) ─────────────────────────────────
function pmModel(data, today){
  const { props, tenants, maint } = data;
  const payments  = data.payments.filter(p=>!p.void);
  const transfers = data.transfers.filter(t=>!t.void);
  const vouchers  = data.vouchers.filter(v=>!v.void);
  const contracts = data.contracts.map(c=>({...c,payments:payments.filter(p=>p.cid===c.id)}));

  const pname = id => props.find(p=>p.id===id)?.name||"";
  const tenantOf = c => tenants.find(t=>t.id===c.tid)||{name:"—",phone:""};
  const left = c => pmDays(today,pmD(c.end));
  const ageOf = s => s ? Math.floor(pmDays(pmD(s),today)/365.2425) : null;
  const dobTxt = s => s ? `${pmFmt(pmD(s))} (${ageOf(s)} سنة)` : "—";

  const instOf = c => {
    const step=12/c.freq, amt=c.rent/c.freq, s=pmD(c.start);
    return Array.from({length:c.freq},(_,idx)=>{
      const paidAmt=pmSum(c.payments.filter(p=>p.idx===idx),p=>p.amount);
      return {c,cid:c.id,t:tenantOf(c),idx,due:pmAddMonths(s,idx*step),amt,paidAmt,rem:Math.max(0,amt-paidAmt)};
    });
  };
  const allInst = () => contracts.flatMap(instOf);
  const st = i => {
    if(i.rem<=0.5) return "paid";
    if(pmDays(today,i.due)<0) return "late";
    if(i.paidAmt>0) return "part";
    return pmDays(today,i.due)<=30?"soon":"up";
  };
  const mChg = m => m.payer==="owner" ? (m.charge ?? m.cost) : 0;
  const mProfit = m => mChg(m)-m.cost;

  const ownerFin = pid => {
    const p=props.find(x=>x.id===pid);
    const pays=[]; contracts.filter(c=>c.pid===pid).forEach(c=>c.payments.forEach(x=>pays.push({c,x})));
    pays.sort((a,b)=>a.x.date.localeCompare(b.x.date)||a.x.seq-b.x.seq);
    const collected=pmSum(pays,o=>o.x.amount);
    const exps=maint.filter(m=>m.pid===pid&&m.status==="مكتمل"&&m.payer==="owner");
    const expenses=pmSum(exps,m=>mChg(m));
    const fee=collected*(p?.fee||0)/100;
    const net=collected-expenses-fee;
    const trs=transfers.filter(t=>t.pid===pid).sort((a,b)=>a.date.localeCompare(b.date)||a.vseq-b.vseq);
    const transferred=pmSum(trs,t=>t.amount);
    const balance=net-transferred;
    const arrears=pmSum(allInst().filter(i=>i.c.pid===pid&&st(i)==="late"),i=>i.rem);
    return {p,pays,collected,exps,expenses,fee,net,trs,transferred,balance,arrears};
  };
  const occupiedUnits = p => p.units.filter(u=>contracts.some(c=>c.pid===p.id&&c.unit===u&&left(c)>=0));
  const tenantFin = t => {
    const cs=contracts.filter(c=>c.tid===t.id), ins=cs.flatMap(instOf);
    const dueTo=pmSum(ins.filter(i=>pmDays(today,i.due)<=0),i=>i.amt);
    const paid=pmSum(cs.flatMap(c=>c.payments),p=>p.amount);
    const arrears=pmSum(ins.filter(i=>st(i)==="late"),i=>i.rem);
    return {cs,ins,dueTo,paid,arrears};
  };

  // دخل المكتب: رسوم الإدارة (تلقائي) + سندات القبض العامة − سندات الصرف العامة − صيانة على المؤسسة ± ربح الصيانة
  const officeData = (from,to) => {
    const inR = d => (!from||d>=from)&&(!to||d<=to);
    const inc=[], exp=[];
    props.forEach(p=>{
      const col=pmSum(contracts.filter(c=>c.pid===p.id).flatMap(c=>c.payments).filter(x=>inR(x.date)),x=>x.amount);
      if(col>0) inc.push({no:"تلقائي",auto:true,date:"",cat:"رسوم إدارة",desc:`${p.fee}% من تحصيل ${pmMoney(col)} — ${p.name}`,amount:col*p.fee/100});
    });
    vouchers.filter(v=>v.kind==="in"&&v.nature==="office"&&inR(v.date)).forEach(v=>inc.push({no:v.no,date:v.date,cat:v.cat,desc:`${v.desc} (${v.party})`,amount:v.amount}));
    vouchers.filter(v=>v.kind==="out"&&v.nature==="office"&&inR(v.date)).forEach(v=>exp.push({no:v.vno||v.no,date:v.date,cat:v.cat,desc:`${v.desc} (${v.party})`,amount:v.amount}));
    maint.filter(m=>m.status==="مكتمل"&&m.cost>0&&m.payer==="firm"&&inR(m.paidDate||m.date)).forEach(m=>exp.push({no:m.vno,date:m.paidDate||m.date,cat:"صيانة على المؤسسة",desc:`${m.title} — ${pname(m.pid)} • ${m.unit}`,amount:m.cost}));
    const mj=maint.filter(m=>m.status==="مكتمل"&&m.cost>0&&inR(m.paidDate||m.date)).sort((a,b)=>(b.paidDate||b.date).localeCompare(a.paidDate||a.date));
    mj.filter(m=>m.payer==="owner").forEach(m=>{
      const pf=mProfit(m), dt=m.paidDate||m.date;
      if(pf>0) inc.push({no:m.vno,date:dt,cat:"ربح صيانة",desc:`${m.title} — ${pname(m.pid)} • ${m.unit} (تكلفة ${pmMoney(m.cost)} / محمّل ${pmMoney(mChg(m))})`,amount:pf});
      else if(pf<0) exp.push({no:m.vno,date:dt,cat:"خسارة صيانة",desc:`${m.title} — ${pname(m.pid)} • محمّل أقل من التكلفة`,amount:-pf});
    });
    inc.sort((a,b)=>(a.auto?0:1)-(b.auto?0:1)||b.date.localeCompare(a.date));
    exp.sort((a,b)=>b.date.localeCompare(a.date));
    const tin=pmSum(inc,l=>l.amount), tout=pmSum(exp,l=>l.amount);
    const tr=vouchers.filter(v=>v.nature==="trust"&&inR(v.date));
    const trust=pmSum(tr.filter(v=>v.kind==="in"),v=>v.amount)-pmSum(tr.filter(v=>v.kind==="out"),v=>v.amount);
    const mCost=pmSum(mj,m=>m.cost), mChgT=pmSum(mj,mChg);
    const mOwnerProfit=pmSum(mj.filter(m=>m.payer==="owner"),mProfit), mFirm=pmSum(mj.filter(m=>m.payer==="firm"),m=>m.cost);
    return {inc,exp,tin,tout,net:tin-tout,trust,mj,mCost,mChgT,mProfit:mOwnerProfit,mFirm,mNet:mChgT-mCost};
  };

  // قائمة السندات الموحّدة (تشمل الملغاة بعلامة void)
  const voucherList = () => {
    const v=[];
    data.payments.forEach(p=>{
      const c=data.contracts.find(x=>x.id===p.cid); if(!c) return;
      v.push({kind:"in",no:p.no,seq:p.seq,date:p.date,void:!!p.void,party:tenantOf(c).name,
        desc:`قسط ${p.idx+1}/${c.freq} — ${c.unit} • ${pname(c.pid)}`,amount:p.amount,view:{k:"r",cid:c.id,id:p.id},voidRef:["pm_payments",p.id]});
    });
    data.transfers.forEach(t=>{
      const p=props.find(x=>x.id===t.pid);
      v.push({kind:"out",no:t.vno,seq:t.vseq,date:t.date,void:!!t.void,party:p?.owner||"—",
        desc:`تحويل للمالك — ${pname(t.pid)}${t.note?" • "+t.note:""}`,amount:t.amount,view:{k:"t",id:t.id},voidRef:["pm_transfers",t.id]});
    });
    maint.forEach(m=>{ if(m.status==="مكتمل"&&m.cost>0) v.push({kind:"out",no:m.vno,seq:m.vseq,date:m.paidDate||m.date,party:m.provider,
      desc:`${m.title} — ${pname(m.pid)} • ${m.unit}`,amount:m.cost,view:{k:"m",id:m.id}}); });
    data.vouchers.forEach(g=>v.push({kind:g.kind,gen:1,trust:g.nature==="trust",no:g.no||g.vno,seq:g.seq||g.vseq,date:g.date,void:!!g.void,party:g.party,
      desc:`${g.cat} — ${g.desc}`,amount:g.amount,view:{k:"g",id:g.id},voidRef:["pm_vouchers",g.id]}));
    return v.sort((a,b)=>b.date.localeCompare(a.date)||(b.seq||0)-(a.seq||0));
  };

  const propHasData = id => data.contracts.some(c=>c.pid===id)||maint.some(m=>m.pid===id)||data.transfers.some(t=>t.pid===id);

  return {raw:data,today,props,tenants,maint,transfers,vouchers,contracts,payments,
    pname,tenantOf,left,ageOf,dobTxt,instOf,allInst,st,mChg,mProfit,ownerFin,occupiedUnits,tenantFin,officeData,voucherList,propHasData};
}

// ── الأوراق القابلة للطباعة (HTML نصي — كل المدخلات تمر عبر pmEsc) ─────────────
const pmHead = (M,title,right) => `<div class="ph"><img src="${PM_LOGO}" alt=""><div><b>${PM_FIRM}</b><small>${PHONE}</small></div><div class="rt"><b>${title}</b><small>${right||pmFmt(M.today)}</small></div></div>`;
const pmFoot = () => `<div class="foot">${PM_FIRM} • ${PHONE} • مرخصة من الهيئة العامة للعقار</div>`;
const pmVoidBanner = x => x&&x.void ? `<div class="big" style="color:#ef4444;border-color:#ef4444">سند ملغي${x.voidReason?` — ${pmEsc(x.voidReason)}`:""}</div>` : "";
const pmRows = (arr,cols) => arr.length?arr.join(""):`<tr><td colspan="${cols}" style="text-align:center;color:#8899bb">لا يوجد</td></tr>`;

function pmStmtHTML(M,pid){
  const f=M.ownerFin(pid), p=f.p; if(!p) return "";
  return `<div class="pmpr">${pmHead(M,"كشف حساب المالك")}
    <div class="kv"><span>العقار</span><b>${pmEsc(p.name)} — ${pmEsc(p.area)}</b></div>
    <div class="kv"><span>المالك</span><b>${pmEsc(p.owner)}</b></div>
    <div class="kv"><span>رقم الهوية</span><b>${pmEsc(p.ownerId||"—")}</b></div>
    <div class="kv"><span>تاريخ الميلاد</span><b>${M.dobTxt(p.ownerDob)}</b></div>
    <div class="sum">
      <div><small>الإيرادات المحصّلة</small><b>${pmMoney(f.collected)}</b></div>
      <div><small>المصروفات</small><b>${pmMoney(f.expenses)}</b></div>
      <div><small>رسوم الإدارة (${p.fee}%)</small><b>${pmMoney(f.fee)}</b></div>
      <div class="hl"><small>صافي المالك</small><b>${pmMoney(f.net)}</b></div>
    </div>
    <div class="sum" style="grid-template-columns:repeat(3,1fr)">
      <div><small>المحوَّل للمالك</small><b>${pmMoney(f.transferred)}</b></div>
      <div class="hl"><small>المتبقي للمالك</small><b>${pmMoney(f.balance)}</b></div>
      <div><small>متأخرات لم تُحصّل</small><b class="${f.arrears?"red":""}">${pmMoney(f.arrears)}</b></div>
    </div>
    <h4>الدفعات المحصّلة</h4>
    <table><thead><tr><th>التاريخ</th><th>المستأجر / الوحدة</th><th>السند</th><th>المبلغ</th></tr></thead><tbody>
    ${pmRows(f.pays.map(o=>`<tr><td>${pmFmt(pmD(o.x.date))}</td><td>${pmEsc(M.tenantOf(o.c).name)} — ${pmEsc(o.c.unit)}</td><td>${pmEsc(o.x.no)}</td><td>${pmMoney(o.x.amount)}</td></tr>`),4)}</tbody></table>
    <h4>المصروفات على المالك</h4>
    <table><thead><tr><th>التاريخ</th><th>البيان</th><th>الوحدة</th><th>المبلغ</th></tr></thead><tbody>
    ${pmRows(f.exps.map(m=>`<tr><td>${pmFmt(pmD(m.date))}</td><td>${pmEsc(m.title)}</td><td>${pmEsc(m.unit)}</td><td>${pmMoney(M.mChg(m))}</td></tr>`),4)}</tbody></table>
    <h4>التحويلات للمالك</h4>
    <table><thead><tr><th>التاريخ</th><th colspan="2">البيان</th><th>المبلغ</th></tr></thead><tbody>
    ${pmRows(f.trs.map(t=>`<tr><td>${pmFmt(pmD(t.date))}</td><td colspan="2">${t.vno?pmEsc(t.vno)+" — ":""}${pmEsc(t.note||"تحويل")}</td><td>${pmMoney(t.amount)}</td></tr>`),4)}</tbody></table>
    ${pmFoot()}</div>`;
}

function pmTenantStmtHTML(M,tid){
  const t=M.tenants.find(x=>x.id===tid); if(!t) return "";
  const f=M.tenantFin(t);
  return `<div class="pmpr">${pmHead(M,"كشف حساب مستأجر")}
    <div class="kv"><span>المستأجر</span><b>${pmEsc(t.name)}</b></div>
    <div class="kv"><span>الجوال</span><b>${pmEsc(t.phone)}</b></div>
    <div class="kv"><span>رقم الهوية</span><b>${pmEsc(t.idNo||"—")}</b></div>
    <div class="kv"><span>تاريخ الميلاد</span><b>${M.dobTxt(t.dob)}</b></div>
    <div class="sum" style="grid-template-columns:repeat(3,1fr)">
      <div><small>المستحق حتى اليوم</small><b>${pmMoney(f.dueTo)}</b></div>
      <div><small>إجمالي المدفوع</small><b class="grn">${pmMoney(f.paid)}</b></div>
      <div class="hl"><small>المتأخر</small><b>${pmMoney(f.arrears)}</b></div>
    </div>
    ${f.cs.map(c=>`<h4>العقد ${pmEsc(c.cno)} — ${pmEsc(c.unit)} • ${pmEsc(M.pname(c.pid))} (${pmFmt(pmD(c.start))} ← ${pmFmt(pmD(c.end))})</h4>
    <table><thead><tr><th>القسط</th><th>الاستحقاق</th><th>القيمة</th><th>المدفوع</th><th>المتبقي</th><th>الحالة</th></tr></thead><tbody>
    ${M.instOf(c).map(i=>`<tr><td>${i.idx+1}/${c.freq}</td><td>${pmFmt(i.due)}</td><td>${pmMoney(i.amt)}</td><td>${pmMoney(i.paidAmt)}</td><td>${pmMoney(i.rem)}</td><td>${PM_ST_META[M.st(i)][1]}</td></tr>`).join("")}</tbody></table>
    <table><thead><tr><th>السند</th><th>التاريخ</th><th>القسط</th><th>الطريقة</th><th colspan="2">المبلغ</th></tr></thead><tbody>
    ${c.payments.length?c.payments.slice().sort((a,b)=>a.seq-b.seq).map(p=>`<tr><td>${pmEsc(p.no)}</td><td>${pmFmt(pmD(p.date))}</td><td>${p.idx+1}</td><td>${pmEsc(p.method)}</td><td colspan="2">${pmMoney(p.amount)}</td></tr>`).join(""):'<tr><td colspan="6" style="text-align:center;color:#8899bb">لا توجد دفعات</td></tr>'}</tbody></table>`).join("")}
    ${pmFoot()}</div>`;
}

function pmReceiptHTML(M,cid,id){
  const p=M.raw.payments.find(x=>x.id===id), c=M.contracts.find(x=>x.id===cid)||M.raw.contracts.find(x=>x.id===cid);
  if(!p||!c) return null;
  const cc={...c,payments:M.raw.payments.filter(x=>x.cid===cid&&!x.void)};
  const i=M.instOf(cc)[p.idx], t=M.tenantOf(c);
  const before=pmSum(M.raw.payments.filter(x=>x.cid===cid&&!x.void&&x.idx===p.idx&&x.seq<=p.seq),x=>x.amount);
  const remAfter=Math.max(0,i.amt-before);
  return `<div class="pmpr">${pmHead(M,"سند قبض",pmEsc(p.no))}
    ${pmVoidBanner(p)}
    <div class="kv"><span>التاريخ</span><b>${pmFmt(pmD(p.date))}</b></div>
    <div class="kv"><span>استلمنا من</span><b>${pmEsc(t.name)}</b></div>
    <div class="big">${pmMoney(p.amount)}</div>
    <div class="kv"><span>وذلك عن</span><b>القسط ${p.idx+1} من ${c.freq} — عقد ${pmEsc(c.cno)}</b></div>
    <div class="kv"><span>العقار / الوحدة</span><b>${pmEsc(M.pname(c.pid))} — ${pmEsc(c.unit)}</b></div>
    <div class="kv"><span>طريقة الدفع</span><b>${pmEsc(p.method)}</b></div>
    <div class="kv"><span>قيمة القسط</span><b>${pmMoney(i.amt)}</b></div>
    <div class="kv"><span>المسدد حتى هذا السند</span><b>${pmMoney(before)}</b></div>
    <div class="kv"><span>المتبقي من القسط</span><b class="${remAfter>0.5?"red":"grn"}">${remAfter>0.5?pmMoney(remAfter):"لا يوجد — تم السداد"}</b></div>
    ${p.note?`<div class="kv"><span>ملاحظة</span><b>${pmEsc(p.note)}</b></div>`:""}
    <div class="sign"><span>المستلم: ____________</span><span>المستأجر: ____________</span></div>
    ${pmFoot()}</div>`;
}

function pmDisbHTML(M,kind,id){
  let no,date,to,amount,rows,x;
  if(kind==="t"){
    const t=M.raw.transfers.find(y=>y.id===id); if(!t) return null; x=t;
    const p=M.props.find(y=>y.id===t.pid), f=M.ownerFin(t.pid);
    const upTo=pmSum(M.transfers.filter(y=>y.pid===t.pid&&y.vseq<=t.vseq),y=>y.amount);
    const after=f.net-upTo;
    no=t.vno; date=t.date; to=p?.owner||"—"; amount=t.amount;
    rows=`<div class="kv"><span>وذلك عن</span><b>تحويل مستحقات المالك — ${pmEsc(p?.name||"")}</b></div>
      <div class="kv"><span>صافي المالك حتى تاريخه</span><b>${pmMoney(f.net)}</b></div>
      <div class="kv"><span>المحوَّل حتى هذا السند</span><b>${pmMoney(upTo)}</b></div>
      <div class="kv"><span>المتبقي للمالك</span><b class="${after>0.5?"red":"grn"}">${after>0.5?pmMoney(after):"لا يوجد"}</b></div>
      <div class="kv"><span>طريقة الصرف / البيان</span><b>${pmEsc(t.note||"—")}</b></div>`;
  } else {
    const m=M.maint.find(y=>y.id===id); if(!m) return null; x=m;
    const p=M.props.find(y=>y.id===m.pid);
    no=m.vno; date=m.paidDate||m.date; to=m.provider; amount=m.cost;
    rows=`<div class="kv"><span>وذلك عن</span><b>${pmEsc(m.title)}</b></div>
      <div class="kv"><span>العقار / الوحدة</span><b>${pmEsc(p?.name||"")} — ${pmEsc(m.unit)}</b></div>
      <div class="kv"><span>تاريخ الطلب</span><b>${pmFmt(pmD(m.date))}</b></div>
      <div class="kv"><span>على حساب</span><b>${m.payer==="owner"?"المالك ("+pmEsc(p?.owner||"")+")":"المؤسسة"}</b></div>`;
  }
  return `<div class="pmpr">${pmHead(M,"سند صرف",pmEsc(no))}
    ${pmVoidBanner(x)}
    <div class="kv"><span>التاريخ</span><b>${pmFmt(pmD(date))}</b></div>
    <div class="kv"><span>صرفنا إلى</span><b>${pmEsc(to)}</b></div>
    <div class="big">${pmMoney(amount)}</div>
    ${rows}
    <div class="sign"><span>المستلم: ____________</span><span>المحاسب: ____________</span></div>
    ${pmFoot()}</div>`;
}

function pmGeneralHTML(M,id){
  const v=M.raw.vouchers.find(y=>y.id===id); if(!v) return null;
  const isIn=v.kind==="in";
  return `<div class="pmpr">${pmHead(M,isIn?"سند قبض":"سند صرف",pmEsc(v.no||v.vno))}
    ${pmVoidBanner(v)}
    <div class="kv"><span>التاريخ</span><b>${pmFmt(pmD(v.date))}</b></div>
    <div class="kv"><span>${isIn?"استلمنا من":"صرفنا إلى"}</span><b>${pmEsc(v.party)}</b></div>
    <div class="big">${pmMoney(v.amount)}</div>
    <div class="kv"><span>وذلك عن</span><b>${pmEsc(v.desc)}</b></div>
    <div class="kv"><span>التصنيف</span><b>${pmEsc(v.cat)}</b></div>
    <div class="kv"><span>طريقة ${isIn?"الدفع":"الصرف"}</span><b>${pmEsc(v.method)}</b></div>
    <div class="kv"><span>طبيعة المبلغ</span><b>${pmNatureText(v)}</b></div>
    <div class="sign"><span>${isIn?"المستلم":"المحاسب"}: ____________</span><span>${isIn?"الدافع":"المستلم"}: ____________</span></div>
    ${pmFoot()}</div>`;
}

function pmOfficeHTML(M,from,to){
  const d=M.officeData(from,to);
  const per=`من ${from?pmFmt(pmD(from)):"البداية"} إلى ${to?pmFmt(pmD(to)):"اليوم"}`;
  const tbl=(rows,total,label)=>`<table><thead><tr><th>السند</th><th>التاريخ</th><th>عن</th><th>المبلغ</th></tr></thead><tbody>
    ${rows.length?rows.map(l=>`<tr><td>${pmEsc(l.no)}</td><td>${l.date?pmFmt(pmD(l.date)):"—"}</td><td><b>${pmEsc(l.cat)}</b> — ${pmEsc(l.desc)}</td><td>${pmMoney(l.amount)}</td></tr>`).join(""):'<tr><td colspan="4" style="text-align:center;color:#8899bb">لا يوجد</td></tr>'}
    <tr style="font-weight:900"><td colspan="3">${label}</td><td>${pmMoney(total)}</td></tr></tbody></table>`;
  return `<div class="pmpr">${pmHead(M,"كشف دخل المكتب",per)}
    <div class="sum" style="grid-template-columns:repeat(3,1fr)">
      <div><small>إجمالي الدخل</small><b class="grn">${pmMoney(d.tin)}</b></div>
      <div><small>إجمالي المصروفات</small><b class="red">${pmMoney(d.tout)}</b></div>
      <div class="hl"><small>صافي دخل المكتب</small><b>${pmMoney(d.net)}</b></div></div>
    <h4>الدخل</h4>${tbl(d.inc,d.tin,"إجمالي الدخل")}
    <h4>المصروفات</h4>${tbl(d.exp,d.tout,"إجمالي المصروفات")}
    <h4>الصيانة — التكلفة والربح وصافي الربح</h4>
    <table><thead><tr><th>السند</th><th>التاريخ</th><th>العمل</th><th>التكلفة</th><th>المحمّل</th><th>الربح / الخسارة</th></tr></thead><tbody>
    ${d.mj.length?d.mj.map(m=>{const pf=M.mProfit(m);return `<tr><td>${pmEsc(m.vno||"—")}</td><td>${pmFmt(pmD(m.paidDate||m.date))}</td><td><b>${pmEsc(m.title)}</b> — ${pmEsc(M.pname(m.pid))} • ${pmEsc(m.unit)} (على ${m.payer==="owner"?"المالك":"المؤسسة"})</td><td>${pmMoney(m.cost)}</td><td>${pmMoney(M.mChg(m))}</td><td class="${pf>=0?"grn":"red"}">${pf<0?"−":""}${pmMoney(Math.abs(pf))}</td></tr>`;}).join(""):'<tr><td colspan="6" style="text-align:center;color:#8899bb">لا يوجد</td></tr>'}
    <tr style="font-weight:900"><td colspan="3">الإجمالي (صافي ربح الصيانة بعد خصم صيانة على المؤسسة)</td><td>${pmMoney(d.mCost)}</td><td>${pmMoney(d.mChgT)}</td><td class="${d.mNet>=0?"grn":"red"}">${d.mNet<0?"−":""}${pmMoney(Math.abs(d.mNet))}</td></tr></tbody></table>
    <div class="kv" style="margin-top:10px"><span>أمانات محتجزة (غير محتسبة في الدخل)</span><b>${pmMoney(d.trust)}</b></div>
    ${pmFoot()}</div>`;
}

function pmViewHTML(M,v){
  if(!v) return null;
  if(v.k==="r") return pmReceiptHTML(M,v.cid,v.id);
  if(v.k==="t") return pmDisbHTML(M,"t",v.id);
  if(v.k==="m") return pmDisbHTML(M,"m",v.id);
  if(v.k==="g") return pmGeneralHTML(M,v.id);
  return null;
}

// ── مكوّنات النوافذ ─────────────────────────────────────────────────────────
const pmUseForm = init => {
  const [f,setF]=useState(init);
  const set=k=>e=>{ const v=e.target.value; setF(s=>({...s,[k]:v})); };
  return [f,set,setF];
};
const PMDbErr = e => { console.error(e); alert("تعذّر الحفظ — تحقق من الاتصال ثم أعد المحاولة"); };
const pmNow = () => new Date().toISOString();

function PMModal({onClose,wide,children}){
  return (
    <div className="modal on" onMouseDown={e=>{ if(e.target===e.currentTarget) onClose(); }}>
      <div className="mbox" style={wide?{maxWidth:840}:undefined}>{children}</div>
    </div>
  );
}
function PMBtns({onSave,saveText,onClose,busy,extra}){
  return (
    <div style={{display:"flex",gap:10}}>
      <button className="btn blue" style={{flex:1}} disabled={busy} onClick={onSave}>{busy?"جاري الحفظ...":saveText}</button>
      {extra}
      <button className="btn ghost" onClick={onClose}>إلغاء</button>
    </div>
  );
}

function PMContractModal({M,pre,onClose,toast}){
  const p0=pre||{}, today=M.today;
  const [f,set,setF]=pmUseForm({
    pid:p0.pid||M.props[0]?.id||"", unit:p0.unit||"", ejar:"", ten:p0.ten||"", phone:p0.phone||"",
    tid:p0.tid||"", tdob:p0.tdob||"", start:p0.start||pmIso(today), end:p0.end||pmIso(pmAddMonths(today,12)),
    rent:p0.rent||"", freq:String(p0.freq||4),
  });
  const [busy,setBusy]=useState(false);
  const prop=M.props.find(x=>x.id===f.pid);
  const fillTen=e=>{
    const v=e.target.value;
    setF(s=>{ const t=M.tenants.find(x=>x.name===v.trim()); return t?{...s,ten:v,phone:t.phone||"",tid:t.idNo||"",tdob:t.dob||""}:{...s,ten:v}; });
  };
  const save=async()=>{
    if(busy) return;
    const unit=f.unit.trim(), name=f.ten.trim(), rent=+f.rent, tid=f.tid.trim(), tdob=f.tdob, phone=f.phone.trim();
    if(!f.pid) return alert("أضف عقاراً أولاً");
    if(!unit||!name||!(rent>0)||!f.start||!f.end) return alert("الوحدة والمستأجر والتواريخ والإيجار مطلوبة");
    if(f.end<=f.start) return alert("تاريخ نهاية العقد يجب أن يكون بعد تاريخ البداية");
    if(tdob&&tdob>pmIso(today)) return alert("تاريخ الميلاد لا يمكن أن يكون في المستقبل");
    const other=tid&&M.tenants.find(x=>x.idNo===tid&&x.name!==name);
    if(other) return alert("هذه الهوية مسجلة للمستأجر «"+other.name+"» — اكتب نفس الاسم لربط العقد به");
    if(!pmOkId("هوية المستأجر",tid)) return;
    const overlap=M.raw.contracts.find(c=>c.pid===f.pid&&c.unit===unit&&c.start<=f.end&&c.end>=f.start);
    if(overlap&&!confirm("الوحدة عليها عقد ("+overlap.cno+") في نفس الفترة. هل تريد المتابعة؟")) return;
    setBusy(true);
    try{
      const t=M.tenants.find(x=>x.name===name); let tenantId;
      if(!t){ const r=await addDoc(collection(db,"pm_tenants"),{name,phone,idNo:tid,dob:tdob,createdAt:pmNow()}); tenantId=r.id; }
      else {
        tenantId=t.id; const upd={}; if(phone) upd.phone=phone; if(tid) upd.idNo=tid; if(tdob) upd.dob=tdob;
        if(Object.keys(upd).length) await updateDoc(doc(db,"pm_tenants",t.id),upd);
      }
      const ref=doc(collection(db,"pm_contracts"));
      await pmTx("cno",(n,tx)=>tx.set(ref,{cno:"C-"+(100+n),pid:f.pid,unit,tid:tenantId,ejar:f.ejar.trim(),start:f.start,end:f.end,rent,freq:+f.freq,createdAt:pmNow()}));
      toast("✔ تم حفظ العقد وإنشاء جدول الأقساط"); onClose();
    }catch(e){ PMDbErr(e); setBusy(false); }
  };
  return (
    <PMModal onClose={onClose}>
      <h3>{pre?.renew?"🔄 تجديد عقد":"📄 إضافة عقد جديد"}</h3>
      {!M.props.length?<div className="info">أضف عقاراً من تبويب «العقارات» قبل إنشاء العقود.</div>:(
        <>
          <div className="fg">
            <div className="full"><label>العقار</label>
              <select value={f.pid} onChange={set("pid")}>{M.props.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></div>
            <div><label>الوحدة</label><input value={f.unit} onChange={set("unit")} list="pm-units" placeholder="اختر أو اكتب"/>
              <datalist id="pm-units">{(prop?.units||[]).map(u=><option key={u} value={u}/>)}</datalist></div>
            <div><label>رقم عقد إيجار</label><input value={f.ejar} onChange={set("ejar")} placeholder="اختياري"/></div>
            <div><label>المستأجر</label><input value={f.ten} onChange={fillTen} list="pm-tens" placeholder="اسم المستأجر"/>
              <datalist id="pm-tens">{M.tenants.map(t=><option key={t.id} value={t.name}/>)}</datalist></div>
            <div><label>الجوال</label><input value={f.phone} onChange={set("phone")} placeholder="05xxxxxxxx"/></div>
            <div><label>هوية المستأجر</label><input value={f.tid} onChange={set("tid")} inputMode="numeric" maxLength={10} placeholder="10 أرقام"/></div>
            <div><label>تاريخ ميلاد المستأجر</label><input type="date" value={f.tdob} onChange={set("tdob")}/></div>
            <div><label>بداية العقد</label><input type="date" value={f.start} onChange={set("start")}/></div>
            <div><label>نهاية العقد</label><input type="date" value={f.end} onChange={set("end")}/></div>
            <div><label>الإيجار السنوي (﷼)</label><input type="number" value={f.rent} onChange={set("rent")}/></div>
            <div><label>عدد الدفعات</label>
              <select value={f.freq} onChange={set("freq")}><option value="1">دفعة واحدة</option><option value="2">دفعتين</option><option value="4">4 دفعات</option><option value="12">شهري</option></select></div>
          </div>
          <PMBtns onSave={save} saveText="✅ حفظ العقد" onClose={onClose} busy={busy}/>
        </>
      )}
    </PMModal>
  );
}

function PMPayModal({M,cid,idx,onClose,onSaved}){
  const c=M.contracts.find(x=>x.id===cid), inst=c?M.instOf(c)[idx]:null;
  const [f,set,setF]=pmUseForm({amt:inst?String(Math.round(inst.rem)):"",date:pmIso(M.today),method:"نقد",note:""});
  const [busy,setBusy]=useState(false);
  if(!inst) return null;
  const a=+f.amt||0, r=inst.rem-a;
  const save=async()=>{
    if(busy) return;
    if(!(a>0)) return alert("أدخل مبلغاً صحيحاً");
    if(a>inst.rem+0.5) return alert("المبلغ أكبر من المتبقي على القسط");
    setBusy(true);
    try{
      const ref=doc(collection(db,"pm_payments"));
      await pmTx("seq",(n,tx)=>tx.set(ref,{cid,idx,amount:a,date:f.date||pmIso(M.today),method:f.method,note:f.note.trim(),seq:n,no:"R-"+(1000+n),createdAt:pmNow()}));
      onSaved(cid,ref.id);
    }catch(e){ PMDbErr(e); setBusy(false); }
  };
  return (
    <PMModal onClose={onClose}>
      <h3>💰 تسجيل دفعة</h3>
      <div className="info">
        <b>{inst.t.name}</b> — {c.unit} • {M.pname(c.pid)}<br/>
        قسط {idx+1} من {c.freq} • استحقاق {pmFmt(inst.due)}<br/>
        قيمة القسط {pmMoney(inst.amt)} • المدفوع سابقاً {pmMoney(inst.paidAmt)} • <b style={{color:"var(--red)"}}>المتبقي {pmMoney(inst.rem)}</b>
      </div>
      <div className="fg">
        <div className="full"><label>المبلغ المستلم (﷼)</label><input type="number" min="0" step="any" value={f.amt} onChange={set("amt")}/></div>
        <div className="full"><button type="button" className="btn ghost" style={{width:"100%"}} onClick={()=>setF(s=>({...s,amt:String(Math.round(inst.rem))}))}>تعبئة المتبقي كاملاً</button></div>
        <div><label>تاريخ الاستلام</label><input type="date" value={f.date} onChange={set("date")}/></div>
        <div><label>طريقة الدفع</label><select value={f.method} onChange={set("method")}><option>نقد</option><option>تحويل بنكي</option><option>شيك</option></select></div>
        <div className="full"><label>ملاحظة (اختياري)</label><input value={f.note} onChange={set("note")} placeholder="مثال: وعد بسداد الباقي نهاية الشهر"/></div>
      </div>
      <div className="pnote">
        {a<=0?"أدخل المبلغ المستلم":r>0.5?<>⚠ دفعة جزئية — يبقى على القسط <b>{pmMoney(r)}</b> مفتوحاً</>:r<-0.5?<span style={{color:"var(--red)"}}>المبلغ أكبر من المتبقي بـ {pmMoney(-r)}</span>:"✔ سداد كامل للقسط"}
      </div>
      <PMBtns onSave={save} saveText="✅ حفظ الدفعة" onClose={onClose} busy={busy}/>
    </PMModal>
  );
}

function PMHistModal({M,cid,idx,onClose,onPay,onView,onVoid}){
  const c=M.contracts.find(x=>x.id===cid); if(!c) return null;
  const i=M.instOf(c)[idx];
  const list=M.raw.payments.filter(p=>p.cid===cid&&p.idx===idx).sort((a,b)=>a.seq-b.seq);
  return (
    <PMModal onClose={onClose}>
      <h3>🧾 سجل دفعات القسط</h3>
      <div className="info"><b>{i.t.name}</b> — {c.unit} • قسط {idx+1}/{c.freq}<br/>القيمة {pmMoney(i.amt)} • المدفوع {pmMoney(i.paidAmt)} • المتبقي <b>{pmMoney(i.rem)}</b></div>
      {list.length?list.map(p=>(
        <div className="hrow" key={p.id} style={p.void?{opacity:.55}:undefined}>
          <div className="hm"><b style={p.void?{textDecoration:"line-through"}:undefined}>{pmMoney(p.amount)}</b> <span className="badge b-gray">{p.no}</span>{p.void&&<span className="badge b-red" style={{marginRight:4}}>ملغي</span>}
            <small>{pmFmt(pmD(p.date))} • {p.method}{p.note?" • "+p.note:""}</small></div>
          <button className="btn ghost sm" onClick={()=>onView({k:"r",cid,id:p.id})}>🧾 سند القبض</button>
          {!p.void&&<button className="btn ghost sm" style={{color:"var(--red)"}} onClick={()=>onVoid("pm_payments",p.id,p.no)}>إلغاء</button>}
        </div>
      )):<div className="empty">لا توجد دفعات</div>}
      <div style={{display:"flex",gap:10,marginTop:14}}>
        {i.rem>0.5&&<button className="btn blue" style={{flex:1}} onClick={()=>onPay(cid,idx)}>+ دفعة أخرى</button>}
        <button className="btn ghost" onClick={onClose}>إغلاق</button>
      </div>
    </PMModal>
  );
}

function PMMaintModal({M,providers,onClose,toast}){
  const [f,set]=pmUseForm({pid:M.props[0]?.id||"",unit:"",date:pmIso(M.today),title:"",prov:""});
  const [busy,setBusy]=useState(false);
  const prop=M.props.find(x=>x.id===f.pid);
  const units=prop?[...prop.units,"المبنى"]:[];
  const save=async()=>{
    if(busy) return;
    if(!f.pid) return alert("أضف عقاراً أولاً");
    if(!f.title.trim()) return alert("اكتب وصف العطل");
    setBusy(true);
    try{
      await addDoc(collection(db,"pm_maint"),{pid:f.pid,unit:f.unit||units[0]||"المبنى",title:f.title.trim(),provider:f.prov.trim()||"—",status:"جديد",cost:0,payer:"owner",date:f.date||pmIso(M.today),createdAt:pmNow()});
      toast("✔ تم إنشاء طلب الصيانة"); onClose();
    }catch(e){ PMDbErr(e); setBusy(false); }
  };
  return (
    <PMModal onClose={onClose}>
      <h3>🔧 طلب صيانة جديد</h3>
      <div className="fg">
        <div className="full"><label>العقار</label><select value={f.pid} onChange={set("pid")}>{M.props.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></div>
        <div><label>الوحدة</label><select value={f.unit||units[0]||""} onChange={set("unit")}>{units.map(u=><option key={u}>{u}</option>)}</select></div>
        <div><label>التاريخ</label><input type="date" value={f.date} onChange={set("date")}/></div>
        <div className="full"><label>وصف العطل</label><input value={f.title} onChange={set("title")} placeholder="مثال: تسرب مياه"/></div>
        <div className="full"><label>المزود (اختر من صفحة مزودي الخدمات أو اكتب اسماً)</label>
          <input value={f.prov} onChange={set("prov")} list="pm-provs" placeholder="اسم المزود"/>
          <datalist id="pm-provs">{providers.map(p=><option key={p} value={p}/>)}</datalist></div>
      </div>
      <PMBtns onSave={save} saveText="✅ حفظ الطلب" onClose={onClose} busy={busy}/>
    </PMModal>
  );
}

function PMCostModal({M,id,onClose,toast,onView}){
  const m=M.maint.find(x=>x.id===id);
  const [f,set,setF]=pmUseForm({cost:"",payer:"owner",charge:""});
  const [touched,setTouched]=useState(false);
  const [busy,setBusy]=useState(false);
  if(!m) return null;
  const firm=f.payer==="firm", c=+f.cost||0;
  const chg=firm?0:(f.charge===""?c:+f.charge), pf=chg-c;
  const onCost=e=>{ const v=e.target.value; setF(s=>({...s,cost:v,charge:(!touched&&s.payer==="owner")?v:s.charge})); };
  const onPayer=e=>{ const v=e.target.value; setF(s=>({...s,payer:v,charge:(v==="owner"&&!touched)?s.cost:s.charge})); };
  const save=async()=>{
    if(busy) return;
    if(f.cost===""||!(c>=0)) return alert("أدخل التكلفة");
    setBusy(true);
    try{
      const upd={cost:c,payer:f.payer,charge:f.payer==="owner"?(f.charge===""?c:+f.charge):0,status:"مكتمل",paidDate:pmIso(M.today)};
      if(c>0){
        await pmTx("vseq",(n,tx)=>tx.update(doc(db,"pm_maint",id),{...upd,vseq:n,vno:"S-"+(2000+n)}));
      } else await updateDoc(doc(db,"pm_maint",id),upd);
      toast("✔ تم إتمام الطلب"+(c>0?" وإصدار سند الصرف":"")+(f.payer==="owner"&&c>0?" وخصم "+pmMoney(upd.charge)+" من كشف المالك":""));
      onClose(); if(c>0) onView({k:"m",id});
    }catch(e){ PMDbErr(e); setBusy(false); }
  };
  return (
    <PMModal onClose={onClose}>
      <h3>✅ إتمام طلب الصيانة</h3>
      <div className="info"><b>{m.title}</b><br/>{M.pname(m.pid)} • {m.unit} • {m.provider}</div>
      <div className="fg">
        <div><label>التكلفة — المدفوع للمزود (﷼)</label><input type="number" min="0" value={f.cost} onChange={onCost}/></div>
        <div><label>على حساب</label><select value={f.payer} onChange={onPayer}><option value="owner">المالك</option><option value="firm">المؤسسة</option></select></div>
        {!firm&&<div className="full"><label>المبلغ المحمّل على المالك (﷼)</label><input type="number" min="0" value={f.charge} onChange={e=>{setTouched(true);set("charge")(e);}}/></div>}
      </div>
      <div className="pnote">
        {firm?`على المؤسسة: تُسجَّل ${pmMoney(c)} مصروفاً (خسارة صيانة) في دخل المكتب`
          :<>الربح = {pmMoney(chg)} − {pmMoney(c)} = <b style={{color:pf>=0?"var(--green)":"var(--red)"}}>{pmMoney(pf)}</b> {pf<0?"(خسارة على المكتب)":"(يُضاف لدخل المكتب)"}</>}
      </div>
      <PMBtns onSave={save} saveText="✅ تأكيد الإتمام" onClose={onClose} busy={busy}/>
    </PMModal>
  );
}

function PMTransModal({M,pid0,onClose,toast,onView}){
  const first=pid0||M.props[0]?.id||"";
  const [f,set,setF]=pmUseForm({pid:first,amt:"",date:pmIso(M.today),note:"تحويل بنكي"});
  const [busy,setBusy]=useState(false);
  const [touched,setTouched]=useState(false);
  const fin=f.pid?M.ownerFin(f.pid):null;
  const amt=touched?f.amt:String(fin?Math.max(0,Math.round(fin.balance)):"");
  const save=async()=>{
    if(busy||!fin||!fin.p) return;
    const a=+amt; if(!(a>0)) return alert("أدخل مبلغاً صحيحاً");
    if(a>fin.balance+0.5&&!confirm("المبلغ أكبر من المتبقي للمالك. هل تريد المتابعة؟")) return;
    setBusy(true);
    try{
      const ref=doc(collection(db,"pm_transfers"));
      await pmTx("vseq",(n,tx)=>tx.set(ref,{pid:f.pid,amount:a,date:f.date||pmIso(M.today),note:f.note.trim(),vseq:n,vno:"S-"+(2000+n),createdAt:pmNow()}));
      toast("✔ تم تسجيل التحويل وإصدار سند الصرف"); onClose(); onView({k:"t",id:ref.id});
    }catch(e){ PMDbErr(e); setBusy(false); }
  };
  return (
    <PMModal onClose={onClose}>
      <h3>💸 تسجيل تحويل للمالك</h3>
      {!M.props.length?<div className="info">أضف عقاراً أولاً.</div>:(
        <>
          {fin&&fin.p&&<div className="info">المالك: <b>{fin.p.owner}</b><br/>صافي المالك {pmMoney(fin.net)} • حُوِّل سابقاً {pmMoney(fin.transferred)} • <b style={{color:"var(--green)"}}>المتبقي {pmMoney(fin.balance)}</b></div>}
          <div className="fg">
            <div className="full"><label>العقار</label><select value={f.pid} onChange={e=>{set("pid")(e);setTouched(false);}}>{M.props.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></div>
            <div><label>المبلغ (﷼)</label><input type="number" min="0" value={amt} onChange={e=>{setTouched(true);set("amt")(e);}}/></div>
            <div><label>التاريخ</label><input type="date" value={f.date} onChange={set("date")}/></div>
            <div className="full"><label>البيان</label><input value={f.note} onChange={set("note")} placeholder="مثال: تحويل بنكي"/></div>
          </div>
          <PMBtns onSave={save} saveText="✅ حفظ التحويل" onClose={onClose} busy={busy}/>
        </>
      )}
    </PMModal>
  );
}

function PMTenantModal({M,id,onClose,toast}){
  const t=M.tenants.find(x=>x.id===id);
  const [f,set]=pmUseForm({name:t?.name||"",phone:t?.phone||"",idNo:t?.idNo||"",dob:t?.dob||""});
  const [busy,setBusy]=useState(false);
  if(!t) return null;
  const save=async()=>{
    if(busy) return;
    const name=f.name.trim(), idNo=f.idNo.trim();
    if(!name) return alert("الاسم مطلوب");
    if(f.dob&&f.dob>pmIso(M.today)) return alert("تاريخ الميلاد لا يمكن أن يكون في المستقبل");
    const other=idNo&&M.tenants.find(x=>x.idNo===idNo&&x.id!==id);
    if(other) return alert("هذه الهوية مسجلة للمستأجر «"+other.name+"»");
    const dupName=M.tenants.find(x=>x.name===name&&x.id!==id);
    if(dupName) return alert("يوجد مستأجر آخر بنفس الاسم");
    if(!pmOkId("هوية المستأجر",idNo)) return;
    setBusy(true);
    try{ await updateDoc(doc(db,"pm_tenants",id),{name,phone:f.phone.trim(),idNo,dob:f.dob}); toast("✔ تم تحديث بيانات المستأجر"); onClose(); }
    catch(e){ PMDbErr(e); setBusy(false); }
  };
  return (
    <PMModal onClose={onClose}>
      <h3>✏️ بيانات المستأجر</h3>
      <div className="fg">
        <div className="full"><label>الاسم</label><input value={f.name} onChange={set("name")}/></div>
        <div><label>الجوال</label><input value={f.phone} onChange={set("phone")} placeholder="05xxxxxxxx"/></div>
        <div><label>رقم الهوية / الإقامة</label><input value={f.idNo} onChange={set("idNo")} inputMode="numeric" maxLength={10} placeholder="10 أرقام"/></div>
        <div className="full"><label>تاريخ الميلاد</label><input type="date" value={f.dob} onChange={set("dob")}/></div>
      </div>
      <PMBtns onSave={save} saveText="✅ حفظ" onClose={onClose} busy={busy}/>
    </PMModal>
  );
}

const pmParseUnits = s => [...new Set(s.split(/[,،\n]+/).map(x=>x.trim()).filter(Boolean))];
function PMPropModal({M,id,onClose,toast}){
  const p=id?M.props.find(x=>x.id===id):null;
  const [f,set,setF]=pmUseForm({name:p?.name||"",area:p?.area||"",owner:p?.owner||"",ownerPhone:p?.ownerPhone||"",ownerId:p?.ownerId||"",ownerDob:p?.ownerDob||"",
    fee:String(p?p.fee:5),icon:p?.icon||"🏢",units:p?p.units.join("، "):"",pre:"شقة",cnt:"6"});
  const [busy,setBusy]=useState(false);
  const gen=()=>{ const pre=f.pre.trim()||"وحدة", n=Math.min(200,Math.max(1,+f.cnt||1)); setF(s=>({...s,units:Array.from({length:n},(_,i)=>pre+" "+(i+1)).join("، ")})); };
  const save=async()=>{
    if(busy) return;
    const name=f.name.trim(), units=pmParseUnits(f.units), fee=+f.fee, ownerId=f.ownerId.trim();
    if(!name||!units.length) return alert("اسم العقار والوحدات مطلوبة");
    if(!(fee>=0)||fee>100) return alert("نسبة رسوم الإدارة غير صحيحة");
    if(f.ownerDob&&f.ownerDob>pmIso(M.today)) return alert("تاريخ الميلاد لا يمكن أن يكون في المستقبل");
    if(!pmOkId("هوية المالك",ownerId)) return;
    if(p){
      const used=[...new Set(M.raw.contracts.filter(c=>c.pid===id).map(c=>c.unit).filter(u=>!units.includes(u)))];
      if(used.length) return alert("لا يمكن إزالة وحدات عليها عقود: "+used.join("، "));
    }
    const data={name,area:f.area.trim(),owner:f.owner.trim()||"—",ownerPhone:f.ownerPhone.trim(),ownerId,ownerDob:f.ownerDob,fee,icon:f.icon,units};
    setBusy(true);
    try{
      if(p) await updateDoc(doc(db,"pm_props",id),data); else await addDoc(collection(db,"pm_props"),{...data,createdAt:pmNow()});
      toast(p?"✔ تم تحديث العقار":"✔ تمت إضافة العقار — أضف له عقداً من تبويب العقود"); onClose();
    }catch(e){ PMDbErr(e); setBusy(false); }
  };
  const del=async()=>{
    if(!p||M.propHasData(id)) return alert("لا يمكن حذف عقار عليه عقود أو حركات");
    if(!confirm("حذف هذا العقار؟")) return;
    try{ await deleteDoc(doc(db,"pm_props",id)); logAudit("حذف عقار (إدارة أملاك)",id,p?.name||""); toast("✔ تم حذف العقار"); onClose(); }catch(e){ PMDbErr(e); }
  };
  return (
    <PMModal onClose={onClose}>
      <h3>{p?"✏️ تعديل العقار":"🏢 إضافة عقار جديد"}</h3>
      <div className="fg">
        <div className="full"><label>اسم العقار</label><input value={f.name} onChange={set("name")} placeholder="مثال: عمارة النخيل"/></div>
        <div className="full"><label>الموقع</label><input value={f.area} onChange={set("area")} placeholder="مثال: الخبر - العليا"/></div>
        <div><label>اسم المالك</label><input value={f.owner} onChange={set("owner")} placeholder="اسم المالك"/></div>
        <div><label>جوال المالك</label><input value={f.ownerPhone} onChange={set("ownerPhone")} placeholder="05xxxxxxxx"/></div>
        <div><label>هوية المالك</label><input value={f.ownerId} onChange={set("ownerId")} inputMode="numeric" maxLength={10} placeholder="10 أرقام"/></div>
        <div><label>تاريخ ميلاد المالك</label><input type="date" value={f.ownerDob} onChange={set("ownerDob")}/></div>
        <div><label>رسوم الإدارة (%)</label><input type="number" min="0" max="100" step="0.5" value={f.fee} onChange={set("fee")}/></div>
        <div><label>نوع العقار</label>
          <select value={f.icon} onChange={set("icon")}><option value="🏢">🏢 عمارة</option><option value="🏘️">🏘️ مجمع سكني</option><option value="🏬">🏬 محلات / مبنى تجاري</option><option value="🏡">🏡 فيلا</option><option value="🏗️">🏗️ أرض / أخرى</option></select></div>
        <div className="full"><label>الوحدات (افصل بينها بفاصلة أو سطر جديد)</label><textarea rows={3} value={f.units} onChange={set("units")} placeholder="شقة 1، شقة 2، شقة 3"/></div>
        <div><label>توليد سريع — البادئة</label><input value={f.pre} onChange={set("pre")}/></div>
        <div><label>العدد</label><input type="number" min="1" max="200" value={f.cnt} onChange={set("cnt")}/></div>
        <div className="full"><button type="button" className="btn ghost" style={{width:"100%"}} onClick={gen}>⚡ توليد الوحدات تلقائياً</button></div>
      </div>
      <PMBtns onSave={save} saveText="✅ حفظ العقار" onClose={onClose} busy={busy}
        extra={p&&!M.propHasData(id)?<button className="btn ghost" style={{color:"var(--red)"}} onClick={del}>🗑️ حذف</button>:null}/>
    </PMModal>
  );
}

function PMGenModal({M,kind,onClose,toast,onView}){
  const isIn=kind==="in";
  const [f,set,setF]=pmUseForm({date:pmIso(M.today),amt:"",party:"",cat:PM_GCATS[kind][0],method:"نقد",desc:"",nature:"office"});
  const [busy,setBusy]=useState(false);
  const onCat=e=>{ const v=e.target.value; setF(s=>({...s,cat:v,nature:["عربون","تأمين","رد أمانة"].includes(v)?"trust":"office"})); };
  const save=async()=>{
    if(busy) return;
    const amt=+f.amt, party=f.party.trim(), desc=f.desc.trim();
    if(!(amt>0)||!party||!desc) return alert('المبلغ والاسم وبيان "وذلك عن" مطلوبة');
    setBusy(true);
    try{
      const ref=doc(collection(db,"pm_vouchers"));
      const base={kind,date:f.date||pmIso(M.today),party,amount:amt,cat:f.cat,desc,method:f.method,nature:f.nature,createdAt:pmNow()};
      const n=await pmTx(isIn?"seq":"vseq",(n,tx)=>tx.set(ref,isIn?{...base,seq:n,no:"R-"+(1000+n)}:{...base,vseq:n,vno:"S-"+(2000+n)}));
      toast("✔ تم إصدار "+(isIn?"سند القبض ":"سند الصرف ")+(isIn?"R-"+(1000+n):"S-"+(2000+n))); onClose(); onView({k:"g",id:ref.id});
    }catch(e){ PMDbErr(e); setBusy(false); }
  };
  return (
    <PMModal onClose={onClose}>
      <h3>{isIn?"📥 سند قبض عام":"📤 سند صرف عام"}</h3>
      <div className="fg">
        <div><label>التاريخ</label><input type="date" value={f.date} onChange={set("date")}/></div>
        <div><label>المبلغ (﷼)</label><input type="number" min="0" value={f.amt} onChange={set("amt")}/></div>
        <div className="full"><label>{isIn?"استلمنا من":"صرفنا إلى"}</label><input value={f.party} onChange={set("party")} placeholder="الاسم"/></div>
        <div><label>التصنيف</label><select value={f.cat} onChange={onCat}>{PM_GCATS[kind].map(c=><option key={c}>{c}</option>)}</select></div>
        <div><label>{isIn?"طريقة الدفع":"طريقة الصرف"}</label><select value={f.method} onChange={set("method")}><option>نقد</option><option>تحويل بنكي</option><option>شيك</option></select></div>
        <div className="full"><label>وذلك عن</label><input value={f.desc} onChange={set("desc")} placeholder="مثال: عمولة تأجير شقة في العليا"/></div>
        <div className="full"><label>طبيعة المبلغ</label><select value={f.nature} onChange={set("nature")}>{PM_GNAT[kind].map(n=><option key={n[0]} value={n[0]}>{n[1]}</option>)}</select></div>
      </div>
      <div className="pnote">{f.nature==="trust"?"ℹ "+pmNatureText({kind,nature:f.nature}):"✔ سيُحتسب ضمن "+(isIn?"دخل المكتب":"مصروفات المكتب")}</div>
      <PMBtns onSave={save} saveText="✅ حفظ وإصدار السند" onClose={onClose} busy={busy}/>
    </PMModal>
  );
}

function PMVoucherModal({html,onClose}){
  return (
    <PMModal onClose={onClose} wide>
      <div className="stmtWrap">{html?<div dangerouslySetInnerHTML={{__html:html}}/>:<div className="empty">جاري تحميل السند...</div>}</div>
      <div style={{display:"flex",gap:10,marginTop:14}}>
        <button className="btn blue" style={{flex:1}} disabled={!html} onClick={()=>pmPrint(html)}>🖨️ طباعة السند</button>
        <button className="btn ghost" onClick={onClose}>إغلاق</button>
      </div>
      <div className="pnote" style={{marginTop:12}}>💡 زر الطباعة يفتح نافذة طباعة المتصفح: اختر منها «حفظ كـ PDF» إذا تبغى تعاين الشكل قبل الطباعة على الورق.</div>
    </PMModal>
  );
}

// ── الصفحة الرئيسية لإدارة الأملاك ─────────────────────────────────────────────
const PM_TABS = [["dash","📊 لوحة التحكم"],["props","🏢 العقارات"],["tenants","👥 المستأجرين"],["contracts","📄 العقود"],["collect","💰 التحصيل"],["vouchers","📑 السندات"],["office","💼 دخل المكتب"],["maint","🔧 الصيانة"],["owners","🧾 كشوف الملاك"]];
const PM_COLLS = ["props","tenants","contracts","payments","maint","transfers","vouchers"];

function PropertyMgmtPage({ lang, T, darkMode }) {
  const [data,setData]       = useState({props:[],tenants:[],contracts:[],payments:[],maint:[],transfers:[],vouchers:[]});
  const [ready,setReady]     = useState({});
  const [loadErr,setLoadErr] = useState(false);
  const [providers,setProviders] = useState([]);
  const [tab,setTab]         = useState("dash");
  const [mod,setMod]         = useState(null);   // {t:"contract"|"pay"|"hist"|"maint"|"cost"|"trans"|"tenant"|"prop"|"gen", ...}
  const [view,setView]       = useState(null);   // السند المعروض {k,id,cid?}
  const [toastMsg,setToastMsg] = useState("");
  const toastT = useRef(null);
  const [filt,setFilt]       = useState("all");
  const [mfilt,setMfilt]     = useState("all");
  const [vfilt,setVfilt]     = useState("all");
  const [vq,setVq]           = useState("");
  const [stPid,setStPid]     = useState("");
  const [oPreset,setOPreset] = useState("year");
  const [oFrom,setOFrom]     = useState(()=>new Date().getFullYear()+"-01-01");
  const [oTo,setOTo]         = useState(()=>new Date().getFullYear()+"-12-31");

  const todayKey = pmIso(new Date());
  const today = useMemo(()=>pmToday(),[todayKey]);
  const M = useMemo(()=>pmModel(data,today),[data,today]);

  useEffect(()=>{
    const unsubs = PM_COLLS.map(n=>onSnapshot(collection(db,"pm_"+n),snap=>{
      setData(d=>({...d,[n]:snap.docs.map(x=>({id:x.id,...x.data()}))}));
      setReady(r=>({...r,[n]:true}));
    },err=>{ console.error(err); setLoadErr(true); setReady(r=>({...r,[n]:true})); }));
    const u2 = onSnapshot(collection(db,"providers"),snap=>setProviders(snap.docs.map(x=>x.data().name).filter(Boolean)),()=>{});
    return ()=>{ unsubs.forEach(u=>u()); u2(); };
  },[]);

  const toast = useCallback(m=>{ setToastMsg(m); clearTimeout(toastT.current); toastT.current=setTimeout(()=>setToastMsg(""),2600); },[]);
  const closeMod = ()=>setMod(null);
  const openView = v=>{ setView(v); };

  const voidItem = async (coll,id,no)=>{
    if(!confirm(`إلغاء السند ${no||""}؟\nسيُستبعد من الحسابات ويبقى رقمه ظاهراً كسند ملغي.`)) return;
    const reason = prompt("سبب الإلغاء (اختياري)");
    if(reason===null) return;
    try{ await updateDoc(doc(db,coll,id),{void:true,voidAt:pmNow(),voidReason:reason.trim()}); logAudit("إلغاء سند",coll+"/"+id,`${no||""} — ${reason.trim()}`); toast("✔ تم إلغاء السند"); }
    catch(e){ PMDbErr(e); }
  };
  const setMaintStatus = async (id,status)=>{ try{ await updateDoc(doc(db,"pm_maint",id),{status}); }catch(e){ PMDbErr(e); } };
  const openRenew = c=>{
    const t=M.tenantOf(c), s=pmD(c.end); s.setDate(s.getDate()+1);
    const e=pmAddMonths(s,12); e.setDate(e.getDate()-1);
    setMod({t:"contract",pre:{renew:1,pid:c.pid,unit:c.unit,ten:t.name,phone:t.phone,tid:t.idNo,tdob:t.dob,start:pmIso(s),end:pmIso(e),rent:c.rent,freq:c.freq}});
  };
  const remind = t=>{
    const f=M.tenantFin(t);
    const units=f.cs.map(c=>c.unit+" ("+M.pname(c.pid)+")").join("، ");
    const msg=`السلام عليكم ${t.name}،\nنذكّركم بوجود مبلغ متأخر قدره ${pmMoney(f.arrears)} على ${units}.\nنأمل التكرم بالسداد، وشاكرين لكم تعاونكم.\n${PM_FIRM} — ${PHONE}`;
    const a=document.createElement("a");
    a.href="https://wa.me/966"+String(t.phone||"").replace(/^0/,"")+"?text="+encodeURIComponent(msg);
    a.target="_blank"; a.rel="noopener"; a.click();
  };
  const setPreset = k=>{
    const y=today.getFullYear(), m=today.getMonth();
    setOPreset(k);
    if(k==="all"){ setOFrom(""); setOTo(""); }
    else if(k==="year"){ setOFrom(y+"-01-01"); setOTo(y+"-12-31"); }
    else if(k==="month"){ setOFrom(pmIso(new Date(y,m,1))); setOTo(pmIso(new Date(y,m+1,0))); }
    else if(k==="q3"){ setOFrom(pmIso(pmAddMonths(today,-3))); setOTo(pmIso(today)); }
  };
  const sel = M.props.some(p=>p.id===stPid) ? stPid : (M.props[0]?.id||"");

  const allReady = PM_COLLS.every(n=>ready[n]);

  // ───────── لوحة التحكم ─────────
  const renderDash = ()=>{
    const yr=today.getFullYear(), ins=M.allInst(), late=ins.filter(i=>M.st(i)==="late");
    const collectedY=pmSum(M.contracts.flatMap(c=>c.payments).filter(p=>p.date.startsWith(String(yr))),p=>p.amount);
    const owed=pmSum(M.props,p=>Math.max(0,M.ownerFin(p.id).balance));
    const endingN=M.contracts.filter(c=>M.left(c)>=0&&M.left(c)<=90).length;
    const ofc=M.officeData(yr+"-01-01",yr+"-12-31");
    const al=[];
    late.forEach(i=>{ const dl=-pmDays(today,i.due);
      al.push({pr:3000+dl,k:"late",ic:"⚠️",t:`متأخر ${dl} يوم — ${i.t.name}`,
        s:`${i.c.unit} • ${M.pname(i.c.pid)} • متبقي ${pmMoney(i.rem)}${i.paidAmt>0?` (دفع ${pmMoney(i.paidAmt)} من ${pmMoney(i.amt)})`:""}`,
        b:"تسجيل دفعة",fn:()=>setMod({t:"pay",cid:i.cid,idx:i.idx})}); });
    ins.filter(i=>["soon","part"].includes(M.st(i))).forEach(i=>{ const dl=pmDays(today,i.due);
      al.push({pr:1500-dl,k:"soon",ic:"🔔",t:`قسط مستحق بعد ${dl} يوم — ${i.t.name}`,
        s:`${i.c.unit} • ${pmMoney(i.rem)} بتاريخ ${pmFmt(i.due)}${i.paidAmt>0?" • دفع جزءاً مسبقاً":""}`,
        b:"تسجيل دفعة",fn:()=>setMod({t:"pay",cid:i.cid,idx:i.idx})}); });
    M.contracts.filter(c=>M.left(c)>=0&&M.left(c)<=90).forEach(c=>{
      al.push({pr:2000-M.left(c),k:"end",ic:"📄",t:`عقد ينتهي بعد ${M.left(c)} يوم — ${M.tenantOf(c).name}`,
        s:`${c.unit} • ${M.pname(c.pid)} • ينتهي ${pmFmt(pmD(c.end))}`,b:"تجديد",fn:()=>openRenew(c)}); });
    M.maint.filter(m=>m.status==="جديد").forEach(m=>{
      al.push({pr:1000,k:"mnt",ic:"🔧",t:`طلب صيانة جديد — ${m.title}`,s:`${m.unit} • ${M.pname(m.pid)}`,b:"عرض",fn:()=>setTab("maint")}); });
    al.sort((a,b)=>b.pr-a.pr);
    const months=["ينا","فبر","مار","أبر","ماي","يون","يول","أغس","سبت","أكت","نوف","ديس"];
    const per=Array(12).fill(0);
    M.contracts.flatMap(c=>c.payments).filter(p=>p.date.startsWith(String(yr))).forEach(p=>{ per[+p.date.slice(5,7)-1]+=p.amount; });
    const mx=Math.max(...per,1);
    const totalUnits=pmSum(M.props,p=>p.units.length), occ=pmSum(M.props,p=>M.occupiedUnits(p).length);
    const pct=totalUnits?Math.round(occ/totalUnits*100):0;
    return (
      <div>
        <h2 className="title">لوحة التحكم</h2>
        <p className="sub">ملخص الأملاك المُدارة حتى {pmFmt(today)}</p>
        {!M.props.length&&<div className="card" style={{marginBottom:16}}><b>ابدأ من هنا 👋</b><div style={{fontSize:12,color:"var(--muted)",marginTop:6,lineHeight:1.9}}>1) من تبويب «العقارات» أضف العقار وبيانات المالك والوحدات.<br/>2) من تبويب «العقود» أضف عقد المستأجر — يُنشأ جدول الأقساط تلقائياً.<br/>3) من تبويب «التحصيل» سجّل الدفعات وستُصدر سندات القبض تلقائياً.</div></div>}
        <div className="gridS">
          <div className="card stat"><div className="ic">🏢</div><div className="v">{M.props.length}</div><div className="l">عقارات مُدارة</div></div>
          <div className="card stat green"><div className="ic">💰</div><div className="v">{pmMoney(collectedY)}</div><div className="l">المحصّل في {yr}</div></div>
          <div className="card stat red"><div className="ic">⚠️</div><div className="v">{pmMoney(pmSum(late,i=>i.rem))}</div><div className="l">متأخرات ({late.length} أقساط)</div></div>
          <div className="card stat"><div className="ic">💸</div><div className="v">{pmMoney(owed)}</div><div className="l">مستحق للملاك (لم يُحوَّل)</div></div>
          <div className="card stat amber"><div className="ic">📄</div><div className="v">{endingN}</div><div className="l">عقود تنتهي خلال 90 يوم</div></div>
          <div className={"card stat "+(ofc.net>=0?"green":"red")}><div className="ic">💼</div><div className="v">{pmMoney(ofc.net)}</div><div className="l">صافي دخل المكتب {yr}</div></div>
        </div>
        <div className="cols">
          <div className="card"><div className="ch">🔔 التنبيهات <span className="badge b-blue">{al.length}</span></div>
            <div className="alerts">{al.length?al.map((a,i)=>(
              <div className={"alert "+a.k} key={i}><div className="d">{a.ic}</div><div className="t">{a.t}<small>{a.s}</small></div><button onClick={a.fn}>{a.b}</button></div>
            )):<div className="empty">لا توجد تنبيهات 🎉</div>}</div></div>
          <div className="card">
            <div className="ch">📈 المحصّل شهرياً ({yr}) <small style={{fontWeight:400,color:"var(--muted)",fontSize:10}}>بالآلاف</small></div>
            <div className="bars">{per.map((v,i)=>(<div className="bar" key={i} title={pmMoney(v)}><em>{v?Math.round(v/1000):""}</em><i style={{height:(v?Math.max(4,v/mx*100):0)+"%"}}/><span>{months[i]}</span></div>))}</div>
            <div className="occ"><div className="ch" style={{marginBottom:0}}>🏠 نسبة الإشغال <span style={{fontSize:13}}>{occ} من {totalUnits} وحدة ({pct}%)</span></div>
              <div className="prog"><i style={{width:pct+"%"}}/></div></div>
          </div>
        </div>
      </div>
    );
  };

  // ───────── العقارات ─────────
  const renderProps = ()=>(
    <div>
      <div className="row"><div><h2 className="title">العقارات المُدارة</h2><p className="sub">العقارات التي تديرها المؤسسة نيابة عن الملاك — الوحدات وحالة الإشغال والرصيد المستحق للمالك</p></div>
        <button className="btn blue" onClick={()=>setMod({t:"prop"})}>+ إضافة عقار</button></div>
      {!M.props.length&&<div className="empty">لم تُضف عقارات بعد</div>}
      <div className="pgrid">{M.props.map(p=>{
        const occ=M.occupiedUnits(p), f=M.ownerFin(p.id);
        const monthly=pmSum(M.contracts.filter(c=>c.pid===p.id&&M.left(c)>=0),c=>c.rent/12);
        return (
          <div className="card pc" key={p.id}>
            <div className="head"><div className="pi">{p.icon}</div><div><div className="pn">{p.name}</div><div className="pa">📍 {p.area} • {p.owner}</div></div></div>
            <div className="units">{p.units.map(u=><span key={u} className={"uc "+(occ.includes(u)?"occ":"vac")}>{u}{occ.includes(u)?"":" • شاغرة"}</span>)}</div>
            <div className="prog"><i style={{width:Math.round(occ.length/p.units.length*100)+"%"}}/></div>
            <div className="kv"><span>🪪 هوية المالك</span><b>{p.ownerId||"—"}{p.ownerDob?" • "+M.dobTxt(p.ownerDob):""}</b></div>
            <div className="kv"><span>📞 جوال المالك</span><b>{p.ownerPhone||"—"}</b></div>
            <div className="kv"><span>💵 الدخل الشهري المتعاقد</span><b>{pmMoney(monthly)}</b></div>
            <div className="kv"><span>✅ المحصّل الفعلي</span><b>{pmMoney(f.collected)}</b></div>
            <div className="kv"><span>⚠️ متأخرات المستأجرين</span><b style={{color:f.arrears?"var(--red)":"inherit"}}>{pmMoney(f.arrears)}</b></div>
            <div className="kv"><span>📋 رسوم الإدارة</span><b>{p.fee}%</b></div>
            <div className="kv"><span>💸 المتبقي للمالك</span><b style={{color:"var(--green)"}}>{pmMoney(f.balance)}</b></div>
            <div className="btns">
              <button className="btn ghost" onClick={()=>{setStPid(p.id);setTab("owners");}}>🧾 الكشف</button>
              <button className="btn" onClick={()=>setMod({t:"trans",pid:p.id})}>💸 تحويل</button>
              <button className="btn ghost" onClick={()=>setMod({t:"prop",id:p.id})}>✏️ تعديل</button>
            </div>
          </div>);
      })}</div>
    </div>
  );

  // ───────── المستأجرين ─────────
  const renderTenants = ()=>(
    <div>
      <h2 className="title">المستأجرين</h2><p className="sub">رصيد كل مستأجر وكشف حسابه وتذكير واتساب بالمتأخرات</p>
      {!M.tenants.length&&<div className="empty">لا يوجد مستأجرون — يُضافون تلقائياً عند إنشاء العقود</div>}
      <div className="tgrid">{M.tenants.map(t=>{
        const f=M.tenantFin(t);
        return (
          <div className="card tc" key={t.id}>
            <div className="tn">{t.name}</div><div className="tp">📞 {t.phone||"—"}</div>
            <div className="kv"><span>🪪 الهوية</span><b>{t.idNo||"—"}</b></div>
            <div className="kv"><span>🎂 تاريخ الميلاد</span><b>{M.dobTxt(t.dob)}</b></div>
            {f.cs.map(c=><div className="kv" key={c.id}><span>{c.unit} • {M.pname(c.pid)}</span><b>{pmMoney(c.rent)}/سنة</b></div>)}
            <div className="kv"><span>المستحق حتى اليوم</span><b>{pmMoney(f.dueTo)}</b></div>
            <div className="kv"><span>إجمالي المدفوع</span><b style={{color:"var(--green)"}}>{pmMoney(f.paid)}</b></div>
            <div className="kv"><span>المتأخر</span><b style={{color:f.arrears?"var(--red)":"inherit"}}>{pmMoney(f.arrears)}</b></div>
            <div className="btns">
              <button className="btn ghost sm" onClick={()=>setMod({t:"tenant",id:t.id})}>✏️ تعديل</button>
              <button className="btn ghost sm" onClick={()=>pmPrint(pmTenantStmtHTML(M,t.id))}>🖨️ كشف حساب</button>
              {f.arrears>0&&<button className="btn wa sm" onClick={()=>remind(t)}>💬 تذكير واتساب</button>}
            </div>
          </div>);
      })}</div>
    </div>
  );

  // ───────── العقود ─────────
  const renderContracts = ()=>(
    <div>
      <div className="row"><div><h2 className="title">عقود الإيجار</h2><p className="sub">العقود مرتبة حسب أقرب انتهاء — مع التحصيل الفعلي لكل عقد وزر التجديد</p></div>
        <button className="btn blue" onClick={()=>setMod({t:"contract"})}>+ إضافة عقد</button></div>
      <div className="card"><div className="tw"><table>
        <thead><tr><th>العقد</th><th>رقم إيجار</th><th>المستأجر</th><th>العقار / الوحدة</th><th>المدة</th><th>الإيجار السنوي</th><th>المحصّل</th><th>الحالة</th><th></th></tr></thead>
        <tbody>{M.contracts.length?M.contracts.slice().sort((a,b)=>M.left(a)-M.left(b)).map(c=>{
          const l=M.left(c);
          const b=l<0?["b-red","منتهي"]:l<=30?["b-red",`ينتهي بعد ${l} يوم`]:l<=90?["b-amber",`ينتهي بعد ${l} يوم`]:["b-green","ساري"];
          const got=pmSum(c.payments,p=>p.amount), pct=Math.min(100,Math.round(got/c.rent*100));
          return (
            <tr key={c.id}><td><b>{c.cno}</b></td><td>{c.ejar||"—"}</td><td>{M.tenantOf(c).name}</td>
              <td>{c.unit}<br/><small style={{color:"var(--muted)"}}>{M.pname(c.pid)}</small></td>
              <td>{pmFmt(pmD(c.start))} ← {pmFmt(pmD(c.end))}</td><td><b>{pmMoney(c.rent)}</b></td>
              <td style={{minWidth:130}}><b>{pmMoney(got)}</b><div className="prog sm"><i style={{width:pct+"%"}}/></div><small style={{color:"var(--muted)"}}>{pct}% من العقد</small></td>
              <td><span className={"badge "+b[0]}>{b[1]}</span></td>
              <td>{l<=90&&<button className="btn ghost sm" onClick={()=>openRenew(c)}>تجديد</button>}</td></tr>);
        }):<tr><td colSpan={9} className="empty">لا توجد عقود</td></tr>}</tbody></table></div></div>
    </div>
  );

  // ───────── التحصيل ─────────
  const renderCollect = ()=>{
    const ins=M.allInst();
    const order=i=>M.st(i)==="paid"?1:0;
    ins.sort((a,b)=>order(a)-order(b)||(order(a)?b.due-a.due:a.due-b.due));
    const cnt={all:ins.length}; ["late","part","soon","up","paid"].forEach(k=>{ cnt[k]=ins.filter(i=>M.st(i)===k).length; });
    const labels={all:"الكل",late:"متأخرة",part:"جزئية",soon:"مستحقة قريباً",up:"قادمة",paid:"مدفوعة"};
    const list=ins.filter(i=>filt==="all"||M.st(i)===filt);
    return (
      <div>
        <h2 className="title">التحصيل</h2><p className="sub">سجّل الدفعة كاملة أو جزئية — الباقي يبقى مفتوحاً على الدفعة، ولكل دفعة سند قبض</p>
        <div className="chips">{Object.keys(labels).map(k=><button key={k} className={"chip"+(filt===k?" on":"")} onClick={()=>setFilt(k)}>{labels[k]} ({cnt[k]})</button>)}</div>
        {list.length?list.map(i=>{
          const s=M.st(i), m=PM_ST_META[s], dl=pmDays(today,i.due), pct=Math.min(100,Math.round(i.paidAmt/i.amt*100));
          const note=s==="paid"?"":dl<0?`متأخر ${-dl} يوم`:`بعد ${dl} يوم`;
          return (
            <div className="crow" key={i.cid+"-"+i.idx}>
              <div className="who"><b>{i.t.name}</b><small>{i.c.unit} • {M.pname(i.c.pid)} • قسط {i.idx+1}/{i.c.freq}</small></div>
              <div className="due">📅 {pmFmt(i.due)}<br/><small>{note}</small></div>
              <div className="amtbox"><div className="amt">{pmMoney(i.amt)}</div><div className="prog sm"><i style={{width:pct+"%"}}/></div>
                <small>مدفوع {pmMoney(i.paidAmt)}{i.rem>0.5&&<> • متبقي <b>{pmMoney(i.rem)}</b></>}</small></div>
              <span className={"badge "+m[0]}>{m[1]}</span>
              <div className="acts">
                {i.rem>0.5&&<button className="btn blue sm" onClick={()=>setMod({t:"pay",cid:i.cid,idx:i.idx})}>تسجيل دفعة</button>}
                {i.c.payments.some(p=>p.idx===i.idx)&&<button className="btn ghost sm" onClick={()=>setMod({t:"hist",cid:i.cid,idx:i.idx})}>السجل</button>}
              </div>
            </div>);
        }):<div className="empty">لا توجد أقساط</div>}
      </div>
    );
  };

  // ───────── السندات ─────────
  const renderVouchers = ()=>{
    const all=M.voucherList(), live=all.filter(v=>!v.void);
    const inn=live.filter(v=>v.kind==="in"), out=live.filter(v=>v.kind==="out");
    const tin=pmSum(inn,v=>v.amount), tout=pmSum(out,v=>v.amount);
    const labels={all:["الكل",all.length],in:["سندات القبض",all.filter(v=>v.kind==="in").length],out:["سندات الصرف",all.filter(v=>v.kind==="out").length]};
    const q=vq.trim().toLowerCase();
    const list=all.filter(v=>(vfilt==="all"||v.kind===vfilt)&&(!q||`${v.no} ${v.party} ${v.desc}`.toLowerCase().includes(q)));
    return (
      <div>
        <h2 className="title">السندات</h2><p className="sub">سندات القبض (من المستأجرين + عامة) وسندات الصرف (للملاك ولمزودي الخدمات + عامة) — اضغط "عرض" لمعاينة السند ثم اطبعه</p>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:14}}>
          <button className="btn blue" onClick={()=>setMod({t:"gen",kind:"in"})}>+ سند قبض عام</button>
          <button className="btn" onClick={()=>setMod({t:"gen",kind:"out"})}>+ سند صرف عام</button>
        </div>
        <div className="gridS">
          <div className="card stat green"><div className="ic">📥</div><div className="v">{pmMoney(tin)}</div><div className="l">سندات القبض ({inn.length})</div></div>
          <div className="card stat red"><div className="ic">📤</div><div className="v">{pmMoney(tout)}</div><div className="l">سندات الصرف ({out.length})</div></div>
          <div className="card stat"><div className="ic">⚖️</div><div className="v">{pmMoney(tin-tout)}</div><div className="l">الفرق (مقبوض − مصروف)</div></div>
        </div>
        <div className="chips">{Object.keys(labels).map(k=><button key={k} className={"chip"+(vfilt===k?" on":"")} onClick={()=>setVfilt(k)}>{labels[k][0]} ({labels[k][1]})</button>)}</div>
        <div style={{marginBottom:12}}><input value={vq} onChange={e=>setVq(e.target.value)} placeholder="🔍 ابحث برقم السند أو الاسم أو البيان..."/></div>
        <div className="card"><div className="tw"><table>
          <thead><tr><th>رقم السند</th><th>النوع</th><th>التاريخ</th><th>الطرف</th><th>البيان</th><th>المبلغ</th><th></th></tr></thead>
          <tbody>{list.length?list.map((v,i)=>(
            <tr key={v.kind+v.no+i} className={v.void?"void-row":""}>
              <td><b>{v.no}</b></td>
              <td><span className={"badge "+(v.kind==="in"?"b-green":"b-amber")}>{v.gen?(v.kind==="in"?"قبض عام":"صرف عام"):(v.kind==="in"?"سند قبض":"سند صرف")}</span>{v.trust&&<> <span className="badge b-gray">أمانة</span></>}{v.void&&<> <span className="badge b-red">ملغي</span></>}</td>
              <td>{pmFmt(pmD(v.date))}</td><td>{v.party}</td><td>{v.desc}</td>
              <td><b style={{color:v.kind==="in"?"var(--green)":"var(--red)"}}>{v.kind==="in"?"+":"−"} {pmMoney(v.amount)}</b></td>
              <td style={{whiteSpace:"nowrap"}}>
                <button className="btn ghost sm" onClick={()=>openView(v.view)}>👁 عرض / طباعة</button>
                {v.voidRef&&!v.void&&<> <button className="btn ghost sm" style={{color:"var(--red)"}} onClick={()=>voidItem(v.voidRef[0],v.voidRef[1],v.no)}>إلغاء</button></>}
              </td></tr>
          )):<tr><td colSpan={7} className="empty">لا توجد سندات</td></tr>}</tbody></table></div></div>
      </div>
    );
  };

  // ───────── دخل المكتب ─────────
  const renderOffice = ()=>{
    const d=M.officeData(oFrom,oTo);
    const labels={month:"هذا الشهر",q3:"آخر 3 أشهر",year:"هذه السنة",all:"الكل"};
    const rows=(arr,total,label,cls)=>(
      <>
        {arr.length?arr.map((l,i)=>(
          <tr key={i}><td><span className={"badge "+(l.auto?"b-blue":cls)}>{l.no}</span></td><td>{l.date?pmFmt(pmD(l.date)):"—"}</td>
            <td><b>{l.cat}</b><br/><small style={{color:"var(--muted)"}}>{l.desc}</small></td><td><b>{pmMoney(l.amount)}</b></td></tr>
        )):<tr><td colSpan={4} className="empty">لا توجد بنود في هذه الفترة</td></tr>}
        <tr style={{background:"var(--soft)"}}><td colSpan={3}><b>{label}</b></td><td><b style={{fontSize:14}}>{pmMoney(total)}</b></td></tr>
      </>
    );
    return (
      <div>
        <div className="row"><div><h2 className="title">دخل المكتب</h2><p className="sub">دخل المؤسسة نفسها: رسوم الإدارة (تلقائي من تحصيل الملاك) + السندات العامة — بسند ومبلغ وبيان، والإجمالي</p></div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            <button className="btn blue" onClick={()=>setMod({t:"gen",kind:"in"})}>+ سند قبض عام</button>
            <button className="btn" onClick={()=>setMod({t:"gen",kind:"out"})}>+ سند صرف عام</button>
            <button className="btn ghost" onClick={()=>pmPrint(pmOfficeHTML(M,oFrom,oTo))}>🖨️ طباعة الكشف</button>
          </div></div>
        <div className="chips">{Object.keys(labels).map(k=><button key={k} className={"chip"+(oPreset===k?" on":"")} onClick={()=>setPreset(k)}>{labels[k]}</button>)}</div>
        <div className="fg" style={{maxWidth:440}}>
          <div><label>من تاريخ</label><input type="date" value={oFrom} onChange={e=>{setOFrom(e.target.value);setOPreset("custom");}}/></div>
          <div><label>إلى تاريخ</label><input type="date" value={oTo} onChange={e=>{setOTo(e.target.value);setOPreset("custom");}}/></div>
        </div>
        <div className="gridS">
          <div className="card stat green"><div className="ic">📥</div><div className="v">{pmMoney(d.tin)}</div><div className="l">إجمالي الدخل ({d.inc.length} بنود)</div></div>
          <div className="card stat red"><div className="ic">📤</div><div className="v">{pmMoney(d.tout)}</div><div className="l">إجمالي المصروفات ({d.exp.length} بنود)</div></div>
          <div className={"card stat "+(d.net>=0?"green":"red")}><div className="ic">💼</div><div className="v">{pmMoney(d.net)}</div><div className="l">صافي دخل المكتب</div></div>
          <div className="card stat"><div className="ic">🔒</div><div className="v">{pmMoney(d.trust)}</div><div className="l">أمانات محتجزة (غير محتسبة في الدخل)</div></div>
        </div>
        <div className="card" style={{marginBottom:14}}><div className="ch">📥 الدخل</div><div className="tw"><table style={{minWidth:560}}>
          <thead><tr><th>السند</th><th>التاريخ</th><th>عن</th><th>المبلغ</th></tr></thead><tbody>{rows(d.inc,d.tin,"إجمالي الدخل","b-green")}</tbody></table></div></div>
        <div className="card"><div className="ch">📤 المصروفات</div><div className="tw"><table style={{minWidth:560}}>
          <thead><tr><th>السند</th><th>التاريخ</th><th>عن</th><th>المبلغ</th></tr></thead><tbody>{rows(d.exp,d.tout,"إجمالي المصروفات","b-amber")}</tbody></table></div></div>
        <div className="card" style={{marginTop:14}}><div className="ch">🔧 الصيانة — التكلفة والربح وصافي الربح</div>
          <div className="gridS" style={{marginBottom:12}}>
            <div className="card stat"><div className="ic">🔧</div><div className="v">{pmMoney(d.mCost)}</div><div className="l">إجمالي التكلفة (المدفوع للمزودين)</div></div>
            <div className="card stat"><div className="ic">🏷️</div><div className="v">{pmMoney(d.mChgT)}</div><div className="l">المحمّل على الملاك</div></div>
            <div className="card stat green"><div className="ic">📈</div><div className="v">{pmMoney(d.mProfit)}</div><div className="l">الربح (من أعمال الملاك)</div></div>
            <div className={"card stat "+(d.mNet>=0?"green":"red")}><div className="ic">💼</div><div className="v">{pmMoney(d.mNet)}</div><div className="l">صافي الربح بعد صيانة على المؤسسة ({pmMoney(d.mFirm)})</div></div>
          </div>
          <div className="tw"><table style={{minWidth:680}}>
            <thead><tr><th>السند</th><th>التاريخ</th><th>العمل</th><th>التكلفة</th><th>المحمّل</th><th>الربح / الخسارة</th></tr></thead>
            <tbody>
              {d.mj.length?d.mj.map(m=>{ const pf=M.mProfit(m); return (
                <tr key={m.id}><td><span className="badge b-amber">{m.vno||"—"}</span></td><td>{pmFmt(pmD(m.paidDate||m.date))}</td>
                  <td><b>{m.title}</b><br/><small style={{color:"var(--muted)"}}>{M.pname(m.pid)} • {m.unit} • على {m.payer==="owner"?"المالك":"المؤسسة"}</small></td>
                  <td>{pmMoney(m.cost)}</td><td>{pmMoney(M.mChg(m))}</td>
                  <td><b style={{color:pf>=0?"var(--green)":"var(--red)"}}>{pf<0?"−":""}{pmMoney(Math.abs(pf))}</b></td></tr>); })
                :<tr><td colSpan={6} className="empty">لا توجد أعمال صيانة مكتملة في هذه الفترة</td></tr>}
              <tr style={{background:"var(--soft)"}}><td colSpan={3}><b>الإجمالي (صافي الربح)</b></td><td><b>{pmMoney(d.mCost)}</b></td><td><b>{pmMoney(d.mChgT)}</b></td>
                <td><b style={{fontSize:14,color:d.mNet>=0?"var(--green)":"var(--red)"}}>{d.mNet<0?"−":""}{pmMoney(Math.abs(d.mNet))}</b></td></tr>
            </tbody></table></div></div>
      </div>
    );
  };

  // ───────── الصيانة ─────────
  const renderMaint = ()=>{
    const labels={all:"الكل","جديد":"جديد","جاري":"جاري","مكتمل":"مكتمل"};
    const bc={"جديد":"b-blue","جاري":"b-amber","مكتمل":"b-green"};
    const list=M.maint.filter(m=>mfilt==="all"||m.status===mfilt).sort((a,b)=>b.date.localeCompare(a.date));
    return (
      <div>
        <div className="row"><div><h2 className="title">طلبات الصيانة</h2><p className="sub">عند إتمام الطلب وإدخال التكلفة تُخصم تلقائياً من كشف المالك (إذا كانت عليه)</p></div>
          <button className="btn blue" onClick={()=>setMod({t:"maint"})}>+ طلب صيانة</button></div>
        <div className="chips">{Object.keys(labels).map(k=><button key={k} className={"chip"+(mfilt===k?" on":"")} onClick={()=>setMfilt(k)}>{labels[k]} ({k==="all"?M.maint.length:M.maint.filter(m=>m.status===k).length})</button>)}</div>
        <div className="mgrid">{list.length?list.map(m=>(
          <div className="card mc" key={m.id}>
            <div className="row" style={{marginBottom:6}}><span className={"badge "+bc[m.status]}>{m.status}</span><small style={{color:"var(--muted)"}}>{pmFmt(pmD(m.date))}</small></div>
            <div className="mt">{m.title}</div><div className="ms">{M.pname(m.pid)} • {m.unit}</div>
            <div className="kv"><span>🔧 المزود</span><b>{m.provider}</b></div>
            {m.status==="مكتمل"?(<>
              <div className="kv"><span>💵 التكلفة (للمزود)</span><b>{pmMoney(m.cost)}</b></div>
              <div className="kv"><span>🏷️ المحمّل على {m.payer==="owner"?"المالك":"المؤسسة"}</span><b>{pmMoney(M.mChg(m))}</b></div>
              <div className="kv"><span>{M.mProfit(m)>=0?"📈 الربح":"📉 الخسارة"}</span><b style={{color:M.mProfit(m)>=0?"var(--green)":"var(--red)"}}>{pmMoney(Math.abs(M.mProfit(m)))}</b></div>
            </>):<div className="kv"><span>💵 التكلفة</span><b>—</b></div>}
            <div className="btns">
              {m.status==="جديد"&&<button className="btn ghost sm" onClick={()=>setMaintStatus(m.id,"جاري")}>▶ بدء التنفيذ</button>}
              {m.status!=="مكتمل"&&<button className="btn blue sm" onClick={()=>setMod({t:"cost",id:m.id})}>✔ إتمام</button>}
              {m.status==="مكتمل"&&m.cost>0&&<button className="btn ghost sm" onClick={()=>openView({k:"m",id:m.id})}>🧾 سند الصرف</button>}
            </div>
          </div>
        )):<div className="empty">لا توجد طلبات</div>}</div>
      </div>
    );
  };

  // ───────── كشوف الملاك ─────────
  const renderOwners = ()=>(
    <div>
      <div className="row"><div><h2 className="title">كشوف الملاك</h2><p className="sub">الإيرادات والمصروفات ورسوم الإدارة والتحويلات — والمتبقي للمالك</p></div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <select value={sel} onChange={e=>setStPid(e.target.value)} style={{minWidth:200}}>{M.props.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select>
          <button className="btn ghost" onClick={()=>setMod({t:"trans",pid:sel})}>💸 تسجيل تحويل للمالك</button>
          <button className="btn" disabled={!sel} onClick={()=>pmPrint(pmStmtHTML(M,sel))}>🖨️ طباعة الكشف</button>
        </div></div>
      {sel?<div className="stmtWrap" dangerouslySetInnerHTML={{__html:pmStmtHTML(M,sel)}}/>:<div className="empty">أضف عقاراً لعرض كشف المالك</div>}
    </div>
  );

  const body = { dash:renderDash, props:renderProps, tenants:renderTenants, contracts:renderContracts, collect:renderCollect, vouchers:renderVouchers, office:renderOffice, maint:renderMaint, owners:renderOwners }[tab];

  return (
    <div className={"pm"+(darkMode?" dark":"")} dir="rtl" style={{paddingTop:96,minHeight:"100vh"}}>
      <style>{PM_CSS}</style>
      <div style={{background:"linear-gradient(135deg,#1e3a7a,#2a4d9b)",padding:"22px 24px 18px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,flexWrap:"wrap"}}>
        <div><div style={{color:"#fff",fontWeight:900,fontSize:20}}>🏛️ نظام إدارة الأملاك</div><div style={{color:"#7ab8ff",fontSize:12,marginTop:2}}>العقارات والمستأجرين والعقود والتحصيل والسندات وكشوف الملاك</div></div>
        <span style={{background:"rgba(74,158,255,.18)",border:"1px solid rgba(74,158,255,.35)",color:"#7ab8ff",fontSize:10,fontWeight:700,borderRadius:20,padding:"4px 12px"}}>👑 المدير فقط</span>
      </div>
      <div className="pmtabs">{PM_TABS.map(([k,l])=><button key={k} className={"pmtab"+(tab===k?" on":"")} onClick={()=>{setTab(k);window.scrollTo(0,0);}}>{l}</button>)}</div>
      <div className="wrap">
        {loadErr&&<div className="card" style={{marginBottom:14,borderColor:"var(--red)"}}>⚠️ تعذّر تحميل بعض البيانات من Firebase — تحقق من الاتصال وقواعد Firestore.</div>}
        {!allReady?<div className="empty">جاري التحميل...</div>:body()}
      </div>

      {mod?.t==="contract"&&<PMContractModal M={M} pre={mod.pre} onClose={closeMod} toast={toast}/>}
      {mod?.t==="pay"&&<PMPayModal M={M} cid={mod.cid} idx={mod.idx} onClose={closeMod} onSaved={(cid,id)=>{ closeMod(); toast("✔ تم تسجيل الدفعة وإصدار سند القبض"); openView({k:"r",cid,id}); }}/>}
      {mod?.t==="hist"&&<PMHistModal M={M} cid={mod.cid} idx={mod.idx} onClose={closeMod} onPay={(cid,idx)=>setMod({t:"pay",cid,idx})} onView={openView} onVoid={voidItem}/>}
      {mod?.t==="maint"&&<PMMaintModal M={M} providers={providers} onClose={closeMod} toast={toast}/>}
      {mod?.t==="cost"&&<PMCostModal M={M} id={mod.id} onClose={closeMod} toast={toast} onView={openView}/>}
      {mod?.t==="trans"&&<PMTransModal M={M} pid0={mod.pid} onClose={closeMod} toast={toast} onView={openView}/>}
      {mod?.t==="tenant"&&<PMTenantModal M={M} id={mod.id} onClose={closeMod} toast={toast}/>}
      {mod?.t==="prop"&&<PMPropModal M={M} id={mod.id} onClose={closeMod} toast={toast}/>}
      {mod?.t==="gen"&&<PMGenModal M={M} kind={mod.kind} onClose={closeMod} toast={toast} onView={openView}/>}
      {view&&<PMVoucherModal html={pmViewHTML(M,view)} onClose={()=>setView(null)}/>}
      {toastMsg&&<div className="pmtoast">{toastMsg}</div>}
    </div>
  );
}

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
  const [privMap,setPrivMap]     = useState({});
  const justLogged = useRef(false);

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

  // ── Firebase Authentication: الدور يُقرأ من users/{uid} ──
  useEffect(()=>onAuthStateChanged(auth, async (u)=>{
    if(!u){ setIsAdmin(false); setUserRole(null); setCurrentUser(null); return; }
    try{
      const snap = await getDoc(doc(db,"users",u.uid));
      const d = snap.exists() ? snap.data() : null;
      if(!d || d.active===false || !["admin","employee"].includes(d.role)){
        await signOut(auth);
        showToast("هذا الحساب لا يملك صلاحية دخول","err");
        return;
      }
      setCurrentUser({uid:u.uid,email:u.email,displayName:d.displayName||u.email,role:d.role,username:d.username||""});
      setUserRole(d.role); setIsAdmin(true);
      if(justLogged.current){ justLogged.current=false; showToast(d.role==="admin"?`مرحباً ${d.displayName||""}! 👑`:`مرحباً ${d.displayName||""}! 👤`); }
    }catch(e){ console.error(e); try{ await signOut(auth); }catch(_){} showToast("تعذّر التحقق من الحساب","err"); }
  }),[]);

  // خروج تلقائي بعد 30 دقيقة بلا نشاط
  useEffect(()=>{
    if(!isAdmin) return;
    let t; const reset=()=>{ clearTimeout(t); t=setTimeout(()=>{ signOut(auth); showToast("تم تسجيل الخروج لعدم النشاط","err"); },30*60*1000); };
    const ev=["mousemove","keydown","click","touchstart","scroll"];
    ev.forEach(e=>window.addEventListener(e,reset,{passive:true})); reset();
    return ()=>{ clearTimeout(t); ev.forEach(e=>window.removeEventListener(e,reset)); };
  },[isAdmin]);

  // بيانات المالك الخاصة (للمدير فقط) تُدمج مع العقار للعرض
  useEffect(()=>{
    if(userRole!=="admin"){ setPrivMap({}); return; }
    const unsub = onSnapshot(collection(db,"properties_private"), snap=>{ const m={}; snap.docs.forEach(d=>{ m[d.id]=d.data(); }); setPrivMap(m); }, ()=>{});
    return ()=>unsub();
  },[userRole]);
  const [backupDue,setBackupDue]=useState(false);
  useEffect(()=>{
    if(userRole!=="admin"){ setBackupDue(false); return; }
    getDoc(doc(db,"meta","backup")).then(d=>{
      const t = d.exists()&&d.data().lastAt ? new Date(d.data().lastAt).getTime() : 0;
      setBackupDue(Date.now()-t > 30*24*3600*1000);
    }).catch(()=>{});
  },[userRole,page]);
  const propsView = useMemo(()=>props.map(p=>privMap[p.id]?{...p,...privMap[p.id]}:p),[props,privMap]);

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
      const { ownerName="", ownerPhone="", id:_id, ...pub } = { ...form, updatedAt:today() };
      if(editId){
        // بيانات المالك تُحفظ في مجموعة خاصة (المدير فقط) — الوثيقة العامة لا تحملها
        if(userRole==="admin") await setDoc(doc(db,"properties_private",editId),{ownerName,ownerPhone},{merge:true});
        await updateDoc(doc(db,"properties",editId),{...pub,ownerName:deleteField(),ownerPhone:deleteField()});
        showToast(isEn?"Updated ✓":"تم التحديث ✓");
      } else {
        const ref = doc(collection(db,"properties"));
        if(userRole==="admin"&&(ownerName||ownerPhone)) await setDoc(doc(db,"properties_private",ref.id),{ownerName,ownerPhone});
        await setDoc(ref,{...pub,views:0,createdAt:today()});
        showToast(isEn?"Added ✓":"تم الإضافة ✓");
      }
    } catch(e){ showToast(isEn?"Error saving":"خطأ في الحفظ","err"); }
    setSaving(false);
    setShowForm(false);
  };

  const del=async(id)=>{
    try { const pr=props.find(x=>x.id===id); await deleteDoc(doc(db,"properties",id)); try{ await deleteDoc(doc(db,"properties_private",id)); }catch(e){} logAudit("حذف عقار",id,pr?.name||""); showToast(isEn?"Deleted":"تم الحذف"); }
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
      {showLogin&&<LoginModal onAttempt={()=>{justLogged.current=true;}} onClose={()=>setShowLogin(false)} lang={lang}/>}
      {showForm&&<PropForm form={form} setForm={setForm} onSave={save} onClose={()=>setShowForm(false)} editId={editId} T={T} userRole={userRole}/>}
      {shareP&&<ShareModal p={shareP} onClose={()=>setShareP(null)}/>}

      {backupDue&&page!=="security"&&userRole==="admin"&&(<div onClick={()=>setPage("security")} style={{position:"fixed",bottom:20,right:20,background:"#b45309",color:"#fff",padding:"10px 18px",borderRadius:12,zIndex:9998,fontWeight:700,fontSize:13,cursor:"pointer",boxShadow:"0 8px 32px #000a"}}>⚠️ حان وقت النسخة الاحتياطية الشهرية — اضغط هنا</div>)}
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

      <Navbar page={page} setPage={setPage} isAdmin={isAdmin} onLoginClick={()=>setShowLogin(true)} onLogout={()=>{signOut(auth);setPage("home");showToast(isEn?"Logged out":"تم تسجيل الخروج");}} lang={lang} setLang={setLang} scrolled={scrolled} darkMode={darkMode} setDarkMode={setDarkMode} T={T} userRole={userRole}/>

      {page==="home"       && <HomePage setPage={setPage} lang={lang} darkMode={darkMode} T={T}/>}
      {page==="properties" && <PropertiesPage props={propsView} isAdmin={isAdmin} userRole={userRole} onEdit={openEdit} onDelete={id=>setDelId(id)} onChangeStatus={changeStatus} setLightbox={setLightbox} onShare={setShareP} onOpenAdd={openAdd} lang={lang} darkMode={darkMode} T={T}/>}
      {page==="services"   && <HomePage setPage={setPage} lang={lang} darkMode={darkMode} T={T}/>}
      {page==="about"      && <AboutPage lang={lang} darkMode={darkMode} T={T}/>}
      {page==="clients"    && isAdmin && <ClientsPage lang={lang} darkMode={darkMode} T={T} userRole={userRole} currentUser={currentUser}/>}
      {page==="providers"  && userRole==="admin" && <ProvidersPage lang={lang} darkMode={darkMode} T={T}/>}
      {page==="employees"  && userRole==="admin" && <EmployeesPage lang={lang} darkMode={darkMode} T={T} currentUser={currentUser}/>}
      {page==="security"   && userRole==="admin" && <SecurityPage lang={lang} darkMode={darkMode} T={T} currentUser={currentUser} rawProps={props}/>}
      {page==="pm"         && userRole==="admin" && <PropertyMgmtPage lang={lang} darkMode={darkMode} T={T}/>}

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
