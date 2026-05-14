
import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import Papa from "papaparse";

/* ─── localStorage shim ─── */
if (typeof window !== "undefined") {
  window.storage = {
    get: (key) => {
      const v = localStorage.getItem(key);
      if (v === null) return Promise.reject(new Error("Key not found: " + key));
      return Promise.resolve({ value: v });
    },
    set: (key, value) => {
      try { localStorage.setItem(key, String(value)); } catch(e) { return Promise.resolve(null); }
      return Promise.resolve({ key, value });
    },
    delete: (key) => {
      localStorage.removeItem(key);
      return Promise.resolve({ key, deleted: true });
    },
    list: (prefix) => {
      const keys = Object.keys(localStorage).filter(k => !prefix || k.startsWith(prefix));
      return Promise.resolve({ keys });
    }
  };
}


/* ─── GLOBAL STYLES ─── */
const GS = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=Nunito:wght@400;500;600;700&display=swap');
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    html, body, #root { height: 100%; width: 100%; overflow: hidden; }
    :root {
      --bg:#0F0E0C; --bg2:#1A1814; --bg3:#232017; --border:#302C22;
      --amber:#E8930A; --amber2:#F5B234; --amber-bg:rgba(232,147,10,0.10);
      --green:#2DB87A; --green-bg:rgba(45,184,122,0.10);
      --red:#E85A4F;   --red-bg:rgba(232,90,79,0.10);
      --blue:#4A9EE8;  --blue-bg:rgba(74,158,232,0.10);
      --text:#F0EAD6;  --text2:#A09070; --text3:#5A5040;
      --radius:14px;   --radius-sm:8px;
    }
    body { background:var(--bg); font-family:'Nunito',sans-serif; color:var(--text); -webkit-font-smoothing:antialiased; }
    ::-webkit-scrollbar { width:3px; }
    ::-webkit-scrollbar-thumb { background:var(--border); border-radius:2px; }

    .shell { display:flex; flex-direction:column; height:100%; max-width:430px; margin:0 auto; background:var(--bg); position:relative; overflow:hidden; }

    /* TOP BAR */
    .topbar { display:flex; align-items:center; justify-content:space-between; padding:14px 18px 10px; border-bottom:1px solid var(--border); background:var(--bg2); flex-shrink:0; position:relative; z-index:10; }
    .topbar-left { display:flex; align-items:center; gap:10px; }
    .logo-pill { background:var(--amber); color:#000; font-family:'Syne',sans-serif; font-weight:800; font-size:13px; padding:5px 10px; border-radius:8px; }
    .topbar-title { font-family:'Syne',sans-serif; font-size:17px; font-weight:700; }
    .topbar-right { display:flex; gap:8px; }
    .icon-btn { width:36px; height:36px; background:var(--bg3); border:1px solid var(--border); border-radius:10px; display:flex; align-items:center; justify-content:center; cursor:pointer; font-size:16px; transition:background 0.15s; flex-shrink:0; position:relative; }
    .icon-btn:hover { background:var(--border); }
    .doc-badge { position:absolute; top:-4px; right:-4px; width:16px; height:16px; background:var(--amber); border-radius:50%; font-size:9px; font-weight:700; color:#000; display:flex; align-items:center; justify-content:center; }
    .key-dot { position:absolute; top:-3px; right:-3px; width:10px; height:10px; border-radius:50%; border:2px solid var(--bg2); }

    /* CHAT */
    .view { flex:1; display:flex; flex-direction:column; overflow:hidden; }
    .chat-scroll { flex:1; overflow-y:auto; padding:16px; display:flex; flex-direction:column; gap:14px; }

    .welcome { flex:1; display:flex; flex-direction:column; align-items:center; justify-content:center; padding:24px 20px; text-align:center; }
    .welcome-avatar { width:80px; height:80px; background:linear-gradient(135deg,var(--amber),var(--amber2)); border-radius:24px; display:flex; align-items:center; justify-content:center; font-size:38px; margin-bottom:18px; box-shadow:0 8px 32px rgba(232,147,10,0.25); }
    .welcome-title { font-family:'Syne',sans-serif; font-size:24px; font-weight:800; margin-bottom:8px; }
    .welcome-sub { font-size:14px; color:var(--text2); line-height:1.65; max-width:280px; }
    .welcome-chips { display:flex; flex-direction:column; gap:8px; width:100%; margin-top:24px; }
    .chip { background:var(--bg3); border:1px solid var(--border); border-radius:12px; padding:12px 14px; font-size:13px; color:var(--text2); text-align:left; cursor:pointer; transition:all 0.15s; line-height:1.4; }
    .chip:hover { border-color:var(--amber); color:var(--amber); background:var(--amber-bg); }

    .alert { border-radius:12px; padding:12px 14px; font-size:12.5px; line-height:1.5; margin-bottom:10px; display:flex; align-items:flex-start; gap:8px; }
    .alert.warn { background:var(--amber-bg); border:1px solid rgba(232,147,10,0.25); color:var(--amber2); }
    .alert.info { background:var(--green-bg); border:1px solid rgba(45,184,122,0.25); color:var(--green); }

    /* MESSAGES */
    .msg-row { display:flex; gap:8px; align-items:flex-end; }
    .msg-row.user { flex-direction:row-reverse; }
    .msg-avi { width:30px; height:30px; border-radius:10px; display:flex; align-items:center; justify-content:center; font-size:14px; flex-shrink:0; }
    .msg-avi.bot { background:linear-gradient(135deg,var(--amber),var(--amber2)); }
    .msg-avi.user { background:var(--bg3); border:1px solid var(--border); font-family:'Syne',sans-serif; font-size:11px; font-weight:800; color:var(--amber); }
    .bubble { max-width:78%; padding:11px 14px; border-radius:16px; font-size:13.5px; line-height:1.65; }
    .msg-row.user .bubble { background:var(--amber); color:#1A0E00; font-weight:600; border-bottom-right-radius:4px; }
    .msg-row.bot  .bubble { background:var(--bg2); border:1px solid var(--border); color:var(--text); border-bottom-left-radius:4px; }
    .bubble p { margin-bottom:7px; } .bubble p:last-child { margin-bottom:0; }
    .bubble strong { color:var(--amber2); font-weight:700; }
    .bubble ul, .bubble ol { padding-left:16px; margin:4px 0; }
    .bubble li { margin-bottom:3px; }
    .bubble code { background:var(--bg3); padding:1px 5px; border-radius:5px; font-size:12px; font-family:monospace; color:var(--amber2); }
    .dots { display:flex; gap:4px; padding:4px 2px; }
    .dots span { width:6px; height:6px; border-radius:50%; background:var(--amber); animation:bop 1.1s infinite; }
    .dots span:nth-child(2) { animation-delay:.18s; } .dots span:nth-child(3) { animation-delay:.36s; }
    @keyframes bop { 0%,60%,100%{transform:translateY(0);opacity:.4;} 30%{transform:translateY(-5px);opacity:1;} }

    /* INPUT */
    .input-bar { padding:10px 14px 16px; background:var(--bg2); border-top:1px solid var(--border); flex-shrink:0; }
    .input-inner { display:flex; align-items:flex-end; gap:8px; background:var(--bg3); border:1.5px solid var(--border); border-radius:16px; padding:10px 10px 10px 14px; transition:border-color 0.2s; }
    .input-inner:focus-within { border-color:var(--amber); }
    .input-inner.listening { border-color:var(--red); background:rgba(232,90,79,0.06); }
    .chat-ta { flex:1; background:none; border:none; outline:none; font-family:'Nunito',sans-serif; font-size:14px; color:var(--text); resize:none; max-height:100px; line-height:1.5; }
    .chat-ta::placeholder { color:var(--text3); }
    .send-btn { width:36px; height:36px; background:var(--amber); border:none; border-radius:11px; cursor:pointer; color:#000; display:flex; align-items:center; justify-content:center; font-size:16px; transition:all 0.15s; flex-shrink:0; }
    .send-btn:hover { background:var(--amber2); }
    .send-btn:disabled { background:var(--bg3); color:var(--text3); cursor:not-allowed; }

    /* MIC BUTTON */
    .mic-btn { width:36px; height:36px; border:none; border-radius:11px; cursor:pointer; display:flex; align-items:center; justify-content:center; font-size:17px; transition:all 0.15s; flex-shrink:0; background:var(--bg2); border:1px solid var(--border); }
    .mic-btn:hover { background:var(--border); }
    .mic-btn.active { background:var(--red); border-color:var(--red); animation:micPulse 1s ease-in-out infinite; }
    @keyframes micPulse {
      0%,100% { box-shadow:0 0 0 0 rgba(232,90,79,0.5); }
      50%      { box-shadow:0 0 0 8px rgba(232,90,79,0); }
    }

    /* LISTENING OVERLAY */
    .listen-bar { display:flex; align-items:center; justify-content:center; gap:10px; padding:8px 14px; background:var(--red-bg); border:1px solid rgba(232,90,79,0.25); border-radius:10px; margin-bottom:8px; font-size:12.5px; color:var(--red); }
    .listen-waves { display:flex; gap:3px; align-items:center; }
    .listen-waves span { width:3px; border-radius:2px; background:var(--red); animation:wave 0.8s ease-in-out infinite; }
    .listen-waves span:nth-child(1){height:8px;animation-delay:0s;}
    .listen-waves span:nth-child(2){height:14px;animation-delay:0.1s;}
    .listen-waves span:nth-child(3){height:20px;animation-delay:0.2s;}
    .listen-waves span:nth-child(4){height:14px;animation-delay:0.3s;}
    .listen-waves span:nth-child(5){height:8px;animation-delay:0.4s;}
    @keyframes wave { 0%,100%{transform:scaleY(0.5);opacity:0.5;} 50%{transform:scaleY(1);opacity:1;} }

    /* SPEAK BUTTON on bot messages */
    .speak-btn { display:inline-flex; align-items:center; gap:5px; margin-top:8px; padding:4px 10px; background:var(--bg3); border:1px solid var(--border); border-radius:20px; font-size:11px; color:var(--text3); cursor:pointer; transition:all 0.15s; }
    .speak-btn:hover { border-color:var(--amber); color:var(--amber); }
    .speak-btn.speaking { border-color:var(--green); color:var(--green); background:var(--green-bg); }

    /* VOICE MODE full screen */
    .voice-modal { position:absolute; inset:0; background:var(--bg); z-index:50; display:flex; flex-direction:column; align-items:center; justify-content:center; gap:0; animation:fadeIn .2s ease; }
    .voice-orb-wrap { position:relative; width:140px; height:140px; margin-bottom:32px; display:flex; align-items:center; justify-content:center; }
    .voice-orb { width:100px; height:100px; border-radius:50%; background:linear-gradient(135deg,var(--amber),var(--amber2)); display:flex; align-items:center; justify-content:center; font-size:44px; position:relative; z-index:2; transition:transform 0.2s; }
    .voice-orb.recording { animation:orbPulse 1s ease-in-out infinite; }
    .voice-orb.replying  { animation:orbSpin 2s linear infinite; }
    .voice-ring { position:absolute; inset:0; border-radius:50%; border:2px solid var(--amber); opacity:0; animation:ringOut 1.5s ease-out infinite; }
    .voice-ring:nth-child(2){ animation-delay:.5s; }
    @keyframes orbPulse { 0%,100%{transform:scale(1);} 50%{transform:scale(1.08);} }
    @keyframes orbSpin  { to{filter:hue-rotate(40deg);} }
    @keyframes ringOut  { 0%{transform:scale(0.8);opacity:0.6;} 100%{transform:scale(1.8);opacity:0;} }

    .voice-status { font-family:'Syne',sans-serif; font-size:20px; font-weight:700; color:var(--text); margin-bottom:8px; }
    .voice-transcript { font-size:14px; color:var(--text2); text-align:center; max-width:280px; min-height:44px; line-height:1.6; margin-bottom:32px; }
    .voice-close { width:52px; height:52px; border-radius:50%; background:var(--bg3); border:1px solid var(--border); display:flex; align-items:center; justify-content:center; font-size:20px; cursor:pointer; transition:background 0.15s; }
    .voice-close:hover { background:var(--border); }

    /* ── DRAWER (docs + settings shared) ── */
    .drawer-overlay { position:absolute; inset:0; background:rgba(0,0,0,0.6); z-index:20; animation:fadeIn .2s ease; }
    @keyframes fadeIn { from{opacity:0} to{opacity:1} }
    .drawer { position:absolute; bottom:0; left:0; right:0; background:var(--bg2); border-top:1px solid var(--border); border-radius:20px 20px 0 0; z-index:30; display:flex; flex-direction:column; max-height:88%; animation:slideUp .25s ease; }
    @keyframes slideUp { from{transform:translateY(100%);} to{transform:translateY(0);} }
    .drawer-handle { width:36px; height:4px; background:var(--border); border-radius:2px; margin:12px auto 0; flex-shrink:0; }
    .drawer-head { display:flex; align-items:center; justify-content:space-between; padding:14px 18px 10px; border-bottom:1px solid var(--border); flex-shrink:0; }
    .drawer-title { font-family:'Syne',sans-serif; font-size:16px; font-weight:700; }
    .drawer-close { width:32px; height:32px; background:var(--bg3); border:1px solid var(--border); border-radius:9px; cursor:pointer; color:var(--text2); display:flex; align-items:center; justify-content:center; font-size:14px; }
    .drawer-body { flex:1; overflow-y:auto; padding:16px 16px 24px; }

    /* SETTINGS FORM */
    .settings-section { margin-bottom:24px; }
    .settings-label { font-size:10.5px; letter-spacing:1.5px; text-transform:uppercase; color:var(--text3); margin-bottom:10px; }
    .settings-card { background:var(--bg3); border:1px solid var(--border); border-radius:14px; padding:16px; }
    .settings-card p { font-size:13px; color:var(--text2); line-height:1.6; margin-bottom:12px; }
    .settings-card a { color:var(--amber); text-decoration:none; }

    .api-input-wrap { position:relative; margin-bottom:10px; }
    .api-input {
      width:100%; background:var(--bg2); border:1.5px solid var(--border);
      border-radius:10px; padding:11px 44px 11px 14px;
      font-family:'Nunito',sans-serif; font-size:13px; color:var(--text);
      outline:none; transition:border-color 0.2s;
      letter-spacing:0.5px;
    }
    .api-input:focus { border-color:var(--amber); }
    .api-input::placeholder { color:var(--text3); letter-spacing:0; }
    .toggle-vis { position:absolute; right:12px; top:50%; transform:translateY(-50%); background:none; border:none; color:var(--text3); cursor:pointer; font-size:16px; }

    .btn { width:100%; padding:13px; border-radius:11px; border:none; cursor:pointer; font-family:'Syne',sans-serif; font-size:14px; font-weight:700; transition:all 0.15s; }
    .btn-primary { background:var(--amber); color:#000; }
    .btn-primary:hover { background:var(--amber2); }
    .btn-danger  { background:var(--red-bg); color:var(--red); border:1px solid rgba(232,90,79,0.25); margin-top:8px; }
    .btn-danger:hover { background:rgba(232,90,79,0.18); }

    .key-status { display:flex; align-items:center; gap:8px; padding:10px 12px; border-radius:10px; font-size:12.5px; margin-bottom:12px; }
    .key-status.set   { background:var(--green-bg); border:1px solid rgba(45,184,122,0.2); color:var(--green); }
    .key-status.unset { background:var(--red-bg);   border:1px solid rgba(232,90,79,0.2);  color:var(--red); }
    .key-preview { font-family:monospace; font-size:12px; opacity:0.8; }

    /* UPLOAD */
    .upload-zone { border:2px dashed var(--border); border-radius:14px; padding:20px 16px; text-align:center; cursor:pointer; transition:all 0.2s; margin-bottom:16px; }
    .upload-zone:hover, .upload-zone.drag { border-color:var(--amber); background:var(--amber-bg); }
    .uz-icon { font-size:28px; margin-bottom:6px; }
    .uz-text { font-size:12.5px; color:var(--text2); line-height:1.55; }
    .uz-text strong { color:var(--amber2); }
    .uz-note { font-size:11px; color:var(--text3); margin-top:4px; }

    .proc-row { display:flex; align-items:center; gap:8px; background:var(--green-bg); border:1px solid rgba(45,184,122,0.2); border-radius:10px; padding:10px 12px; font-size:12.5px; color:var(--green); margin-bottom:12px; }
    .spin { width:12px; height:12px; border:2px solid var(--green); border-top-color:transparent; border-radius:50%; animation:spin .7s linear infinite; flex-shrink:0; }
    @keyframes spin { to{ transform:rotate(360deg); } }

    .docs-label { font-size:10.5px; letter-spacing:1.5px; text-transform:uppercase; color:var(--text3); margin-bottom:8px; }
    .doc-item { display:flex; align-items:center; gap:10px; padding:11px 12px; background:var(--bg3); border:1px solid var(--border); border-radius:12px; margin-bottom:8px; }
    .doc-ico { width:38px; height:38px; border-radius:10px; display:flex; align-items:center; justify-content:center; font-size:18px; flex-shrink:0; }
    .doc-inf { flex:1; min-width:0; }
    .doc-nm { font-size:13px; font-weight:600; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .doc-mt { font-size:11px; color:var(--text3); margin-top:2px; }
    .doc-del { background:none; border:none; color:var(--red); font-size:18px; cursor:pointer; padding:4px; opacity:0.6; transition:opacity 0.15s; flex-shrink:0; }
    .doc-del:hover { opacity:1; }
    .empty-state { text-align:center; padding:20px 0 8px; color:var(--text3); font-size:13px; line-height:1.6; }
    .storage-bar { background:var(--bg3); border-radius:8px; padding:8px 12px; margin-bottom:14px; display:flex; align-items:center; gap:8px; font-size:11.5px; color:var(--text2); }
    .storage-track { flex:1; height:4px; background:var(--border); border-radius:2px; overflow:hidden; }
    .storage-fill { height:100%; background:var(--amber); border-radius:2px; transition:width 0.4s; }

    /* TOAST */
    .toast { position:absolute; top:70px; left:16px; right:16px; background:var(--bg2); border:1px solid var(--border); border-radius:12px; padding:11px 14px; font-size:13px; color:var(--text); z-index:100; display:flex; align-items:center; gap:8px; box-shadow:0 8px 32px rgba(0,0,0,0.4); animation:slideDown .25s ease; border-left:3px solid var(--amber); }
    .toast.err { border-left-color:var(--red); }
    .toast.ok  { border-left-color:var(--green); }
    @keyframes slideDown { from{transform:translateY(-12px);opacity:0;} to{transform:translateY(0);opacity:1;} }

    /* ONBOARDING */
    .onboard { flex:1; display:flex; flex-direction:column; align-items:center; justify-content:center; padding:28px 24px; }
    .onboard-icon { font-size:48px; margin-bottom:16px; }
    .onboard-title { font-family:'Syne',sans-serif; font-size:22px; font-weight:800; margin-bottom:8px; text-align:center; }
    .onboard-sub { font-size:13.5px; color:var(--text2); line-height:1.65; text-align:center; margin-bottom:28px; max-width:300px; }
    .onboard-form { width:100%; }
    .field-label { font-size:11.5px; color:var(--text2); margin-bottom:6px; letter-spacing:0.5px; }
    .steps { display:flex; flex-direction:column; gap:8px; margin-bottom:20px; }
    .step { display:flex; align-items:flex-start; gap:10px; }
    .step-num { width:22px; height:22px; border-radius:50%; background:var(--amber-bg); border:1px solid rgba(232,147,10,0.3); color:var(--amber); font-size:11px; font-weight:700; display:flex; align-items:center; justify-content:center; flex-shrink:0; margin-top:1px; }
    .step-text { font-size:13px; color:var(--text2); line-height:1.5; }
    .step-text a { color:var(--amber); }
  `}</style>
);

/* ─── helpers ─── */
const TYPES = {
  xlsx:{icon:"📊",bg:"#0d2a1a",label:"Excel"}, xls:{icon:"📊",bg:"#0d2a1a",label:"Excel"},
  csv: {icon:"📋",bg:"#0d2a1a",label:"CSV"},   docx:{icon:"📝",bg:"#0d1e38",label:"Word"},
  doc: {icon:"📝",bg:"#0d1e38",label:"Word"},   pdf: {icon:"📄",bg:"#2a0d0d",label:"PDF"},
  pptx:{icon:"📑",bg:"#2a1a0d",label:"PPT"},    ppt: {icon:"📑",bg:"#2a1a0d",label:"PPT"},
  txt: {icon:"🗒️",bg:"#1a1a1a",label:"Text"},
};
const ext   = n => n.split(".").pop().toLowerCase();
const ftype = n => TYPES[ext(n)] || {icon:"📁",bg:"#1a1a1a",label:"File"};
const fmtSz = b => b<1024?b+"B":b<1e6?(b/1024).toFixed(1)+"KB":(b/1e6).toFixed(1)+"MB";
const maskKey= k => k ? "sk-ant-••••••••" + k.slice(-6) : "";

async function extractText(file) {
  const e = ext(file.name);
  if(e==="txt")  return await file.text();
  if(e==="csv") {
    const t=await file.text();
    const r=Papa.parse(t,{header:true,skipEmptyLines:true});
    const h=r.meta.fields||[];
    return `CSV: ${file.name}\nColumns: ${h.join(", ")}\n\n`+r.data.slice(0,300).map(row=>h.map(k=>`${k}: ${row[k]}`).join(" | ")).join("\n");
  }
  if(e==="xlsx"||e==="xls") {
    const buf=await file.arrayBuffer(); const wb=XLSX.read(buf,{type:"array"});
    let out=`Excel: ${file.name}\n\n`;
    for(const sn of wb.SheetNames) out+=`Sheet: ${sn}\n${XLSX.utils.sheet_to_csv(wb.Sheets[sn]).split("\n").slice(0,150).join("\n")}\n\n`;
    return out;
  }
  if(e==="docx") { const buf=await file.arrayBuffer(); const r=await mammoth.extractRawText({arrayBuffer:buf}); return `Word: ${file.name}\n\n${r.value}`; }
  if(e==="pdf")  return {__pdf:true};
  return `File: ${file.name}`;
}
async function toB64(file){ return new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(r.result.split(",")[1]); r.onerror=rej; r.readAsDataURL(file); }); }

function renderMd(text){
  const lines=text.split("\n"); const els=[]; let i=0;
  while(i<lines.length){
    const l=lines[i];
    if(/^#{1,3} /.test(l)){els.push(<p key={i}><strong>{l.replace(/^#+\s/,"")}</strong></p>);}
    else if(/^[-*] /.test(l)){const it=[];while(i<lines.length&&/^[-*] /.test(lines[i])){it.push(<li key={i}>{iMd(lines[i].replace(/^[-*] /,""))}</li>);i++;}els.push(<ul key={"u"+i}>{it}</ul>);continue;}
    else if(/^\d+\. /.test(l)){const it=[];while(i<lines.length&&/^\d+\. /.test(lines[i])){it.push(<li key={i}>{iMd(lines[i].replace(/^\d+\. /,""))}</li>);i++;}els.push(<ol key={"o"+i}>{it}</ol>);continue;}
    else if(l.trim()){els.push(<p key={i}>{iMd(l)}</p>);}
    i++;
  }
  return els;
}
function iMd(s){return s.split(/(\*\*.*?\*\*|`.*?`)/g).map((p,i)=>{if(p.startsWith("**")&&p.endsWith("**"))return<strong key={i}>{p.slice(2,-2)}</strong>;if(p.startsWith("`")&&p.endsWith("`"))return<code key={i}>{p.slice(1,-1)}</code>;return p;});}

const SUGGESTIONS=[
  {icon:"🗃️",text:"What backend tables are used in the recruitment platform?"},
  {icon:"🔄",text:"Walk me through the end-to-end recruitment process"},
  {icon:"⚙️",text:"What configurations do I need to know about?"},
  {icon:"👤",text:"How is candidate data managed and stored?"},
];

/* ─── STORAGE KEYS ─── */
const IDX_KEY = "sanjog:docs:index";
const DOC_KEY = id => `sanjog:doc:${id}`;
const API_KEY_STORE = "sanjog:apikey";

/* ═══════ MAIN APP ═══════ */
export default function Sanjog() {
  const [apiKey,   setApiKey]   = useState("");
  const [savedKey, setSavedKey] = useState("");
  const [keyInput, setKeyInput] = useState("");
  const [showKey,  setShowKey]  = useState(false);
  const [docs,     setDocs]     = useState([]);
  const [msgs,     setMsgs]     = useState([]);
  const [input,    setInput]    = useState("");
  const [loading,  setLoading]  = useState(false);
  const [drawer,   setDrawer]   = useState(null);
  const [drag,     setDrag]     = useState(false);
  const [proc,     setProc]     = useState(null);
  const [toast,    setToast]    = useState(null);
  const [storePct, setStorePct] = useState(0);
  const [appReady, setAppReady] = useState(false);

  // ── Voice ──
  const [listening,    setListening]    = useState(false);
  const [transcript,   setTranscript]   = useState("");
  const [speakingIdx,  setSpeakingIdx]  = useState(null);  // index of msg being spoken
  const [voiceMode,    setVoiceMode]    = useState(false); // full-screen voice UI
  const [voiceStatus,  setVoiceStatus]  = useState("Tap to speak");
  const recognitionRef = useRef(null);
  const synthRef       = useRef(window.speechSynthesis);
  const speechSupported = !!(window.SpeechRecognition || window.webkitSpeechRecognition);

  const fileRef = useRef(); const endRef = useRef();

  useEffect(()=>{ init(); },[]);
  useEffect(()=>{ endRef.current?.scrollIntoView({behavior:"smooth"}); },[msgs,loading]);

  async function init(){
    try {
      // Load API key
      const kr = await window.storage.get(API_KEY_STORE);
      if(kr){ setSavedKey(kr.value); setApiKey(kr.value); }
      // Load docs
      const ids = await getIndex();
      if(ids.length){
        const loaded=[];
        for(const id of ids){ try{ const r=await window.storage.get(DOC_KEY(id)); if(r) loaded.push(JSON.parse(r.value)); }catch{} }
        setDocs(loaded); calcPct(loaded);
      }
    } catch{}
    setAppReady(true);
  }

  function toast_(msg,type="ok"){ setToast({msg,type}); setTimeout(()=>setToast(null),3000); }

  /* ── API Key ── */
  async function saveApiKey(){
    const k = keyInput.trim();
    if(!k.startsWith("sk-ant-")){ toast_("Key should start with sk-ant-","err"); return; }
    await window.storage.set(API_KEY_STORE, k);
    setSavedKey(k); setApiKey(k);
    toast_("API key saved ✓");
    setDrawer(null);
  }
  async function removeApiKey(){
    await window.storage.delete(API_KEY_STORE);
    setSavedKey(""); setApiKey(""); setKeyInput("");
    toast_("API key removed");
  }

  /* ── Storage helpers ── */
  async function getIndex(){ try{ const r=await window.storage.get(IDX_KEY); return r?JSON.parse(r.value):[]; }catch{return[];} }
  async function saveIndex(ids){ await window.storage.set(IDX_KEY,JSON.stringify(ids)); }
  async function saveDoc(doc){
    const payload={id:doc.id,name:doc.name,size:doc.size,text:doc.text,pdfB64:doc.pdfB64};
    const j=JSON.stringify(payload);
    if(j.length>4.5*1024*1024){ toast_("File too large (>4.5 MB)","err"); return false; }
    await window.storage.set(DOC_KEY(doc.id),j);
    const ids=await getIndex(); if(!ids.includes(doc.id)){ ids.push(doc.id); await saveIndex(ids); }
    return true;
  }
  async function deleteDoc(id){
    await window.storage.delete(DOC_KEY(id));
    const ids=(await getIndex()).filter(x=>x!==id); await saveIndex(ids);
    const updated=docs.filter(d=>d.id!==id); setDocs(updated); calcPct(updated);
    toast_("Document removed");
  }
  function calcPct(dl){ const total=dl.reduce((a,d)=>a+(d.text||"").length+(d.pdfB64||"").length,0); setStorePct(Math.min(100,(total/(90*1024*1024))*100)); }

  /* ── File processing ── */
  async function handleFiles(files){
    for(const file of Array.from(files)){
      const e=ext(file.name); const ok=["txt","csv","xlsx","xls","docx","doc","pdf","pptx","ppt"];
      if(!ok.includes(e)){ toast_(`Skipped – unsupported: ${file.name}`,"err"); continue; }
      setProc(`Reading ${file.name}…`);
      try{
        const extracted=await extractText(file);
        const id=`${Date.now()}-${Math.random().toString(36).slice(2)}`;
        let doc={id,name:file.name,size:file.size,text:null,pdfB64:null};
        if(extracted?.__pdf) doc.pdfB64=await toB64(file); else doc.text=extracted;
        if(await saveDoc(doc)){ setDocs(prev=>{ const n=[...prev,doc]; calcPct(n); return n; }); toast_(`Saved: ${file.name}`); }
      }catch{ toast_(`Error reading ${file.name}`,"err"); }
    }
    setProc(null);
  }

  /* ── Voice Input ── */
  function startListening() {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) { toast_("Speech recognition not supported on this browser","err"); return; }
    stopSpeaking();
    const rec = new SR();
    rec.lang = "en-IN";
    rec.continuous = false;
    rec.interimResults = true;
    rec.onstart = () => { setListening(true); setTranscript(""); setVoiceStatus("Listening…"); };
    rec.onresult = (e) => {
      let interim = "", final = "";
      for (let i = e.resultIndex; i < e.results.length; i++) {
        if (e.results[i].isFinal) final += e.results[i][0].transcript;
        else interim += e.results[i][0].transcript;
      }
      const t = final || interim;
      setTranscript(t);
      setInput(t);
    };
    rec.onend = () => {
      setListening(false);
      setVoiceStatus("Tap to speak");
      // Auto-send if there's a transcript
      setInput(prev => { if (prev.trim()) { setTimeout(() => send(prev.trim()), 100); } return prev; });
    };
    rec.onerror = (e) => {
      setListening(false);
      setVoiceStatus("Tap to speak");
      if (e.error !== "aborted") toast_("Mic error: " + e.error, "err");
    };
    recognitionRef.current = rec;
    rec.start();
  }

  function stopListening() {
    recognitionRef.current?.stop();
    setListening(false);
  }

  function toggleMic() {
    if (listening) stopListening(); else startListening();
  }

  /* ── Voice Output ── */
  function speakText(text, idx) {
    const synth = synthRef.current;
    if (!synth) return;
    if (synth.speaking) { synth.cancel(); if (speakingIdx === idx) { setSpeakingIdx(null); return; } }
    // strip markdown symbols for cleaner speech
    const clean = text.replace(/\*\*/g,"").replace(/[#`*_]/g,"").replace(/\n+/g,". ");
    const utt = new SpeechSynthesisUtterance(clean);
    utt.lang = "en-IN";
    utt.rate = 0.95;
    utt.pitch = 1;
    // prefer a natural-sounding voice
    const voices = synth.getVoices();
    const preferred = voices.find(v => v.lang.startsWith("en") && v.localService) || voices[0];
    if (preferred) utt.voice = preferred;
    utt.onstart = () => setSpeakingIdx(idx);
    utt.onend   = () => setSpeakingIdx(null);
    utt.onerror = () => setSpeakingIdx(null);
    synth.speak(utt);
  }

  function stopSpeaking() {
    synthRef.current?.cancel();
    setSpeakingIdx(null);
  }

  /* ── Send ── */
  async function send(override){
    const q=(override||input).trim(); if(!q||loading) return;
    if(!apiKey){ toast_("Add your API key in Settings ⚙️","err"); return; }
    setInput("");
    setMsgs(prev=>[...prev,{role:"user",text:q}]);
    setLoading(true);
    try{
      const content=[];
      for(const d of docs) if(d.pdfB64) content.push({type:"document",source:{type:"base64",media_type:"application/pdf",data:d.pdfB64},title:d.name});
      const textCtx=docs.filter(d=>d.text).map(d=>`===== ${d.name} =====\n${d.text?.slice(0,18000)}`).join("\n\n");
      content.push({type:"text",text:textCtx?`Reference documents:\n\n${textCtx}\n\n---\nUser question: ${q}`:`No documents uploaded yet.\n\nUser question: ${q}`});
      const res=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json","x-api-key":apiKey,"anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},
        body:JSON.stringify({
          model:"claude-sonnet-4-20250514", max_tokens:1000,
          system:"You are Sanjog — a knowledgeable, friendly HR and recruitment platform advisor. Help the user understand their recruitment platform: processes, configurations, backend tables, workflows. Reference document names when citing. Be concise and practical. Use **bold** for key terms.",
          messages:[{role:"user",content}]
        })
      });
      const data=await res.json();
      if(data.error){ setMsgs(prev=>[...prev,{role:"bot",text:`⚠️ API Error: ${data.error.message}`}]); }
      else {
        const answer = data.content?.[0]?.text||"No response.";
        setMsgs(prev => {
          const next = [...prev, {role:"bot", text:answer}];
          // auto-speak if voice mode is open
          if (voiceMode) {
            setTimeout(() => { setVoiceStatus("Speaking…"); speakText(answer, next.length-1); }, 200);
          }
          return next;
        });
      }
    }catch{ setMsgs(prev=>[...prev,{role:"bot",text:"⚠️ Connection error. Please try again."}]); }
    setLoading(false);
  }

  function onKey(e){ if(e.key==="Enter"&&!e.shiftKey){ e.preventDefault(); send(); } }

  if(!appReady) return (
    <>
      <GS/>
      <div className="shell" style={{alignItems:"center",justifyContent:"center"}}>
        <div style={{fontSize:40,marginBottom:16}}>🧑‍💼</div>
        <div style={{fontFamily:"'Syne',sans-serif",fontSize:18,color:"var(--amber)"}}>Loading Sanjog…</div>
      </div>
    </>
  );

  /* ── ONBOARDING (no key yet) ── */
  if(!savedKey && drawer !== "settings") return (
    <>
      <GS/>
      <div className="shell">
        <div className="topbar">
          <div className="topbar-left">
            <div className="logo-pill">SJ</div>
            <div className="topbar-title">Sanjog</div>
          </div>
        </div>
        <div className="onboard">
          <div className="onboard-icon">🔑</div>
          <div className="onboard-title">Connect Claude AI</div>
          <div className="onboard-sub">Sanjog uses Claude AI to answer questions about your documents. You need a free Anthropic API key — it takes 2 minutes.</div>
          <div className="steps">
            <div className="step"><div className="step-num">1</div><div className="step-text">Go to <a href="https://console.anthropic.com" target="_blank">console.anthropic.com</a> and sign up for free</div></div>
            <div className="step"><div className="step-num">2</div><div className="step-text">Go to <strong style={{color:"var(--text)"}}>API Keys</strong> and create a new key</div></div>
            <div className="step"><div className="step-num">3</div><div className="step-text">Copy the key (starts with <code style={{background:"var(--bg3)",padding:"1px 5px",borderRadius:4,color:"var(--amber2)"}}>sk-ant-</code>) and paste below</div></div>
          </div>
          <div className="onboard-form">
            <div className="field-label">Your Anthropic API Key</div>
            <div className="api-input-wrap" style={{marginBottom:14}}>
              <input
                className="api-input"
                type={showKey?"text":"password"}
                placeholder="sk-ant-api03-..."
                value={keyInput}
                onChange={e=>setKeyInput(e.target.value)}
              />
              <button className="toggle-vis" onClick={()=>setShowKey(v=>!v)}>{showKey?"🙈":"👁️"}</button>
            </div>
            <button className="btn btn-primary" onClick={saveApiKey}>Save & Continue →</button>
            <p style={{fontSize:11.5,color:"var(--text3)",textAlign:"center",marginTop:10,lineHeight:1.5}}>
              Your key is stored only on this device.<br/>Never shared with anyone.
            </p>
          </div>
        </div>
        {toast && <div className={`toast ${toast.type}`}>{toast.msg}</div>}
      </div>
    </>
  );

  /* ── MAIN APP ── */
  return (
    <>
      <GS/>
      <div className="shell">

        {/* TOP BAR */}
        <div className="topbar">
          <div className="topbar-left">
            <div className="logo-pill">SJ</div>
            <div className="topbar-title">Sanjog</div>
          </div>
          <div className="topbar-right">
            <div className="icon-btn" onClick={()=>setDrawer("docs")}>
              📂
              {docs.length>0 && <div className="doc-badge">{docs.length}</div>}
            </div>
            <div className="icon-btn" onClick={()=>{ setKeyInput(savedKey); setDrawer("settings"); }}>
              ⚙️
              <div className="key-dot" style={{background:savedKey?"var(--green)":"var(--red)"}}/>
            </div>
          </div>
        </div>

        {/* TOAST */}
        {toast && <div className={`toast ${toast.type}`}>{toast.msg}</div>}

        {/* CHAT */}
        <div className="view">
          <div className="chat-scroll">
            {msgs.length===0 ? (
              <div className="welcome">
                <div className="welcome-avatar">🧑‍💼</div>
                <div className="welcome-title">Namaste! I'm Sanjog.</div>
                <div className="welcome-sub">
                  {docs.length>0 ? "Your documents are loaded. Ask me anything!" : "Tap 📂 to upload your recruitment platform documents."}
                </div>
                {!savedKey && (
                  <div className="alert warn" style={{marginTop:16,width:"100%"}}>
                    ⚠️ No API key set. Tap ⚙️ to add your key.
                  </div>
                )}
                {docs.length===0 && savedKey && (
                  <div className="alert info" style={{marginTop:16,width:"100%"}}>
                    💡 Documents saved here persist forever — upload once, ask anytime.
                  </div>
                )}
                {docs.length>0 && (
                  <div className="welcome-chips">
                    {SUGGESTIONS.map((s,i)=>(<div key={i} className="chip" onClick={()=>send(s.text)}><span style={{marginRight:8}}>{s.icon}</span>{s.text}</div>))}
                  </div>
                )}
              </div>
            ) : (
              <>
                {msgs.map((m,i)=>(
                  <div key={i} className={`msg-row ${m.role}`}>
                    <div className={`msg-avi ${m.role}`}>{m.role==="bot"?"🧑‍💼":"P"}</div>
                    <div className="bubble">
                      {renderMd(m.text)}
                      {m.role==="bot" && (
                        <div
                          className={`speak-btn ${speakingIdx===i?"speaking":""}`}
                          onClick={()=>speakText(m.text,i)}
                        >
                          {speakingIdx===i ? "■ Stop" : "🔊 Listen"}
                        </div>
                      )}
                    </div>
                  </div>
                ))}
                {loading && <div className="msg-row bot"><div className="msg-avi bot">🧑‍💼</div><div className="bubble"><div className="dots"><span/><span/><span/></div></div></div>}
              </>
            )}
            <div ref={endRef}/>
          </div>
          <div className="input-bar">
            {listening && (
              <div className="listen-bar">
                <div className="listen-waves"><span/><span/><span/><span/><span/></div>
                <span>{transcript || "Listening…"}</span>
              </div>
            )}
            <div className={`input-inner ${listening?"listening":""}`}>
              <textarea className="chat-ta" rows={1} placeholder={listening?"Listening…":docs.length?"Ask about your documents…":"Upload documents first…"} value={input} onChange={e=>setInput(e.target.value)} onKeyDown={onKey}/>
              {speechSupported && (
                <button className={`mic-btn ${listening?"active":""}`} onClick={toggleMic} title={listening?"Stop":"Speak"}>
                  {listening ? "⏹" : "🎙️"}
                </button>
              )}
              <button className="send-btn" onClick={()=>send()} disabled={!input.trim()||loading}>➤</button>
            </div>
            {speechSupported && (
              <div style={{textAlign:"center",marginTop:8}}>
                <span style={{fontSize:11,color:"var(--text3)",cursor:"pointer"}} onClick={()=>{ stopSpeaking(); setVoiceMode(true); setVoiceStatus("Tap to speak"); }}>
                  🎤 Switch to Voice Mode
                </span>
              </div>
            )}
          </div>
        </div>

        {/* VOICE MODE MODAL */}
        {voiceMode && (
          <div className="voice-modal">
            <div className="voice-orb-wrap">
              <div className="voice-ring"/>
              <div className="voice-ring"/>
              <div className={`voice-orb ${listening?"recording":loading?"replying":""}`}>
                {loading ? "⏳" : listening ? "🎙️" : "🧑‍💼"}
              </div>
            </div>
            <div className="voice-status">{loading?"Thinking…":voiceStatus}</div>
            <div className="voice-transcript">{transcript || (msgs.length>0 && msgs[msgs.length-1].role==="bot" ? msgs[msgs.length-1].text.slice(0,120)+"…" : "Tap the orb and ask your question")}</div>
            <div
              style={{width:80,height:80,borderRadius:"50%",background:listening?"var(--red)":"var(--amber)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:32,cursor:"pointer",marginBottom:28,boxShadow:listening?"0 0 0 12px rgba(232,90,79,0.15)":"0 0 0 12px rgba(232,147,10,0.15)",transition:"all 0.2s"}}
              onClick={()=>{ if(loading) return; if(listening) stopListening(); else { setVoiceStatus("Listening…"); startListening(); } }}
            >
              {listening?"⏹":"🎙️"}
            </div>
            <div className="voice-close" onClick={()=>{ stopListening(); stopSpeaking(); setVoiceMode(false); setTranscript(""); }}>✕</div>
          </div>
        )}

        {/* DRAWER */}
        {drawer && (
          <>
            <div className="drawer-overlay" onClick={()=>setDrawer(null)}/>
            <div className="drawer">
              <div className="drawer-handle"/>
              <div className="drawer-head">
                <div className="drawer-title">{drawer==="docs"?"📂 My Documents":"⚙️ Settings"}</div>
                <div className="drawer-close" onClick={()=>setDrawer(null)}>✕</div>
              </div>
              <div className="drawer-body">

                {/* ── DOCS ── */}
                {drawer==="docs" && <>
                  {docs.length>0 && (
                    <div className="storage-bar">
                      <span>Storage</span>
                      <div className="storage-track"><div className="storage-fill" style={{width:`${storePct}%`}}/></div>
                      <span>{docs.length} file{docs.length!==1?"s":""}</span>
                    </div>
                  )}
                  <div className="upload-zone"
                    onClick={()=>fileRef.current.click()}
                    onDragOver={e=>{e.preventDefault();setDrag(true);}}
                    onDragLeave={()=>setDrag(false)}
                    onDrop={e=>{e.preventDefault();setDrag(false);handleFiles(e.dataTransfer.files);}}
                    className={`upload-zone${drag?" drag":""}`}
                  >
                    <div className="uz-icon">⬆️</div>
                    <div className="uz-text"><strong>Tap to upload documents</strong><br/>Excel · Word · PDF · CSV · PPT · TXT</div>
                    <div className="uz-note">Saved permanently · No re-upload needed</div>
                  </div>
                  <input ref={fileRef} type="file" multiple accept=".xlsx,.xls,.csv,.docx,.doc,.pdf,.pptx,.ppt,.txt" style={{display:"none"}} onChange={e=>handleFiles(e.target.files)}/>
                  {proc && <div className="proc-row"><div className="spin"/><span>{proc}</span></div>}
                  {docs.length===0
                    ? <div className="empty-state">No documents yet.<br/>Upload your platform files above.</div>
                    : <>{<div className="docs-label">Saved Documents</div>}{docs.map(d=>{ const ft=ftype(d.name); return (
                        <div key={d.id} className="doc-item">
                          <div className="doc-ico" style={{background:ft.bg}}>{ft.icon}</div>
                          <div className="doc-inf"><div className="doc-nm">{d.name}</div><div className="doc-mt">{ft.label} · {fmtSz(d.size)}</div></div>
                          <button className="doc-del" onClick={()=>deleteDoc(d.id)}>🗑</button>
                        </div>
                      );})}</>
                  }
                </>}

                {/* ── SETTINGS ── */}
                {drawer==="settings" && <>
                  <div className="settings-section">
                    <div className="settings-label">Claude API Key</div>
                    <div className="settings-card">
                      <p>Your key is used to power AI answers. It's stored only on this device and never sent anywhere else. Get your free key at <a href="https://console.anthropic.com" target="_blank">console.anthropic.com</a></p>
                      {savedKey && (
                        <div className="key-status set">
                          ✓ Key connected &nbsp;<span className="key-preview">{maskKey(savedKey)}</span>
                        </div>
                      )}
                      {!savedKey && <div className="key-status unset">✕ No key set</div>}
                      <div className="field-label">Enter / Replace Key</div>
                      <div className="api-input-wrap">
                        <input className="api-input" type={showKey?"text":"password"} placeholder="sk-ant-api03-..." value={keyInput} onChange={e=>setKeyInput(e.target.value)}/>
                        <button className="toggle-vis" onClick={()=>setShowKey(v=>!v)}>{showKey?"🙈":"👁️"}</button>
                      </div>
                      <button className="btn btn-primary" onClick={saveApiKey}>Save Key</button>
                      {savedKey && <button className="btn btn-danger" onClick={removeApiKey}>Remove Key</button>}
                    </div>
                  </div>

                  <div className="settings-section">
                    <div className="settings-label">About</div>
                    <div className="settings-card">
                      <p>
                        <strong style={{color:"var(--text)"}}>Sanjog</strong> — your personal recruitment platform guide.<br/><br/>
                        Documents and your API key are stored locally on this device only. Nothing is sent to any external server except your questions + document text, which go directly to Anthropic's Claude API.
                      </p>
                    </div>
                  </div>
                </>}

              </div>
            </div>
          </>
        )}
      </div>
    </>
  );
}
