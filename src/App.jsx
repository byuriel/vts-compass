import { useState, useRef, useEffect, useCallback } from "react";
import mammoth from "mammoth";

// ─── THEME ────────────────────────────────────────────────────────────────────
const T = {
  bg:        "#0c0a14",
  panel:     "#12101e",
  card:      "#1a1728",
  cardHover: "#1f1d30",
  border:    "#2a2440",
  borderHi:  "#3d3560",
  violet:    "#7c3aed",
  violetMid: "#a855f7",
  violetSoft:"#2d1f5e",
  violetGlow:"#7c3aed22",
  teal:      "#0ea5e9",
  tealSoft:  "#0c2d3d",
  green:     "#10b981",
  greenSoft: "#0a2e22",
  amber:     "#f59e0b",
  amberSoft: "#3a2400",
  red:       "#ef4444",
  redSoft:   "#2e0f0f",
  text:      "#ede9fe",
  sub:       "#a89bc2",
  muted:     "#6b5f8a",
};

// ─── SUPABASE STORAGE ─────────────────────────────────────────────────────────
const SUPA_URL  = "https://xobvmenbenijjspgyzci.supabase.co";
const SUPA_KEY  = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhvYnZtZW5iZW5pampzcGd5emNpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk5OTkxNDcsImV4cCI6MjA5NTU3NTE0N30.K7u2lVHOBa8vK7f4xKHx-LrmjqyJrwTFW8Yhas2TJ5k";
const SUPA_RECORD_ID = "dimari_main"; // one row per user — change if adding more users

const supaHeaders = {
  "Content-Type": "application/json",
  "apikey": SUPA_KEY,
  "Authorization": `Bearer ${SUPA_KEY}`,
  "Prefer": "resolution=merge-duplicates",
};

async function storageSave(data) {
  try {
    await fetch(`${SUPA_URL}/rest/v1/progress`, {
      method: "POST",
      headers: supaHeaders,
      body: JSON.stringify({ id: SUPA_RECORD_ID, data, updated_at: new Date().toISOString() }),
    });
  } catch(_) {
    // fail silently — export/import buttons are the fallback
  }
}

async function storageLoad() {
  try {
    const res = await fetch(
      `${SUPA_URL}/rest/v1/progress?id=eq.${SUPA_RECORD_ID}&select=data`,
      { headers: supaHeaders }
    );
    const rows = await res.json();
    return rows?.[0]?.data || null;
  } catch(_) { return null; }
}

// ─── API ──────────────────────────────────────────────────────────────────────
const ANTHROPIC_KEY = ["sk-ant-api03-","2kwDLA7db-_NfmgkbhL6HaTG7nzra4JcojlAcC1Oe62nCHNdDLakHlRx1bwoWtoLBVi6","g1OqN2C85EbqOf12Dg-6PUJ_AAA"].join("");

async function callClaude(system, user, maxTokens=1400) {
  try {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": ANTHROPIC_KEY,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify({ model:"claude-haiku-4-5-20251001", max_tokens:maxTokens, system, messages:[{role:"user",content:user}] })
    });
    const d = await res.json();
    if (d.error) throw new Error(`API error: ${d.error.message} (type: ${d.error.type})`);
    if (!d.content) throw new Error(`Unexpected response: ${JSON.stringify(d).slice(0,200)}`);
    return d.content.map(b=>b.text||"").join("") || "No response content";
  } catch(err) {
    if (err.message.includes("fetch")) throw new Error("Network error — check internet connection and try again");
    throw err;
  }
}

// ─── JSZIP LOADER ─────────────────────────────────────────────────────────────
// dynamic import() is blocked by sandbox CSP — load via script tag instead
let _jszipReady = null;
function loadJSZip() {
  if (_jszipReady) return _jszipReady;
  _jszipReady = new Promise((resolve, reject) => {
    if (window.JSZip) { resolve(window.JSZip); return; }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
    s.onload  = () => window.JSZip ? resolve(window.JSZip) : reject(new Error("JSZip not on window after load"));
    s.onerror = () => reject(new Error("JSZip script failed to load"));
    document.head.appendChild(s);
  });
  return _jszipReady;
}

// ─── DOCX XML PARSER ──────────────────────────────────────────────────────────
// Parse raw OOXML to extract font, size, spacing, margin checks
async function parseDocxXml(arrayBuffer) {
  const results = [];
  try {
    const JSZip = await loadJSZip();
    const zip   = await JSZip.loadAsync(arrayBuffer);

    // ── document.xml — font, spacing, margins ────────────────────────────────
    const docXml = await zip.file("word/document.xml")?.async("string");
    if (docXml) {
      // Font names (filter out Symbol/Wingdings which appear in bullet runs)
      const allFonts = [...docXml.matchAll(/w:ascii="([^"]+)"/g)]
        .map(m => m[1])
        .filter(f => !/(Symbol|Wingdings|Webdings)/i.test(f));
      const uniqueFonts = [...new Set(allFonts)];
      const hasTNR    = uniqueFonts.some(f => /times\s*new\s*roman|times/i.test(f));
      const hasOther  = uniqueFonts.some(f => !/times/i.test(f));
      results.push({
        label:"Font",
        value: uniqueFonts.length ? uniqueFonts.join(", ") : "Not detected in XML",
        ok: uniqueFonts.length === 0 ? null : (hasTNR && !hasOther),
        expected:"Times New Roman only",
      });

      // Line spacing — w:line in twips (1.5 spacing = 360, single = 240, double = 480)
      const spacingNums = [...docXml.matchAll(/w:line="(\d+)"/g)].map(m => parseInt(m[1]));
      if (spacingNums.length) {
        const freq = {};
        spacingNums.forEach(v => { freq[v] = (freq[v]||0)+1; });
        const dominant = parseInt(Object.entries(freq).sort((a,b)=>b[1]-a[1])[0][0]);
        const spacingLabel =
          dominant < 280 ? "Single spacing" :
          dominant < 320 ? "~1.15 spacing" :
          dominant < 420 ? "1.5 spacing ✓" :
          dominant < 520 ? "~Double spacing" : "Custom spacing";
        results.push({
          label:"Line Spacing",
          value:`${dominant} twips — ${spacingLabel}`,
          ok: dominant >= 340 && dominant <= 400,
          expected:"1.5 lines (360 twips)",
        });
      } else {
        results.push({ label:"Line Spacing", value:"Not found in XML — check manually", ok:null, expected:"1.5 lines" });
      }

      // Page margins — in sectPr, w:pgMar twips (720 = 0.5 inch, 1440 = 1 inch)
      const pgMarMatch = docXml.match(/w:pgMar[^>]*w:top="(\d+)"[^>]*w:right="(\d+)"[^>]*w:bottom="(\d+)"[^>]*w:left="(\d+)"/);
      if (pgMarMatch) {
        const [, top, right, bottom, left] = pgMarMatch.map(Number);
        const toIn = v => (v / 1440).toFixed(2);
        const allHalf = [top, right, bottom, left].every(v => v >= 680 && v <= 760);
        results.push({
          label:"Margins",
          value:`T:${toIn(top)}" R:${toIn(right)}" B:${toIn(bottom)}" L:${toIn(left)}"`,
          ok: allHalf,
          expected:'0.5" all sides (720 twips)',
        });
      } else {
        results.push({ label:"Margins", value:"Not found in XML — check manually", ok:null, expected:'0.5" all sides' });
      }
    }

    // ── styles.xml — default font size ───────────────────────────────────────
    const stylesXml = await zip.file("word/styles.xml")?.async("string");
    if (stylesXml) {
      // w:sz = half-points: 10.5pt = 21, 11pt = 22, 12pt = 24
      const allSizes = [...stylesXml.matchAll(/w:sz\s+w:val="(\d+)"/g)].map(m => parseInt(m[1]));
      if (allSizes.length) {
        const freq = {};
        allSizes.forEach(v => { freq[v] = (freq[v]||0)+1; });
        const dominant = parseInt(Object.entries(freq).sort((a,b)=>b[1]-a[1])[0][0]);
        results.push({
          label:"Font Size",
          value:`${dominant / 2}pt (most common in styles)`,
          ok: dominant >= 20 && dominant <= 22,   // 10pt–11pt range, 10.5 = 21
          expected:"10.5pt (21 half-points)",
        });
      }
    }

  } catch(err) {
    results.push({ label:"XML Parse Error", value:err.message, ok:false, expected:"Valid .docx" });
  }
  return results;
}

// ─── DOCX FORMAT CHECK ────────────────────────────────────────────────────────
async function checkDocxFormatting(originalBuffer) {
  const passes = [];
  try {
    // Always clone — mammoth and JSZip each consume the buffer
    const bufForMammoth = originalBuffer.slice(0);
    const bufForXml     = originalBuffer.slice(0);

    // Text extraction via mammoth
    const rawResult  = await mammoth.extractRawText({ arrayBuffer: bufForMammoth });
    const text       = rawResult.value || "";
    const wordCount  = text.split(/\s+/).filter(Boolean).length;
    const estPages   = (wordCount / 230).toFixed(1);

    passes.push({ label:"File Format",     value:".docx ✓", ok:true,                          expected:".docx required" });
    passes.push({ label:"Est. Page Count", value:`~${estPages} pages (${wordCount} words)`,
                  ok: parseFloat(estPages) <= 5.5, expected:"Max 5 pages" });

    // Deep XML checks via JSZip
    const xmlResults = await parseDocxXml(bufForXml);
    passes.push(...xmlResults);

    // Drug notation check from extracted text
    const mlOnly = (text.match(/\b\d+\.?\d*\s*m[Ll]\b/g) || []).length;
    const mgkg   = (text.match(/\d+\.?\d*\s*(?:mg|mcg)\/kg/gi) || []).length;
    passes.push({
      label:"Drug Notation",
      value: mlOnly > 0 ? `${mlOnly} bare mL reference(s) — verify mg/kg also present` : "No bare mL doses detected",
      ok: mlOnly === 0 || mgkg > mlOnly,
      expected:"mg or mcg/kg — not mL alone",
    });

    // Blinding check — scan for facility names that should not appear in Folder 2
    const clinicPattern = /animal\s+hospital|vet(?:erinary)?\s+(?:center|clinic|hospital)|university\s+of|college\s+of\s+vet/i;
    const hasClinic = clinicPattern.test(text);
    passes.push({
      label:"Blinding Check",
      value: hasClinic ? "⚠ Possible facility name detected — review Folder 2" : "No facility names detected",
      ok: !hasClinic,
      expected:"No identifying info in Folder 2",
    });

    // Abbreviation check — common ones that should be spelled out on first use
    const abbrevs = ["CRT","PCV","TP","ETT","pRBC","CRI","IPPV","NSR","ICU","PE","HR","RR","MAP","SpO2","EtCO2"];
    const undefined_abbrevs = abbrevs.filter(a => {
      const firstUse = text.indexOf(a);
      if (firstUse === -1) return false;
      // Check if a definition appears within 150 chars before first use
      const context = text.slice(Math.max(0, firstUse - 150), firstUse);
      return !context.includes("("+a+")") && !context.includes(a+"(");
    });
    passes.push({
      label:"Abbreviations",
      value: undefined_abbrevs.length > 0
        ? `${undefined_abbrevs.length} possibly undefined: ${undefined_abbrevs.slice(0,5).join(", ")}${undefined_abbrevs.length>5?" +more":""}`
        : "No unannounced abbreviations detected",
      ok: undefined_abbrevs.length === 0,
      expected:"All abbreviations spelled out on first use",
    });

    return { passes, text, wordCount, estPages };
  } catch(err) {
    return { passes:[{ label:"Parse Error", value:err.message, ok:false, expected:"Valid .docx file" }], text:"", wordCount:0, estPages:0 };
  }
}

// ─── SHARED STYLES ────────────────────────────────────────────────────────────
const S = {
  card: { background:T.card, border:`1px solid ${T.border}`, borderRadius:16, padding:24, marginBottom:20 },
  btn: (v="primary") => ({
    padding:"10px 20px", borderRadius:10, border:"none", cursor:"pointer", fontSize:13, fontWeight:700,
    background: v==="primary" ? `linear-gradient(135deg,${T.violet},${T.violetMid})` : v==="teal" ? `linear-gradient(135deg,${T.teal},#0284c7)` : v==="ghost" ? "transparent" : T.violetSoft,
    color: v==="ghost" ? T.muted : "#fff",
    border: v==="ghost" ? `1px solid ${T.border}` : "none",
    transition:"all 0.15s", letterSpacing:0.2,
  }),
  chip: (c="violet") => ({
    display:"inline-block", padding:"3px 10px", borderRadius:20, fontSize:11, fontWeight:700,
    background: c==="green"?T.greenSoft:c==="red"?T.redSoft:c==="amber"?T.amberSoft:c==="teal"?T.tealSoft:T.violetSoft,
    color: c==="green"?T.green:c==="red"?T.red:c==="amber"?T.amber:c==="teal"?T.teal:T.violetMid,
    marginRight:6, marginBottom:4,
  }),
  scoreBar: { height:6, borderRadius:4, background:T.border, overflow:"hidden", margin:"6px 0 2px" },
  scoreFill: (pct,c) => ({ height:"100%", width:`${pct}%`, borderRadius:4, transition:"width 0.6s ease",
    background: c==="green"?T.green:c==="amber"?T.amber:c==="red"?T.red:c==="teal"?T.teal:`linear-gradient(90deg,${T.violet},${T.violetMid})` }),
  aiBlock: { background:"#100e1a", border:`1px solid ${T.violetSoft}`, borderLeft:`3px solid ${T.violet}`, borderRadius:"0 10px 10px 0", padding:18, marginTop:16, fontSize:13, lineHeight:1.75, color:T.sub, whiteSpace:"pre-wrap" },
  simpleBlock: { background:"#0f1e18", border:`1px solid #1a3d2e`, borderLeft:`3px solid ${T.green}`, borderRadius:"0 10px 10px 0", padding:18, marginTop:12, fontSize:13, lineHeight:1.75, color:"#86efac", whiteSpace:"pre-wrap" },
  spinner: { display:"inline-block",width:14,height:14,border:"2px solid rgba(255,255,255,0.15)",borderTop:"2px solid #fff",borderRadius:"50%",animation:"spin 0.7s linear infinite",marginRight:8,verticalAlign:"middle" },
  input: { background:"#0e0c18", border:`1px solid ${T.border}`, borderRadius:8, padding:"10px 14px", color:T.text, fontSize:13, outline:"none", width:"100%", boxSizing:"border-box" },
  label: { fontSize:11, color:T.muted, fontWeight:700, marginBottom:6, display:"block", letterSpacing:0.5, textTransform:"uppercase" },
  textarea: { width:"100%", background:"#0e0c18", border:`1px solid ${T.border}`, borderRadius:10, padding:"12px 14px", color:T.text, fontSize:13, fontFamily:"inherit", resize:"vertical", outline:"none", lineHeight:1.65, boxSizing:"border-box" },
  sectionTitle: { fontSize:22, fontWeight:800, color:T.text, marginBottom:8, letterSpacing:-0.5 },
  sectionSub: { fontSize:14, color:T.muted, marginBottom:28 },
  grid2: { display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(230px,1fr))", gap:14, marginBottom:20 },
  statCard: (c="violet") => ({
    background:T.card, borderRadius:12, padding:18,
    border:`1px solid ${c==="green"?T.green+"44":c==="red"?T.red+"44":c==="amber"?T.amber+"44":c==="teal"?T.teal+"44":T.violet+"44"}`,
  }),
  dropzone: (drag) => ({ border:`2px dashed ${drag?T.violet:T.border}`, borderRadius:12, padding:"32px 20px", textAlign:"center", cursor:"pointer", transition:"all 0.2s", background:drag?T.violetGlow:"transparent" }),
  checkRow: { display:"flex",alignItems:"flex-start",gap:12,padding:"10px 0",borderBottom:`1px solid ${T.border}`,cursor:"pointer" },
  checkbox: (ch) => ({ width:18,height:18,borderRadius:4,border:`2px solid ${ch?T.green:T.border}`,background:ch?T.green:"transparent",display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",flexShrink:0,marginTop:2,fontSize:10,color:"#fff",transition:"all 0.15s" }),
};

// ─── PLAIN ENGLISH HELPER ─────────────────────────────────────────────────────
async function simplify(text) {
  return await callClaude(
    "You simplify veterinary credential application feedback into plain, friendly English. No jargon. Use short sentences. Be encouraging but honest. Keep it under 300 words. Use bullet points for action items.",
    `Simplify this AVTAA feedback into plain English a non-expert can easily understand:\n\n${text}`
  );
}

// ─── SAVE INDICATOR ───────────────────────────────────────────────────────────
function SaveIndicator({saved}) {
  return (
    <div style={{position:"fixed",bottom:20,right:20,zIndex:999,padding:"8px 14px",borderRadius:20,background:saved?T.greenSoft:T.violetSoft,border:`1px solid ${saved?T.green:T.violet}`,fontSize:12,color:saved?T.green:T.violetMid,fontWeight:700,transition:"all 0.4s",display:"flex",alignItems:"center",gap:6}}>
      {saved ? <>☁️ Saved to cloud</> : <><span style={{display:"inline-block",width:10,height:10,border:`2px solid ${T.violetMid}`,borderTop:`2px solid transparent`,borderRadius:"50%",animation:"spin 0.7s linear infinite"}}/>Syncing...</>}
    </div>
  );
}

// ─── DROPZONE ─────────────────────────────────────────────────────────────────
function DropZone({accept,label,sublabel,onFile,fileName}) {
  const [drag,setDrag]=useState(false);
  const ref=useRef();
  return (
    <div style={S.dropzone(drag)} onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
      onDrop={e=>{e.preventDefault();setDrag(false);if(e.dataTransfer.files[0])onFile(e.dataTransfer.files[0]);}} onClick={()=>ref.current.click()}>
      <input ref={ref} type="file" accept={accept} style={{display:"none"}} onChange={e=>e.target.files[0]&&onFile(e.target.files[0])}/>
      {fileName
        ? <div><div style={{fontSize:26,marginBottom:6}}>📄</div><div style={{fontWeight:700,color:T.green,fontSize:14}}>{fileName}</div><div style={{fontSize:11,color:T.muted,marginTop:4}}>Click to replace</div></div>
        : <div><div style={{fontSize:32,marginBottom:8}}>⬆️</div><div style={{fontWeight:700,color:T.text,fontSize:14,marginBottom:4}}>{label}</div><div style={{fontSize:12,color:T.muted}}>{sublabel}</div></div>}
    </div>
  );
}

// ─── FORMAT RESULTS ───────────────────────────────────────────────────────────
function FormatResults({results}) {
  if(!results?.length)return null;
  return (
    <div style={{marginTop:14}}>
      <div style={{fontSize:11,color:T.muted,fontWeight:700,marginBottom:8,letterSpacing:0.5}}>FORMATTING COMPLIANCE</div>
      <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
        {results.map((r,i)=>(
          <div key={i} style={{background:T.panel,border:`1px solid ${r.ok===true?T.green+"55":r.ok===false?T.red+"55":T.border}`,borderRadius:10,padding:"8px 12px",minWidth:140}}>
            <div style={{fontSize:10,color:T.muted,fontWeight:700,marginBottom:2,letterSpacing:0.3}}>{r.label}</div>
            <div style={{fontSize:12,fontWeight:700,color:r.ok===true?T.green:r.ok===false?T.red:T.amber}}>{r.ok===true?"✓":r.ok===false?"✗":"?"} {r.value}</div>
            <div style={{fontSize:10,color:T.muted,marginTop:2}}>{r.expected}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── AI RESULT WITH PLAIN ENGLISH TOGGLE ─────────────────────────────────────
function AIResult({result, blockStyle}) {
  const [simple,setSimple]=useState("");
  const [loading,setLoading]=useState(false);
  const [showSimple,setShowSimple]=useState(false);
  const handleSimplify=async()=>{
    if(simple){setShowSimple(s=>!s);return;}
    setLoading(true);setShowSimple(true);
    const r=await simplify(result);
    setSimple(r);setLoading(false);
  };
  return (
    <div>
      <div style={blockStyle||S.aiBlock}>
        <div style={{fontSize:10,color:T.violet,fontWeight:700,marginBottom:8,letterSpacing:1}}>AVTAA COMMITTEE SIMULATION</div>
        {result}
      </div>
      <div style={{marginTop:8,display:"flex",gap:8,alignItems:"center"}}>
        <button style={{...S.btn("ghost"),fontSize:12,padding:"6px 14px"}} onClick={handleSimplify} disabled={loading}>
          {loading?<><span style={S.spinner}/>Simplifying...</>:showSimple?"🔬 Show Technical Version":"💬 Explain in Plain English"}
        </button>
      </div>
      {showSimple&&simple&&(
        <div style={S.simpleBlock}>
          <div style={{fontSize:10,color:T.green,fontWeight:700,marginBottom:8,letterSpacing:1}}>PLAIN ENGLISH SUMMARY</div>
          {simple}
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// TAB 1 — CHECKLIST
// ═══════════════════════════════════════════════════════════════════════════════
const CHECKLIST_ITEMS = {
  "Pre-Application (Already Done ✓)":[
    {id:"pa1",text:"Professional History & Experience form submitted"},
    {id:"pa2",text:"Current in-date license submitted"},
    {id:"pa3",text:"Letter of Good Standing from vet medical board"},
    {id:"pa4",text:"$60.00 application fee paid via PayPal"},
    {id:"pa5",text:"Letter of Agreement signed (DACVAA or VTS)"},
    {id:"pa6",text:"Pre-application approval email + applicant number received"},
  ],
  "Folder 1 — Your Identity Documents (Due Dec 31)":[
    {id:"f1a",text:"Application Waiver & Release signed → [firstname.lastname.waiver.pdf]"},
    {id:"f1b",text:"Plagiarism Affidavit signed → [firstname.lastname.affidavit.pdf]"},
    {id:"f1c",text:"Letter of Recommendation #1 from DACVAA or VTS — signed"},
    {id:"f1d",text:"Letter of Recommendation #2 signed"},
    {id:"f1e",text:'Statement of Purpose: 1 page, 12pt Times New Roman, 1" margins, signed'},
    {id:"f1f",text:"40+ CE hours — each conference form bundled WITH its attendance proof"},
    {id:"f1g",text:"All CE speaker credentials listed on form (DACVAA, VTS, etc.)"},
    {id:"f1h",text:"Skills Verification Form signed by DACVAA or VTS for each work location"},
  ],
  "Folder 2 — Blinded Documents (Applicant # Only, Due Dec 31)":[
    {id:"f2a",text:"Employment Location Form → [applicant#.location]"},
    {id:"f2b",text:"Case Logs: 50–60 cases, strongly recommend all 60 → [applicant#.caselog.pdf]"},
    {id:"f2c",text:"ASA I/II cases: maximum 12, within first 12 logs (mark SKILLS ONLY if beyond)"},
    {id:"f2d",text:"Sedation-only cases: maximum 3 (small animal track)"},
    {id:"f2e",text:"All drug doses written as mg or mcg/kg — never just mL"},
    {id:"f2f",text:"Location dropdown filled in on every single case log entry"},
    {id:"f2g",text:"Skills list submitted (core 90%, supplemental 50%) → [applicant#.skills.pdf]"},
    {id:"f2h",text:"Each mastered skill has one specific case log number assigned"},
    {id:"f2i",text:"Case Reports 1–4 uploaded as .docx files → [applicant#.casereport1-4.docx]"},
    {id:"f2j",text:'Reports: 10.5pt Times New Roman, 1.5 line spacing, 0.5" margins, max 5 pages'},
    {id:"f2k",text:"Zero identifying info anywhere in Folder 2 (no names, no clinic names)"},
    {id:"f2l",text:"Anesthesia Records 1–4 uploaded, one per case report, blinded"},
  ],
  "Final Submission":[
    {id:"s1",text:"Folder 1 compressed → [firstname.lastname.AVTAA2025.zip]"},
    {id:"s2",text:"Folder 2 compressed → [applicant#.AVTAA2025.zip]"},
    {id:"s3",text:"Both zipped folders uploaded TOGETHER via WeTransfer"},
    {id:"s4",text:"Sent to: avtaa.vts.credential@gmail.com — NOT the pre-app email"},
    {id:"s5",text:"Both WeTransfer confirmation emails saved with timestamps"},
    {id:"s6",text:"Submitted before Dec 31, 11:59:59 PM Eastern — zero exceptions"},
  ],
};

function ChecklistTab({checked,setChecked}) {
  const allIds=Object.values(CHECKLIST_ITEMS).flat().map(i=>i.id);
  const total=allIds.length,done=Object.values(checked).filter(Boolean).length;
  const pct=Math.round(done/total*100);
  const bc=pct>=80?"green":pct>=50?"amber":"red";
  const toggle=id=>setChecked(p=>({...p,[id]:!p[id]}));
  return (
    <div>
      <div style={S.sectionTitle}>Application Checklist</div>
      <div style={S.sectionSub}>Every required document. Pre-app items are already checked. Progress saves automatically.</div>
      <div style={{...S.card,marginBottom:28}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <div style={{fontWeight:700,fontSize:15}}>Overall Completion</div>
          <div style={{fontSize:24,fontWeight:900,color:bc==="green"?T.green:bc==="amber"?T.amber:T.red}}>{pct}%</div>
        </div>
        <div style={S.scoreBar}><div style={S.scoreFill(pct,bc)}/></div>
        <div style={{fontSize:12,color:T.muted,marginTop:6}}>{done} of {total} items complete</div>
      </div>
      {Object.entries(CHECKLIST_ITEMS).map(([section,items])=>{
        const sd=items.filter(i=>checked[i.id]).length;
        return (
          <div key={section} style={{...S.card,marginBottom:14}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
              <div style={{fontWeight:700,fontSize:14,color:T.text}}>{section}</div>
              <span style={S.chip(sd===items.length?"green":"violet")}>{sd}/{items.length}</span>
            </div>
            {items.map(item=>(
              <div key={item.id} style={S.checkRow} onClick={()=>toggle(item.id)}>
                <div style={S.checkbox(checked[item.id])}>{checked[item.id]?"✓":""}</div>
                <div style={{fontSize:13,color:checked[item.id]?T.muted:T.sub,textDecoration:checked[item.id]?"line-through":"none",lineHeight:1.5}}>{item.text}</div>
              </div>
            ))}
          </div>
        );
      })}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// TAB 2 — READINESS SCORE
// ═══════════════════════════════════════════════════════════════════════════════
function ReadinessTab({checked,caseLogCount,caseLogRecords,asaCounts,sedationOnly,reportScores,reportUploadCount,ceHours,coreSkillsPct,suppSkillsPct}) {
  const lowASA=(asaCounts.I||0)+(asaCounts.II||0);
  const allIds=Object.values(CHECKLIST_ITEMS).flat().map(i=>i.id);
  const checklistPct=Math.round(Object.values(checked).filter(Boolean).length/allIds.length*100);
  // Use caseLogRecords.length as source of truth if records exist, else fall back to manual count
  const effectiveLogCount = caseLogRecords.length > 0 ? caseLogRecords.length : caseLogCount;
  // Linear: each case = 1.67 points. 50 = 83%, 60 = 100%. No jump from 1 log = 40%
  const logScore = Math.min(100, Math.round(effectiveLogCount / 60 * 100));
  const asaScore=lowASA>0&&lowASA<=12?100:lowASA===0?60:0;
  const sedScore=sedationOnly<=3?100:0;
  // Case reports: blend upload progress (25% each) + scored quality
  // 4 uploaded = 40% base, then scored reports push toward 100%
  const scoredReports = reportScores.filter(s=>s!==null);
  const uploadPct = Math.round((reportUploadCount / 4) * 40); // max 40 pts for uploading
  const scorePct = scoredReports.length > 0
    ? Math.round(scoredReports.reduce((a,b)=>a+b,0) / scoredReports.length * 0.60) // max 60 pts for score quality
    : 0;
  const avgReport = Math.min(100, uploadPct + scorePct);
  const ceScore=ceHours>=40?100:ceHours>0?Math.round(ceHours/40*100):0;
  const coreScore=coreSkillsPct>=90?100:Math.round(coreSkillsPct/90*100);
  const suppScore=suppSkillsPct>=50?100:Math.round(suppSkillsPct/50*100);
  const overall=Math.round((checklistPct*0.15+logScore*0.15+asaScore*0.05+sedScore*0.05+avgReport*0.30+ceScore*0.15+coreScore*0.10+suppScore*0.05));
  const readinessColor=overall>=80?"green":overall>=60?"amber":"red";
  const readinessLabel=overall>=80?"Ready to Submit":overall>=60?"Almost Ready":"Needs More Work";
  const todos=[];
  if(effectiveLogCount<50)todos.push({priority:"🔴 Critical",text:`Add ${50-effectiveLogCount} more case logs (need minimum 50, recommend 60)`});
  if(lowASA>12)todos.push({priority:"🔴 Critical",text:`Remove ${lowASA-12} ASA I/II cases — only 12 allowed`});
  if(sedationOnly>3)todos.push({priority:"🔴 Critical",text:`Reduce sedation-only cases to 3 maximum`});
  if(scoredReports.length>0&&scoredReports.reduce((a,b)=>a+b,0)/scoredReports.length<60)todos.push({priority:"🔴 Critical",text:`Improve case reports — avg score ${Math.round(scoredReports.reduce((a,b)=>a+b,0)/scoredReports.length)}% (aim for 80%+)`});
  if(reportUploadCount<4)todos.push({priority:"🔴 Critical",text:`Upload ${4-reportUploadCount} more case report${4-reportUploadCount!==1?"s":""} in the Case Reports tab`});
  if(reportUploadCount===4&&scoredReports.length<4)todos.push({priority:"🟡 Important",text:`Score ${4-scoredReports.length} remaining case report${4-scoredReports.length!==1?"s":""} — click Analyze & Score`});
  if(ceHours<40)todos.push({priority:"🟡 Important",text:`Add ${40-ceHours} more CE hours (need 40 minimum)`});
  if(coreSkillsPct<90)todos.push({priority:"🟡 Important",text:`Document ${Math.ceil(32*0.9)-Math.round(32*coreSkillsPct/100)} more core skills in case logs`});
  if(suppSkillsPct<50)todos.push({priority:"🟡 Important",text:`Document more supplemental skills — currently at ${suppSkillsPct}%`});
  if(checklistPct<100)todos.push({priority:"🔵 To Do",text:`Complete remaining checklist items (${checklistPct}% done)`});
  if(todos.length===0)todos.push({priority:"✅ All Clear",text:"All tracked metrics look good — final review recommended before submission"});

  const metrics=[
    {label:"Checklist",pct:checklistPct,color:"violet"},
    {label:`Case Logs (${effectiveLogCount}/60)`,pct:logScore,color:logScore>=100?"green":logScore>=83?"amber":"red"},
    {label:"ASA Distribution",pct:asaScore,color:asaScore===100?"green":"red"},
    {label:`Case Reports (${reportUploadCount}/4 uploaded, ${scoredReports.length}/4 scored)`,pct:avgReport,color:avgReport>=80?"green":avgReport>=40?"amber":"red"},
    {label:"CE Hours",pct:ceScore,color:ceScore===100?"green":ceScore>0?"amber":"red"},
    {label:"Core Skills",pct:coreScore,color:coreScore>=100?"green":coreScore>=70?"amber":"red"},
    {label:"Suppl. Skills",pct:suppScore,color:suppScore>=100?"green":suppScore>=70?"amber":"red"},
  ];

  return (
    <div>
      <div style={S.sectionTitle}>Readiness Score</div>
      <div style={S.sectionSub}>One number that pulls everything together. Updates automatically as you work through the other tabs.</div>
      <div style={{...S.card,textAlign:"center",padding:"36px 24px",marginBottom:24,background:`linear-gradient(135deg,${T.panel},${T.violetSoft}22)`,border:`1px solid ${T.violet}44`}}>
        <div style={{fontSize:80,fontWeight:900,lineHeight:1,color:readinessColor==="green"?T.green:readinessColor==="amber"?T.amber:T.red,textShadow:`0 0 40px ${readinessColor==="green"?T.green:readinessColor==="amber"?T.amber:T.red}44`}}>{overall}%</div>
        <div style={{fontSize:18,fontWeight:700,color:T.sub,marginTop:8}}>{readinessLabel}</div>
        <div style={{fontSize:13,color:T.muted,marginTop:4}}>Overall application readiness</div>
      </div>
      <div style={S.grid2}>
        {metrics.map(m=>(
          <div key={m.label} style={S.statCard(m.color)}>
            <div style={{fontSize:11,color:T.muted,fontWeight:700,marginBottom:4,letterSpacing:0.3}}>{m.label}</div>
            <div style={{fontSize:22,fontWeight:900,color:m.color==="green"?T.green:m.color==="amber"?T.amber:m.color==="red"?T.red:T.violetMid}}>{m.pct}%</div>
            <div style={S.scoreBar}><div style={S.scoreFill(m.pct,m.color)}/></div>
          </div>
        ))}
      </div>
      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:15,marginBottom:16}}>Prioritized To-Do List</div>
        {todos.map((t,i)=>(
          <div key={i} style={{display:"flex",gap:12,padding:"10px 0",borderBottom:`1px solid ${T.border}`}}>
            <div style={{fontSize:13,fontWeight:700,minWidth:110,flexShrink:0}}>{t.priority}</div>
            <div style={{fontSize:13,color:T.sub,lineHeight:1.5}}>{t.text}</div>
          </div>
        ))}
      </div>
      <div style={{...S.card,background:T.violetSoft,border:`1px solid ${T.violet}44`}}>
        <div style={{fontSize:13,color:T.violetMid,fontWeight:700,marginBottom:6}}>💡 How to improve your score fast</div>
        <div style={{fontSize:12,color:T.sub,lineHeight:1.8}}>Case Reports carry the most weight (30%). Upload and score all 4 in the Case Reports tab — each point of improvement there moves the needle the most. Case logs and CE are next.</div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// TAB 3 — REJECTION ANALYSIS
// ═══════════════════════════════════════════════════════════════════════════════
function buildRejectionSystem() {
  return `You are a senior AVTAA VTS (Anesthesia & Analgesia) application specialist. You have reviewed hundreds of failed and successful applications and know exactly what the credentials committee looks for.

A veterinary technician is on her SECOND attempt after being rejected. You are analyzing her rejection materials to build a comprehensive correction plan AND a personalized "mistake memory" she can apply to every future case log and case report she writes.

YOUR ANALYSIS MUST BE EXHAUSTIVE. Read every word of the rejection materials carefully. Do not summarize or skip any deficiency — every single issue mentioned must be addressed.

STRUCTURE YOUR RESPONSE EXACTLY AS FOLLOWS:

═══════════════════════════════════════════
REJECTION ANALYSIS REPORT
═══════════════════════════════════════════

OVERVIEW
Write 2-3 sentences summarizing the overall picture of why this application was rejected and what the committee's main concerns were.

═══════════════════════════════════════════
SECTION 1 — EVERY DEFICIENCY IDENTIFIED
═══════════════════════════════════════════
For each deficiency found, format exactly like this:

DEFICIENCY #[number]:
Category: [Case Logs / Case Reports / CE Hours / Skills / Formatting / Other]
What AVTAA said: [quote or close paraphrase of their exact language]
What this actually means: [plain English explanation of the problem]
Why it caused rejection: [explain the committee's reasoning]
Effort to fix: [Easy / Moderate / Significant]
Specific fix: [exactly what she needs to do differently — be concrete and detailed]

═══════════════════════════════════════════
SECTION 2 — CASE LOG SPECIFIC FAILURES
═══════════════════════════════════════════
If the rejection mentions case log issues, list every specific thing that was wrong with the case logs. For each:
- What was missing or wrong
- Which cases were likely affected (if mentioned)
- The exact standard she needed to meet
- Word-for-word example of how a correct entry would read

═══════════════════════════════════════════
SECTION 3 — CASE REPORT SPECIFIC FAILURES
═══════════════════════════════════════════
If the rejection mentions case report issues, analyze each report section that was flagged. For each:
- Which section failed (Signalment, Complications, Plan, Intra-op, Recovery, etc.)
- What the committee found lacking
- The specific depth and content level required
- What a passing version of that section would contain

═══════════════════════════════════════════
SECTION 4 — PERSONALIZED MISTAKE MEMORY
(THE MOST IMPORTANT SECTION)
═══════════════════════════════════════════
Based on her specific rejection reasons, create a personalized checklist of mistakes she personally made that she must NEVER repeat. This is her individual pattern of errors. Format as:

MY PERSONAL MISTAKE PATTERN — Things I Must Fix Every Time:

For EVERY CASE LOG she writes going forward:
□ [specific thing she personally missed — be very concrete]
□ [another personal mistake pattern]
[continue for all case log issues found]

For EVERY CASE REPORT she writes going forward:
□ [specific thing she personally missed]
□ [another personal mistake pattern]
[continue for all case report issues found]

Red flags that previously got her rejected — she must double-check these before submitting:
⚠ [specific red flag from her rejection]
⚠ [another red flag]
[continue for all]

═══════════════════════════════════════════
SECTION 5 — SECOND ATTEMPT PRIORITY PLAN
═══════════════════════════════════════════
Rank order the top 5 things she must fix first, with reasoning for why each is highest priority. Include time estimates for each fix.

PRIORITY 1: [title]
Why: [reasoning]
Time needed: [estimate]
Exactly how to fix it: [concrete steps]

[repeat for priorities 2-5]

═══════════════════════════════════════════
SECTION 6 — WHAT SHE DID RIGHT
═══════════════════════════════════════════
Identify anything in the rejection materials that suggests she was close or had strengths. Be specific and genuine — this is to help her confidence and know what NOT to change.

═══════════════════════════════════════════
SECTION 7 — DIRECT MESSAGE TO SUNNY DEE
═══════════════════════════════════════════
Write a brief, direct, honest and encouraging message specifically to her about what this rejection means and what she needs to hear to succeed on her second attempt. Be real with her — not generic encouragement, but specific truth about what will make the difference.`;
}

function RejectionTab() {
  const blank = () => ({ file:null, phase:"idle", step:"", text:"", result:"", errMsg:"" });
  const [state, setState] = useState(blank());
  const [manualText, setManualText] = useState("");
  const [inputMode, setInputMode] = useState("file"); // "file" | "text"
  const patch = (obj) => setState(prev => ({...prev, ...obj}));

  const onFile = (f) => setState({...blank(), file:f, phase:"ready"});

  const run = async (textOverride) => {
    const content = textOverride || manualText;
    if (!content?.trim() && !state.file) return;

    patch({phase:"running", step:"Reading rejection materials...", result:"", errMsg:""});

    try {
      let text = content || "";

      if (state.file && !textOverride) {
        patch({step:"Extracting text from rejection document..."});
        if (state.file.name.endsWith(".pdf") || state.file.type.startsWith("image/")) {
          text = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onerror = () => reject(new Error("Could not read file"));
            reader.onload = async (e) => {
              try {
                const isImage = state.file.type.startsWith("image/");
                const mediaType = isImage ? state.file.type : "application/pdf";
                const base64 = e.target.result.split(",")[1];
                const res = await fetch("https://api.anthropic.com/v1/messages", {
                  method:"POST", headers:{"Content-Type":"application/json"},
                  body: JSON.stringify({
                    model:"claude-sonnet-4-20250514", max_tokens:2000,
                    messages:[{role:"user", content:[
                      {type: isImage ? "image" : "document",
                       source:{type:"base64", media_type:mediaType, data:base64}},
                      {type:"text", text:"Extract ALL text from this AVTAA rejection letter or feedback document exactly as written. Include every word, number, and detail. Return only the document text."}
                    ]}]
                  })
                });
                const d = await res.json();
                resolve(d.content?.map(b=>b.text||"").join("") || "");
              } catch(e) { reject(e); }
            };
            if (state.file.type.startsWith("image/")) {
              reader.readAsDataURL(state.file);
            } else {
              reader.readAsDataURL(state.file);
            }
          });
        } else {
          text = await state.file.text();
        }
      }

      if (!text.trim()) throw new Error("No text could be extracted. Try uploading a different format or pasting the text manually.");

      patch({step:"Running deep rejection analysis — building your personalized fix plan...", text});

      const result = await callClaude(
        buildRejectionSystem(),
        `Analyze this AVTAA VTS rejection material thoroughly. Build a complete correction plan and personalized mistake memory:\n\n${text.slice(0, 12000)}`,
        2400
      );

      patch({phase:"done", step:"", result, text});

    } catch(err) {
      patch({phase:"error", step:"", errMsg: err.message || "Analysis failed — please try again"});
    }
  };

  return (
    <div>
      <div style={S.sectionTitle}>Rejection Analysis</div>
      <div style={S.sectionSub}>
        Upload or paste the rejection letter AVTAA sent after your first attempt. Claude runs a deep analysis — mapping every deficiency, building a personalized mistake memory, and creating a priority fix plan for your second attempt.
      </div>

      {/* Input mode toggle */}
      <div style={{display:"flex",gap:10,marginBottom:20}}>
        {[["file","📎 Upload File"],["text","✏️ Paste Text"]].map(([mode,label])=>(
          <button key={mode} onClick={()=>setInputMode(mode)}
            style={{padding:"9px 20px",borderRadius:8,border:`1px solid ${inputMode===mode?T.violet:T.border}`,background:inputMode===mode?T.violetSoft:"transparent",color:inputMode===mode?T.violetMid:T.muted,fontSize:13,fontWeight:700,cursor:"pointer",transition:"all 0.15s"}}>
            {label}
          </button>
        ))}
      </div>

      {/* Info card */}
      <div style={{...S.card,background:T.amberSoft,border:`1px solid ${T.amber}44`,marginBottom:20}}>
        <div style={{fontSize:13,color:T.amber,fontWeight:700,marginBottom:6}}>📧 Where to find your rejection materials</div>
        <div style={{fontSize:12,color:T.sub,lineHeight:1.7}}>
          After your application was rejected, AVTAA emailed a feedback overview from <strong style={{color:T.text}}>avtaa.vts.credential@gmail.com</strong>. Upload that email as a PDF, screenshot it as an image, or paste the text directly. The more detail you provide, the more specific the fix plan will be.
        </div>
      </div>

      <div style={S.card}>
        {inputMode === "file" ? (
          <>
            <div style={{fontWeight:700,fontSize:15,marginBottom:4}}>Upload Rejection Letter</div>
            <div style={{fontSize:12,color:T.muted,marginBottom:16}}>
              Accepts PDF, images (screenshot of email), or plain text files.
            </div>
            <DropZone
              accept=".pdf,.txt,.png,.jpg,.jpeg,.webp"
              label="Drop your rejection letter here"
              sublabel="PDF · Image screenshot · Text file · Click or drag & drop"
              onFile={onFile}
              fileName={state.file?.name}
            />
            {state.phase==="ready" && (
              <div style={{marginTop:14}}>
                <button style={{...S.btn("primary"),padding:"12px 28px"}} onClick={()=>run()}>
                  🔍 Run Deep Analysis
                </button>
              </div>
            )}
          </>
        ) : (
          <>
            <div style={{fontWeight:700,fontSize:15,marginBottom:4}}>Paste Rejection Text</div>
            <div style={{fontSize:12,color:T.muted,marginBottom:12}}>
              Copy and paste everything from the rejection email or feedback document.
            </div>
            <textarea
              style={{...S.textarea,minHeight:220}}
              placeholder={"Paste the full rejection feedback here. Include everything AVTAA wrote — the more detail the better.\n\nExamples of what it might say:\n• 'Case logs did not demonstrate adequate skill description...'\n• 'Drug doses listed in mL only — mg/kg required...'\n• 'Case report section 3 lacked contingency planning...'\n• 'ASA classification not adequately justified...'\n\nEven partial notes or bullet points will work."}
              value={manualText}
              onChange={e=>setManualText(e.target.value)}
            />
            <div style={{marginTop:14}}>
              <button style={{...S.btn("primary"),padding:"12px 28px"}} onClick={()=>run(manualText)} disabled={state.phase==="running"||!manualText.trim()}>
                🔍 Run Deep Analysis
              </button>
            </div>
          </>
        )}

        {/* Progress */}
        {state.phase==="running" && (
          <div style={{marginTop:14,padding:"13px 16px",background:T.violetSoft,border:`1px solid ${T.violet}44`,borderRadius:10,display:"flex",alignItems:"center",gap:10}}>
            <span style={S.spinner}/>
            <span style={{fontSize:13,color:T.violetMid,fontWeight:600}}>{state.step}</span>
          </div>
        )}

        {/* Error */}
        {state.phase==="error" && (
          <div style={{marginTop:14,padding:"13px 16px",background:T.redSoft,border:`1px solid ${T.red}44`,borderRadius:10}}>
            <div style={{fontSize:13,color:T.red,marginBottom:8}}>❌ {state.errMsg}</div>
            <button style={{...S.btn("ghost"),fontSize:12,padding:"5px 12px"}} onClick={()=>patch({phase:state.file?"ready":"idle",errMsg:""})}>Try Again</button>
          </div>
        )}

        {/* Results */}
        {state.phase==="done" && state.result && (
          <>
            <div style={{marginTop:20,padding:"10px 16px",background:T.greenSoft,border:`1px solid ${T.green}44`,borderRadius:10,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{fontSize:13,color:T.green,fontWeight:700}}>✓ Analysis complete — your personalized fix plan is ready</div>
              <button style={{...S.btn("ghost"),fontSize:11,padding:"5px 12px"}} onClick={()=>run(state.text||manualText)}>🔄 Re-analyze</button>
            </div>
            <div style={S.aiBlock}>
              <div style={{fontSize:10,color:T.violet,fontWeight:700,marginBottom:10,letterSpacing:1}}>REJECTION ANALYSIS — PERSONALIZED FIX PLAN</div>
              {state.result}
            </div>
            <div style={{display:"flex",gap:10,marginTop:12}}>
              <button style={{...S.btn("ghost"),fontSize:12}} onClick={()=>navigator.clipboard?.writeText(state.result)}>
                📋 Copy Full Report
              </button>
            </div>
          </>
        )}

        {/* Re-run after done with file mode */}
        {state.phase==="done" && (
          <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${T.border}`}}>
            <div style={{fontSize:12,color:T.muted,marginBottom:8}}>Want to analyze different rejection materials?</div>
            <button style={{...S.btn("ghost"),fontSize:12}} onClick={()=>setState(blank())}>
              Upload New File
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// TAB 4 — CASE LOG AUDITOR
// ═══════════════════════════════════════════════════════════════════════════════

// Official AVTAA requirements sourced from 2025 application packet
function buildCaseLogSystem() {
  return `You are an expert AVTAA credentials committee reviewer auditing a VTS (Anesthesia and Analgesia) case log submission. You have full knowledge of the 2025 AVTAA application packet requirements.

OFFICIAL AVTAA CASE LOG REQUIREMENTS (from the 2025 packet):

ELIGIBILITY AND VOLUME:
- Minimum 50 cases, maximum 60 cases (AVTAA strongly recommends submitting all 60)
- Cases must be performed between January 1 and December 31 of the application year
- Cases performed during on-call or overtime hours are acceptable
- The technician must have been the PRIMARY anesthesia provider — observer-only cases do not count
- All cases must be performed at an approved employment location listed on the employment history form

ASA DISTRIBUTION RULES:
- ASA I and II cases combined: maximum 12 total
- All ASA I and II cases must appear within the FIRST 12 case logs of the submission
- If any ASA I or II case appears at position 13 or beyond, it MUST be explicitly marked "SKILLS ONLY" — otherwise it is rejected and does not count toward the passing total
- Sedation-only cases: maximum 3 for small animal track, maximum 3 for large animal track
- The remaining cases should demonstrate a mix of ASA III, IV, and V patients
- IMPORTANT: A rejected case log reduces the effective passing count. An applicant can submit 60 logs but still fail if enough are rejected to drop below the 50 minimum — so EVERY log must be complete and correct, not just present

WHAT EACH CASE LOG ENTRY MUST CONTAIN:
1. Date of the case (month/day/year format)
2. Species, breed, age, sex, and body weight (with units — kg or lbs)
3. ASA physical status rating (I, II, III, IV, V, or E suffix for emergency)
4. Primary diagnosis or reason for anesthesia
5. Anesthetic procedure performed
6. Location dropdown selected (Location 1, Location 2, Secondary 1, etc.)
7. Duration of anesthesia in minutes or hours
8. All drugs administered — names AND doses in mg/kg or mcg/kg (NEVER just mL)
9. Monitoring equipment used
10. Case summary describing pre-anesthetic assessment, intraoperative management, and recovery
11. Skills demonstrated — described IN CONTEXT with technique, rationale, and patient response — NOT just a list of skill names

DRUG DOSE REQUIREMENTS (critical):
- All drug doses MUST be listed in mg/kg or mcg/kg
- Listing only mL volumes without mg/kg is an automatic rejection flag
- Ambiguous abbreviations must be defined — dex could mean dexmedetomidine, dexamethasone, or dextrose
- Generic drug names required throughout (brand names not acceptable)

SKILL DOCUMENTATION REQUIREMENTS:
- Skills must be DESCRIBED IN CONTEXT — not just listed
- A skill description must include: what was done, why it was chosen, how it was performed, and what the patient response was
- Example of WRONG: "Placed arterial catheter"
- Example of RIGHT: "Placed 22g arterial catheter in dorsal pedal artery for continuous direct arterial blood pressure monitoring due to high cardiovascular risk — MAP maintained at 68-74mmHg throughout"

BLINDING REQUIREMENTS:
- No applicant name, technician name, or facility/clinic/hospital name anywhere in the case log
- Location should be referenced as Location 1, Location 2, etc. only
- No identifying doctor names

COMMON REJECTION REASONS FROM AVTAA:
- Drug doses listed only in mL without mg/kg conversion
- Skills listed without contextual description
- More than 12 ASA I/II cases total
- ASA I/II cases at position 13+ not marked SKILLS ONLY
- More than 3 sedation-only cases
- Location dropdown field left blank on any entry
- Abbreviations used without definition (especially dex, bup, ace, val, prop, tele)
- Case summary missing pre-op assessment or recovery information
- No weight listed or weight without units
- Cases performed before January 1 or after December 31 of the application year


KNOWN REJECTION PATTERNS FROM THIS APPLICANT'S PREVIOUS SUBMISSION (Dimari Diaz #2574):
These are the EXACT mistakes that caused her case logs to be rejected. Check EVERY entry for these:

VITAL SIGNS ERRORS (caused multiple rejections):
- Writing "BP in the 50's" or "BP tanked" — NOT acceptable. Must use actual MAP/SAP values with units e.g. MAP 52mmHg
- Writing "ETCO2 in 70's" — must state exact value with units e.g. EtCO2 72mmHg
- "Maintained appropriate vitals" without actual values — committee cannot confirm knowledge without numbers
- Missing units on ANY vital sign — every value needs its unit (mmHg, bpm, degrees C, %)
- Stating hypotension but not providing the MAP or SAP value that triggered the intervention

DRUG AND DOSE ERRORS (caused multiple rejections):
- Missing doses entirely — e.g. midazolam dose missing in Case Log 58
- Epidural drugs listed with no doses — ALL epidural drug doses must be stated in mg/kg
- Giving glycopyrrolate for "hypotension" without stating the HR or BP values — must justify why glyco was the correct choice with the actual numbers
- Vague clinical reasoning — e.g. "50mcg Fent bolus IV per no one to do TAP block" is unclear; the reasoning for every drug decision must be explicit and professional
- Using "dexmed" or other shorthand ambiguously — always write dexmedetomidine with dose in mcg/kg
- Incorrect units on CRIs — e.g. dexmedetomidine CRI must be in mcg/kg/min or mcg/kg/hr, never just mcg
- KCl concentration in fluids (e.g. Plasmalyte) not specified — all additive concentrations must be stated
- Missing volumes of blood products when transfusions are given

CLINICAL REASONING AND TECHNIQUE RED FLAGS (committee judges knowledge, not just completeness):
- Questionable management of hypercapnia — e.g. Case Log 14 used "a mask over ETT to provide less than 100% oxygen to allow hypoxic drive to help vent" with EtCO2 in the 60-70mmHg range; this reflects flawed clinical reasoning, not just a wording problem
- Stating an intervention "worked" or vitals were "maintained" without the values that prove correct decision-making
- Any management choice that a board-certified anesthetist would question should be flagged as a knowledge concern, separate from formatting issues

CONFLICTING/INCONSISTENT INFORMATION:
- Equipment section says NRB circuit but case log says rebreath circuit — must match
- ASA on case report does not match case log or anesthetic sheet — all three must agree
- Patient presentation details repeated in Anesthesia Care section instead of clinical reasoning

PE AND ASSESSMENT GAPS:
- No PE findings provided to justify ASA assignment — must include exam findings AND diagnostics
- No body weight parameters mentioned — weight with units required for every case
- "No PE findings" in reason for anesthesia section — must justify ASA with clinical data
- Vague "depth checks" in the equipment/monitoring section — must specify the actual methods used: palpebral reflex, eye position, jaw tone, etc.
- No anesthetic depth assessment detail — committee needs to see the specific reflexes and signs monitored

SKILLS DOCUMENTATION ERRORS (caused majority of skill rejections):
- Just stating a skill was used without describing drug effects or properties (C3 — acepromazine)
- Administering a drug without describing rationale (C11)
- Calling propofol + inhalant a "multimodal analgesic protocol" — propofol is induction not analgesia (C13)
- Not describing antiarrhythmic drug in the designated case log (C19)
- No mention of temperature probe or temp assessment (C31)
- No Bair Hugger or active warming device mentioned (C32) — use generic term "forced warm air blanket"
- Syringe pump not described in assigned log (C45)
- Dental block: only mentioning drug without technique, landmarks, or how block was performed (C68)
- Pain scoring system not mentioned or described (C70)
- NMBA given but not labeled which drug was the neuromuscular blocking agent (S5)
- TIVA described without reasoning for drug choices (S10)
- Airway exchange catheter used but placement not described (S17)
- IV opioid induction not described in designated case log (S20)

ASA PLACEMENT ERRORS:
- Case Logs 52, 57, and 59 were ASA II appearing AFTER case log position 12 and were NOT marked as "SKILLS ONLY" — this caused rejection
- THE RULE: ASA I and II cases must appear within the first 12 case logs. Any ASA I or II case appearing at position 13 or later MUST be explicitly marked "SKILLS ONLY" or it will be rejected and will not count toward the passing total
- When auditing, check the position of every ASA I/II case and whether it carries the SKILLS ONLY designation if beyond position 12
- Brief or scant ASA I/II case logs (like 57 and 59) are also flagged even when marked correctly — they still need complete information

TERMINOLOGY AND PROFESSIONALISM:
- "Soft palate resection" instead of staphylectomy — use proper medical terminology throughout
- "BP tanked" or "BP in the 50's" — informal language, never acceptable
- Multiple spelling and typographic errors were noted — proofread every entry
- Uneven skill distribution — some logs had 3+ skills, others had none — distribute skills evenly across logs
- Blank space in stable anesthetic cases should be used to document basic skills (machine components, fluid pumps, syringe pumps, capnography, ECG, etc.)

TIMING ERROR:
- Case Log 58 dated 11/26/04 — outside the application year, automatically rejected

RECOMMENDATIONS BASED ON AVTAA GUIDANCE:
- Aim for 60 cases, not 50 — provides buffer if any are rejected
- Include a variety of species if possible to demonstrate breadth
- Include a variety of ASA III-V cases showing increasing complexity
- Ensure at least some cases involve CRIs, epidurals, or regional blocks if applicable to your track
- Each case should demonstrate a different skill or technique where possible
- Cases with complications or challenging management score better than routine cases

YOUR TASK:
Perform a comprehensive line-by-line audit of this case log submission. For each problem you find, cite the specific case number or entry where it occurs.

Structure your response as follows:

CASE LOG AUDIT SUMMARY
Total cases identified: [number]
ASA I/II count: [number] — [COMPLIANT or VIOLATION]
Sedation-only count: [number] — [COMPLIANT or VIOLATION]
Estimated compliance score: [X/10]

CRITICAL VIOLATIONS (automatic rejection risks):
[numbered list — cite specific case numbers]

WARNINGS (weakens application):
[numbered list — cite specific case numbers where possible]

MISSING ELEMENTS BY CASE:
[for each case that is incomplete, list what is missing]

SKILL DOCUMENTATION QUALITY:
[assess whether skills are described in context or just listed]

DRUG NOTATION AUDIT:
[list any mL-only doses found, cite case numbers]

RECOMMENDATIONS TO STRENGTHEN BEFORE SUBMISSION:
[numbered list of specific improvements she should make]

WHAT IS DONE WELL:
[genuine strengths to acknowledge]`;
}

function CaseLogTab({caseLogCount,setCaseLogCount,asaCounts,setAsaCounts,sedationOnly,setSedationOnly,caseLogRecords,setCaseLogRecords}) {
  const blank = () => ({ file:null, phase:"idle", step:"", text:"", result:"", errMsg:"" });
  const [state, setState] = useState(blank());
  const patch = (obj) => setState(prev => ({...prev, ...obj}));

  const [analyzing, setAnalyzing]   = useState(false);
  const [entryResult, setEntryResult] = useState("");

  const analyzeEntry = async () => {
    const content = [
      entry.num    && `Case Log #: ${entry.num}`,
      entry.date   && `Date: ${entry.date}`,
      entry.species&& `Species/Breed: ${entry.species} / ${entry.breed}`,
      entry.age    && `Age/Sex/Weight: ${entry.age} / ${entry.sex} / ${entry.weight}`,
      `ASA Status: ${entry.asa}${entry.skillsOnly?" (SKILLS ONLY)":""}${entry.sedationOnly?" (SEDATION ONLY)":""}`,
      entry.location&&`Location: ${entry.location}`,
      entry.duration&&`Duration: ${entry.duration}`,
      entry.diagnosis&&`Diagnosis: ${entry.diagnosis}`,
      entry.drugs  && `Drugs & Doses:\n${entry.drugs}`,
      entry.monitoring&&`Monitoring: ${entry.monitoring}`,
      entry.summary&&`Case Summary:\n${entry.summary}`,
      entry.skills &&`Skills Documented:\n${entry.skills}`,
    ].filter(Boolean).join("\n\n");

    if (!content.trim()) return;
    setAnalyzing(true);
    setEntryResult("");
    try {
      const result = await callClaude(
        buildCaseLogSystem(),
        `Analyze this SINGLE AVTAA case log entry for compliance issues. Be specific — flag every problem with drug doses, vital sign notation, skill descriptions, ASA placement, missing fields, terminology, and clinical reasoning. Also note what is done well:\n\n${content}`,
        1400
      );
      setEntryResult(result);
    } catch(err) {
      setEntryResult("Analysis failed: " + err.message);
    }
    setAnalyzing(false);
  };

  // ── Export all records as AVTAA-formatted PDF using reportlab via API ──────
  const exportPDF = async () => {
    if (caseLogRecords.length === 0) return;
    try {
      // Build plain text version of all cases for Claude to format
      const allCases = caseLogRecords.map(r => [
        `CASE LOG #${r.num}`,
        `Date: ${r.date||""}  |  Location: ${r.location||""}  |  Duration: ${r.duration||""}`,
        `Species/Breed: ${r.species||""} / ${r.breed||""}`,
        `Age/Sex/Weight: ${r.age||""} / ${r.sex||""} / ${r.weight||""}`,
        `ASA Physical Status: ${r.asa||""}${r.skillsOnly?" — SKILLS ONLY":""}${r.sedationOnly?" — SEDATION ONLY":""}`,
        `Diagnosis/Reason for Anesthesia: ${r.diagnosis||""}`,
        `\nDrugs Administered:\n${r.drugs||""}`,
        `\nMonitoring Equipment:\n${r.monitoring||""}`,
        `\nCase Summary:\n${r.summary||""}`,
        `\nSkills Demonstrated:\n${r.skills||""}`,
      ].join("\n")).join("\n\n" + "═".repeat(60) + "\n\n");

      // Build HTML blob that renders nicely and can be printed to PDF
      const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>AVTAA Case Logs — Applicant XXXX</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:'Times New Roman',Times,serif; font-size:10.5pt; line-height:1.5; color:#1a1a1a; background:#fff; }
  .cover { text-align:center; padding:80px 40px; border-bottom:2px solid #2D1F5E; margin-bottom:40px; }
  .cover h1 { font-size:20pt; font-weight:bold; color:#2D1F5E; margin-bottom:10px; }
  .cover p { font-size:11pt; color:#555; margin-top:8px; }
  .case { page-break-before:always; padding:36px 48px; }
  .case:first-of-type { page-break-before:avoid; }
  .case-header { background:#2D1F5E; color:#fff; padding:10px 16px; border-radius:4px; margin-bottom:16px; }
  .case-header h2 { font-size:13pt; font-weight:bold; }
  .case-header p { font-size:9.5pt; margin-top:3px; opacity:0.85; }
  .meta-grid { display:grid; grid-template-columns:1fr 1fr; gap:4px 24px; margin-bottom:14px; font-size:10pt; }
  .meta-item { display:flex; gap:6px; }
  .meta-label { font-weight:bold; color:#2D1F5E; min-width:120px; flex-shrink:0; }
  .asa-badge { display:inline-block; padding:3px 12px; border-radius:12px; font-weight:bold; font-size:9pt;
    background:${`#2D1F5E`}; color:white; margin-bottom:12px; }
  .section { margin-bottom:14px; }
  .section-title { font-weight:bold; font-size:10pt; color:#2D1F5E; border-bottom:1px solid #C4B5FD; padding-bottom:3px; margin-bottom:6px; text-transform:uppercase; letter-spacing:0.5px; }
  .section-body { font-size:10.5pt; line-height:1.6; white-space:pre-wrap; }
  .tag { display:inline-block; padding:2px 10px; border-radius:10px; font-size:9pt; font-weight:bold; margin-left:8px; }
  .tag-skills { background:#EDE9FE; color:#7C3AED; }
  .tag-sed { background:#FEF3C7; color:#92400E; }
  .footer { text-align:center; font-size:9pt; color:#888; padding:20px; border-top:1px solid #ddd; margin-top:40px; }
  @media print {
    body { -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    .case { page-break-before:always; }
    .case:first-of-type { page-break-before:avoid; }
  }
</style>
</head>
<body>
<div class="cover">
  <h1>AVTAA VTS Case Logs</h1>
  <p>Applicant Number: XXXX &nbsp;|&nbsp; Small Animal Track</p>
  <p>Application Year: 2024 &nbsp;|&nbsp; Total Cases: ${caseLogRecords.length}</p>
  <p style="margin-top:16px;font-size:9.5pt;color:#888;">Generated by VTS Compass &nbsp;·&nbsp; Review all entries before submission</p>
</div>
${caseLogRecords.map((r,i) => `
<div class="case">
  <div class="case-header">
    <h2>Case Log #${r.num}${r.skillsOnly?' <span style="font-size:10pt;opacity:0.85;">[SKILLS ONLY]</span>':''}${r.sedationOnly?' <span style="font-size:10pt;opacity:0.85;">[SEDATION ONLY]</span>':''}</h2>
    <p>${r.date||''} &nbsp;·&nbsp; ${r.location||''} &nbsp;·&nbsp; ${r.duration||''}</p>
  </div>
  <div class="meta-grid">
    <div class="meta-item"><span class="meta-label">Species/Breed:</span><span>${r.species||'—'} / ${r.breed||'—'}</span></div>
    <div class="meta-item"><span class="meta-label">Age/Sex/Weight:</span><span>${r.age||'—'} / ${r.sex||'—'} / ${r.weight||'—'}</span></div>
    <div class="meta-item"><span class="meta-label">ASA Status:</span><span style="font-weight:bold;color:#2D1F5E;">ASA ${r.asa||'—'}</span></div>
    <div class="meta-item"><span class="meta-label">Diagnosis:</span><span>${r.diagnosis||'—'}</span></div>
  </div>
  ${r.drugs?`<div class="section"><div class="section-title">Drugs Administered</div><div class="section-body">${r.drugs}</div></div>`:''}
  ${r.monitoring?`<div class="section"><div class="section-title">Monitoring Equipment</div><div class="section-body">${r.monitoring}</div></div>`:''}
  ${r.summary?`<div class="section"><div class="section-title">Case Summary</div><div class="section-body">${r.summary}</div></div>`:''}
  ${r.skills?`<div class="section"><div class="section-title">Skills Demonstrated</div><div class="section-body">${r.skills}</div></div>`:''}
</div>`).join('')}
<div class="footer">VTS Compass &nbsp;·&nbsp; ${caseLogRecords.length} case logs &nbsp;·&nbsp; Review all entries before AVTAA submission</div>
</body>
</html>`;

      const blob = new Blob([html], {type:"text/html"});
      const url  = URL.createObjectURL(blob);
      const w    = window.open(url, "_blank");
      if (w) {
        w.onload = () => {
          setTimeout(() => { w.print(); }, 500);
        };
      }
      setTimeout(() => URL.revokeObjectURL(url), 10000);
    } catch(err) {
      alert("Export failed: " + err.message);
    }
  };

  const blankEntry = () => ({
    num:"", date:"", species:"", breed:"", age:"", sex:"", weight:"",
    asa:"III", location:"Location 1", duration:"", diagnosis:"",
    drugs:"", monitoring:"", summary:"", skills:"",
    sedationOnly:false, skillsOnly:false,
  });
  const [showRecords,  setShowRecords]  = useState(false);
  const [addingCase,   setAddingCase]   = useState(false);
  const [selectedCase, setSelectedCase] = useState(null);
  const [entry,        setEntry]        = useState(blankEntry());
  const setField = (k,v) => setEntry(p=>({...p,[k]:v}));

  const saveEntry = () => {
    if (!entry.num.trim()) return;
    const updated = caseLogRecords
      .filter(r=>r.num!==entry.num)
      .concat(entry)
      .sort((a,b)=>parseInt(a.num)-parseInt(b.num));
    setCaseLogRecords(updated);
    setEntry(blankEntry()); setAddingCase(false); setSelectedCase(entry.num);
  };

  const deleteCase = (num) => {
    setCaseLogRecords(prev=>prev.filter(r=>r.num!==num));
    if(selectedCase===num) setSelectedCase(null);
  };

  const editCase = (rec) => { setEntry({...rec}); setAddingCase(true); setShowRecords(true); };
  const onFile   = (f)   => setState({...blank(), file:f, phase:"ready"});

  const run = async () => {
    if (!state.file) return;
    patch({phase:"running", step:"Reading and extracting case log...", result:"", errMsg:""});
    try {
      let text = "";
      if (state.file.name.endsWith(".pdf")) {
        text = await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onerror = () => reject(new Error("Could not read PDF"));
          reader.onload = async (e) => {
            try {
              const base64 = e.target.result.split(",")[1];
              const res = await fetch("https://api.anthropic.com/v1/messages", {
                method:"POST", headers:{"Content-Type":"application/json"},
                body: JSON.stringify({
                  model:"claude-sonnet-4-20250514", max_tokens:1400,
                  messages:[{role:"user", content:[
                    {type:"document", source:{type:"base64", media_type:"application/pdf", data:base64}},
                    {type:"text", text:"Extract ALL text from this AVTAA case log PDF exactly as written. Preserve all case numbers, dates, drug names, doses, and descriptions. Return only the document text."}
                  ]}]
                })
              });
              const d = await res.json();
              resolve(d.content?.map(b=>b.text||"").join("") || "");
            } catch(e) { reject(e); }
          };
          reader.readAsDataURL(state.file);
        });
      } else {
        text = await state.file.text();
      }
      if (!text.trim()) throw new Error("No text could be extracted from this file.");
      const caseMatches = text.match(/\bcase\s*(?:log\s*)?#?\s*\d+/gi) || [];
      if (caseMatches.length > 0) setCaseLogCount(caseMatches.length);
      patch({step:"Running comprehensive AVTAA compliance audit...", text});
      const result = await callClaude(
        buildCaseLogSystem(),
        `Perform a full AVTAA case log audit on this submission. Be thorough — check every entry:\n\n${text.slice(0, 12000)}`,
        1800
      );
      patch({phase:"done", step:"", result, text});
    } catch(err) {
      patch({phase:"error", step:"", errMsg: err.message});
    }
  };

  const lowASA = (asaCounts.I||0) + (asaCounts.II||0);
  const selectedRec = caseLogRecords.find(r=>r.num===selectedCase);
  const ASA_COLORS  = {I:T.green,II:T.green,III:T.amber,IV:T.red,V:T.red,E:T.red,IVE:T.red,VE:T.red,IIIE:T.amber};

  return (
    <div>
      <div style={S.sectionTitle}>Case Log Auditor</div>
      <div style={S.sectionSub}>Upload your AVTAA case log PDF for a full compliance audit. Use the Case Log Records section to manually add and review each individual case.</div>

      {/* Stats */}
      <div style={S.grid2}>
        <div style={S.statCard(caseLogCount>=60?"green":caseLogCount>=50?"amber":caseLogCount>0?"red":"violet")}>
          <div style={{fontSize:11,color:T.muted,fontWeight:700,marginBottom:4}}>TOTAL CASE LOGS</div>
          <div style={{fontSize:30,fontWeight:900,color:caseLogCount>=60?T.green:caseLogCount>=50?T.amber:caseLogCount>0?T.red:T.muted}}>{caseLogCount||"—"}</div>
          <div style={{fontSize:11,color:T.muted,marginBottom:6}}>Need 50–60 · {caseLogRecords.length>0?"auto-tracked from records below":"use +/− or add cases below"}</div>
          {caseLogRecords.length===0&&<div style={{display:"flex",gap:6}}>
            <button style={{...S.btn("ghost"),padding:"4px 12px"}} onClick={()=>setCaseLogCount(Math.max(0,caseLogCount-1))}>−</button>
            <button style={{...S.btn("ghost"),padding:"4px 12px"}} onClick={()=>setCaseLogCount(caseLogCount+1)}>+</button>
          </div>}
        </div>
        <div style={S.statCard(lowASA<=12&&lowASA>0?"green":lowASA>12?"red":"violet")}>
          <div style={{fontSize:11,color:T.muted,fontWeight:700,marginBottom:4}}>ASA I/II COUNT</div>
          <div style={{fontSize:30,fontWeight:900,color:lowASA<=12&&lowASA>0?T.green:lowASA>12?T.red:T.muted}}>{lowASA||"—"}</div>
          <div style={{fontSize:11,color:T.muted,marginBottom:6}}>Max 12 — within first 12 logs</div>
          <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
            {["I","II","III","IV","V","E"].map(a=>(
              <div key={a} style={{textAlign:"center"}}>
                <div style={{fontSize:9,color:T.muted,marginBottom:2}}>ASA{a}</div>
                <input type="number" min={0} style={{...S.input,width:34,padding:"3px",textAlign:"center",fontSize:11}}
                  value={asaCounts[a]||0} onChange={e=>setAsaCounts(p=>({...p,[a]:parseInt(e.target.value)||0}))}/>
              </div>
            ))}
          </div>
        </div>
        <div style={S.statCard(sedationOnly<=3?"green":"red")}>
          <div style={{fontSize:11,color:T.muted,fontWeight:700,marginBottom:4}}>SEDATION-ONLY CASES</div>
          <div style={{fontSize:30,fontWeight:900,color:sedationOnly<=3?T.green:T.red}}>{sedationOnly}</div>
          <div style={{fontSize:11,color:T.muted,marginBottom:8}}>Max 3 (both tracks)</div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            <button style={{...S.btn("ghost"),padding:"4px 14px",fontSize:15}} onClick={()=>setSedationOnly(Math.max(0,sedationOnly-1))}>−</button>
            <button style={{...S.btn("ghost"),padding:"4px 14px",fontSize:15}} onClick={()=>setSedationOnly(sedationOnly+1)}>+</button>
          </div>
        </div>
        <div style={S.statCard("violet")}>
          <div style={{fontSize:11,color:T.muted,fontWeight:700,marginBottom:8}}>OFFICIAL REQUIREMENTS</div>
          <div style={{fontSize:11,color:T.sub,lineHeight:1.85}}>
            {"📋 50-60 cases, Jan 1–Dec 31"}<br/>
            {"💊 All doses in mg/kg or mcg/kg"}<br/>
            {"📍 Location dropdown on every case"}<br/>
            {"🔬 Skills described in context — not listed"}<br/>
            {"🚫 No names or facility info (Folder 2)"}<br/>
            {"⚠️ ASA I/II max 12, within first 12 logs"}<br/>
            {"💉 Max 3 sedation-only cases"}<br/>
            {"✍️ Define all abbreviations (dex, bup, ace)"}
          </div>
        </div>
      </div>

      {/* ── CASE LOG RECORDS ─────────────────────────────────────────────── */}
      <div style={S.card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16,flexWrap:"wrap",gap:10}}>
          <div>
            <div style={{fontWeight:800,fontSize:15}}>📋 Case Log Records</div>
            <div style={{fontSize:12,color:T.muted,marginTop:2}}>
              {caseLogRecords.length > 0
                ? `${caseLogRecords.length} case${caseLogRecords.length!==1?"s":""} recorded — select from dropdown to review`
                : "Manually track each individual case log entry"}
            </div>
          </div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {caseLogRecords.length > 0 && (
              <>
                <button style={{...S.btn("ghost"),fontSize:12,padding:"7px 14px"}}
                  onClick={()=>{setShowRecords(s=>!s); setAddingCase(false);}}>
                  {showRecords?"▲ Hide List":"▼ View All"}
                </button>
                <button style={{...S.btn("teal"),fontSize:12,padding:"7px 14px"}}
                  onClick={exportPDF} title="Opens a print-ready HTML page — use Print → Save as PDF">
                  📄 Export PDF
                </button>
              </>
            )}
            <button style={{...S.btn("primary"),fontSize:12,padding:"7px 14px"}}
              onClick={()=>{setAddingCase(true); setShowRecords(true); setEntry(blankEntry()); setEntryResult("");}}>
              + Add Case
            </button>
          </div>
        </div>

        {/* Dropdown */}
        {caseLogRecords.length > 0 && (
          <div style={{marginBottom:14}}>
            <select style={{...S.input,width:"100%",cursor:"pointer"}}
              value={selectedCase||""}
              onChange={e=>{setSelectedCase(e.target.value||null); setAddingCase(false);}}>
              <option value="">— Select a case to review —</option>
              {caseLogRecords.map(r=>(
                <option key={r.num} value={r.num}>
                  {"#"+r.num+" · "+(r.date||"no date")+" · "+(r.species||"?")+"/"+(r.breed||"?")+" · ASA "+r.asa+(r.skillsOnly?" [SKILLS ONLY]":"")+(r.sedationOnly?" [SED ONLY]":"")+" · "+((r.diagnosis||"").slice(0,45)||"no diagnosis")}
                </option>
              ))}
            </select>
          </div>
        )}

        {/* Selected case detail */}
        {selectedRec && !addingCase && (
          <div style={{background:"#0e0c18",border:`1px solid ${T.border}`,borderRadius:12,padding:20,marginBottom:14}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14,flexWrap:"wrap",gap:8}}>
              <div>
                <div style={{fontSize:17,fontWeight:900,color:T.text}}>Case Log #{selectedRec.num}</div>
                <div style={{fontSize:12,color:T.muted,marginTop:2}}>{selectedRec.date} · {selectedRec.duration} · {selectedRec.location}</div>
              </div>
              <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                <span style={{padding:"4px 12px",borderRadius:20,background:(ASA_COLORS[selectedRec.asa]||T.amber)+"22",border:`1px solid ${(ASA_COLORS[selectedRec.asa]||T.amber)}55`,fontSize:12,fontWeight:700,color:ASA_COLORS[selectedRec.asa]||T.amber}}>
                  ASA {selectedRec.asa}{selectedRec.skillsOnly?" · SKILLS ONLY":""}
                </span>
                {selectedRec.sedationOnly && <span style={{padding:"4px 10px",borderRadius:20,background:T.amberSoft,border:`1px solid ${T.amber}55`,fontSize:11,color:T.amber,fontWeight:700}}>Sedation Only</span>}
                <button style={{...S.btn("ghost"),fontSize:11,padding:"5px 10px"}} onClick={()=>editCase(selectedRec)}>✏️ Edit</button>
                <button style={{...S.btn("ghost"),fontSize:11,padding:"5px 10px",color:T.red}} onClick={()=>deleteCase(selectedRec.num)}>🗑 Delete</button>
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"6px 24px",fontSize:12,marginBottom:14}}>
              {[["Patient",`${selectedRec.species||"—"} / ${selectedRec.breed||"—"}`],
                ["Age/Sex/Weight",`${selectedRec.age||"—"} / ${selectedRec.sex||"—"} / ${selectedRec.weight||"—"}`],
                ["Diagnosis",selectedRec.diagnosis||"—"],
              ].map(([l,v])=>(
                <div key={l}><span style={{color:T.muted,fontWeight:700}}>{l}: </span><span style={{color:T.sub}}>{v}</span></div>
              ))}
            </div>
            {[["DRUGS & DOSES","drugs"],["MONITORING","monitoring"],["CASE SUMMARY","summary"],["SKILLS","skills"]].map(([label,key])=>
              selectedRec[key] ? (
                <div key={key} style={{marginBottom:10}}>
                  <div style={{fontSize:11,color:T.muted,fontWeight:700,marginBottom:4,letterSpacing:0.5}}>{label}</div>
                  <div style={{fontSize:12,color:T.sub,lineHeight:1.7,whiteSpace:"pre-wrap"}}>{selectedRec[key]}</div>
                </div>
              ) : null
            )}
          </div>
        )}

        {/* Compact list */}
        {showRecords && !addingCase && caseLogRecords.length > 0 && (
          <div style={{marginBottom:14}}>
            <div style={{fontSize:11,color:T.muted,fontWeight:700,letterSpacing:0.5,marginBottom:8}}>ALL {caseLogRecords.length} RECORDED CASES</div>
            <div style={{display:"flex",flexDirection:"column",gap:5,maxHeight:280,overflowY:"auto"}}>
              {caseLogRecords.map(r=>(
                <div key={r.num} onClick={()=>{setSelectedCase(r.num);setAddingCase(false);}}
                  style={{display:"flex",alignItems:"center",gap:12,padding:"8px 14px",borderRadius:8,
                    border:`1px solid ${selectedCase===r.num?T.violet:T.border}`,
                    background:selectedCase===r.num?T.violetSoft:"transparent",
                    cursor:"pointer",transition:"all 0.15s"}}>
                  <div style={{fontWeight:800,fontSize:13,color:T.violetMid,minWidth:28}}>#{r.num}</div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:12,color:T.text,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                      {r.species}/{r.breed} · {(r.diagnosis||"no diagnosis").slice(0,55)}
                    </div>
                    <div style={{fontSize:11,color:T.muted}}>{r.date} · {r.duration} · ASA {r.asa}{r.skillsOnly?" · SKILLS ONLY":""}{r.sedationOnly?" · SED ONLY":""}</div>
                  </div>
                  <span style={{padding:"2px 8px",borderRadius:10,fontSize:10,fontWeight:700,background:(ASA_COLORS[r.asa]||T.amber)+"22",color:ASA_COLORS[r.asa]||T.amber}}>ASA {r.asa}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Add / Edit form */}
        {addingCase && (
          <div style={{background:"#0e0c18",border:`1px solid ${T.violet}44`,borderRadius:12,padding:20,marginBottom:14}}>
            <div style={{fontWeight:800,fontSize:14,color:T.text,marginBottom:16}}>
              {entry.num && caseLogRecords.find(r=>r.num===entry.num) ? `✏️ Editing Case Log #${entry.num}` : "➕ New Case Log Entry"}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:12}}>
              {[["Case Log #","num","e.g. 1"],["Date","date","MM/DD/YYYY"],["Duration","duration","e.g. 1hr 30min"],
                ["Species","species","e.g. Canine"],["Breed","breed","e.g. Labrador"],["Age","age","e.g. 6yr"],
                ["Sex","sex","MN / FS / MI / FI"],["Weight (with units)","weight","e.g. 32kg"],["Diagnosis","diagnosis","primary diagnosis"],
              ].map(([label,key,ph])=>(
                <div key={key}>
                  <label style={{...S.label,marginBottom:4}}>{label}</label>
                  <input style={S.input} placeholder={ph} value={entry[key]} onChange={e=>setField(key,e.target.value)}/>
                </div>
              ))}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:12}}>
              <div>
                <label style={{...S.label,marginBottom:4}}>ASA Status</label>
                <select style={S.input} value={entry.asa} onChange={e=>setField("asa",e.target.value)}>
                  {["I","II","III","IV","V","IVE","VE","IIIE"].map(a=><option key={a} value={a}>ASA {a}</option>)}
                </select>
              </div>
              <div>
                <label style={{...S.label,marginBottom:4}}>Location</label>
                <select style={S.input} value={entry.location} onChange={e=>setField("location",e.target.value)}>
                  {["Location 1","Location 2","Location 3","Secondary 1","Secondary 2"].map(l=><option key={l}>{l}</option>)}
                </select>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:12,justifyContent:"center",paddingTop:18}}>
                <label style={{display:"flex",alignItems:"center",gap:8,fontSize:12,color:T.sub,cursor:"pointer"}}>
                  <input type="checkbox" checked={entry.sedationOnly} onChange={e=>setField("sedationOnly",e.target.checked)}/>
                  Sedation Only
                </label>
                <label style={{display:"flex",alignItems:"center",gap:8,fontSize:12,color:T.sub,cursor:"pointer"}}>
                  <input type="checkbox" checked={entry.skillsOnly} onChange={e=>setField("skillsOnly",e.target.checked)}/>
                  Skills Only (ASA I/II past position 12)
                </label>
              </div>
            </div>
            {[["Monitoring","monitoring","ECG, SpO2, EtCO2, direct arterial BP, temp probe..."],
              ["Drugs & Doses","drugs","Hydromorphone 0.1mg/kg IV, Propofol 2mg/kg IV, Isoflurane 1.5-2%..."],
              ["Case Summary","summary","Describe anesthetic event with actual vital sign values (e.g. MAP 68mmHg), timestamps, interventions with doses, and patient responses..."],
              ["Skills Documented","skills","Describe each skill in context — what you did, why, technique/landmarks, and patient response..."],
            ].map(([label,key,ph])=>(
              <div key={key} style={{marginBottom:12}}>
                <label style={{...S.label,marginBottom:4}}>{label}</label>
                <textarea style={{...S.textarea,minHeight:key==="summary"||key==="skills"?110:55}}
                  placeholder={ph} value={entry[key]} onChange={e=>setField(key,e.target.value)}/>
              </div>
            ))}
            <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
              <button style={{...S.btn("primary"),padding:"10px 24px"}} onClick={saveEntry} disabled={!entry.num.trim()}>
                💾 Save Case Log
              </button>
              <button style={{...S.btn("teal"),padding:"10px 22px"}} onClick={analyzeEntry}
                disabled={analyzing||(!entry.summary&&!entry.drugs&&!entry.skills)}>
                {analyzing?<><span style={S.spinner}/>Analyzing...</>:"🔍 Analyze This Entry"}
              </button>
              <button style={{...S.btn("ghost"),padding:"10px 20px"}}
                onClick={()=>{setAddingCase(false);setEntry(blankEntry());setEntryResult("");}}>
                Cancel
              </button>
            </div>
            {entryResult && (
              <div style={{marginTop:16}}>
                <div style={{fontSize:11,color:T.teal,fontWeight:700,marginBottom:8,letterSpacing:0.5}}>🔍 ENTRY ANALYSIS</div>
                <AIResult result={entryResult}/>
              </div>
            )}
          </div>
        )}

        {caseLogRecords.length===0 && !addingCase && (
          <div style={{textAlign:"center",padding:"24px 0",color:T.muted,fontSize:13}}>
            No case logs recorded yet. Click <strong style={{color:T.violetMid}}>+ Add Case</strong> to start tracking.
          </div>
        )}
      </div>

      {/* Upload audit card */}
      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:15,marginBottom:4}}>Upload Case Log PDF for Full AI Audit</div>
        <div style={{fontSize:12,color:T.muted,marginBottom:16}}>
          Upload your [applicant#.caselog.pdf] for a comprehensive compliance check against all 2025 AVTAA requirements.
        </div>
        <DropZone accept=".pdf,.txt" label="Drop your case log PDF here" sublabel="applicant#.caselog.pdf · Click or drag & drop" onFile={onFile} fileName={state.file?.name}/>
        {state.phase==="running" && (
          <div style={{marginTop:14,padding:"13px 16px",background:T.violetSoft,border:`1px solid ${T.violet}44`,borderRadius:10,display:"flex",alignItems:"center",gap:10}}>
            <span style={S.spinner}/><span style={{fontSize:13,color:T.violetMid,fontWeight:600}}>{state.step}</span>
          </div>
        )}
        {state.phase==="error" && (
          <div style={{marginTop:12,padding:"10px 14px",background:T.redSoft,border:`1px solid ${T.red}44`,borderRadius:8,fontSize:13,color:T.red}}>
            ❌ {state.errMsg}
            <button style={{...S.btn("ghost"),marginLeft:12,fontSize:12,padding:"4px 10px"}} onClick={()=>patch({phase:"ready",errMsg:""})}>Try Again</button>
          </div>
        )}
        {(state.phase==="ready"||state.phase==="done") && (
          <div style={{marginTop:14,display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
            <button style={{...S.btn("primary"),padding:"11px 24px"}} onClick={run} disabled={state.phase==="running"}>
              {state.phase==="done"?"🔄 Re-Audit Case Log":"🔍 Run Full Compliance Audit"}
            </button>
            {state.phase==="done"&&state.text&&(
              <span style={{fontSize:12,color:T.muted}}>{state.text.split(/\s+/).filter(Boolean).length.toLocaleString()} words extracted</span>
            )}
          </div>
        )}
        {state.result&&<AIResult result={state.result}/>}
      </div>

      <div style={{...S.card,background:"#0f1a0f",border:`1px solid ${T.green}33`}}>
        <div style={{fontSize:13,color:T.green,fontWeight:700,marginBottom:12}}>📋 What AVTAA Recommends You Include</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px 24px",fontSize:12,color:T.sub,lineHeight:1.7}}>
          <div><strong style={{color:T.text}}>Case variety:</strong> Mix of species, ASA levels, and procedure types shows breadth</div>
          <div><strong style={{color:T.text}}>Case complexity:</strong> Include cases with complications — they score better than routine cases</div>
          <div><strong style={{color:T.text}}>Advanced techniques:</strong> Arterial lines, epidurals, nerve blocks, CRIs — map directly to skills list</div>
          <div><strong style={{color:T.text}}>Skill context:</strong> What you did, why, how, and patient response — every time</div>
          <div><strong style={{color:T.text}}>Emergency cases:</strong> ASA E cases demonstrate high-risk patient management ability</div>
          <div><strong style={{color:T.text}}>Volume:</strong> Submit all 60 — buffer if any are rejected during review</div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// TAB 5 — CASE REPORT REVIEWER  (rebuilt scoring engine)
// ═══════════════════════════════════════════════════════════════════════════════

// Each section is scored 1–10. Weighted to 100 points total.
// Weights reflect actual AVTAA emphasis: Sections 3,4,5 carry the most.
const SECTIONS = [
  {
    key:"s1", label:"Signalment & Physical Status", weight:10,
    rubric:`Score 1-2: Missing multiple required elements (species/age/sex/weight/PE/labs). ASA not stated or wrong class.
Score 3-4: Most elements present but labs lack reference ranges, or ASA justification is vague (e.g. "high risk").
Score 5-6: All required elements present. Labs listed but some reference ranges missing. ASA stated with partial explanation.
Score 7-8: Complete PE findings, all labs WITH reference ranges in parentheses, relevant history, current meds, ASA III-V with clear written justification linking specific diseases to anesthetic risk.
Score 9-10: All of the above PLUS: pertinent negatives noted, prior anesthetic history addressed, each disease process explicitly connected to how it elevates anesthetic risk and affects drug choices.`,
  },
  {
    key:"s2", label:"Reason for Anesthesia", weight:5,
    rubric:`Score 1-4: Procedure not clearly stated or diagnosis missing.
Score 5-7: Procedure and diagnosis stated clearly.
Score 8-10: Procedure stated, DVM diagnosis included, brief relevant clinical context provided that sets up the complexity of the case.`,
  },
  {
    key:"s3", label:"Anticipated Complications", weight:18,
    rubric:`Score 1-2: Complications listed as a bare list with no explanation. No contingency plans at all.
Score 3-4: A few complications mentioned with minimal explanation. No specific "if X then Y" contingency plans.
Score 5-6: Multiple complications identified with some explanation of WHY they are anticipated. At least one contingency plan mentioned but others missing.
Score 7-8: All major anticipated complications covered: (a) how patient's comorbidities affect anesthesia, (b) drug-related adverse effects expected, (c) procedural complications. Most have specific contingency plans with drugs and doses.
Score 9-10: Comprehensive anticipation of ALL complications with specific, dose-ready contingency plans for each. Demonstrates proactive clinical reasoning — e.g. "if MAP falls below 60mmHg I will initiate dopamine CRI at 5mcg/kg/min" not just "I will treat hypotension." Drug interactions, species-specific concerns, and emergency scenarios all addressed.`,
  },
  {
    key:"s4", label:"Anesthesia Plan", weight:22,
    rubric:`Score 1-2: Drugs listed with no doses, no rationale. Monitoring not explained. No DVM approval noted.
Score 3-4: Some drugs with doses but rationale missing or superficial. Fluid rate missing. Pain management incomplete. DVM approval absent.
Score 5-6: Drugs listed with mg/kg doses. Basic rationale provided. Fluid type and rate included. Some pain management mentioned. Monitoring listed but normal ranges not given.
Score 7-8: All drugs listed with mg/kg or mcg/kg doses AND clear rationale for each choice (including why certain drugs were AVOIDED). Fluid type, rate, and rationale. Multimodal pain management for pre/intra/post-op phases. All monitoring parameters listed WITH target normal ranges. DVM approval of plan documented.
Score 9-10: All of the above PLUS: drug interactions acknowledged, alternative drugs considered, CRI rates and calculations shown, monitoring parameters explained in terms of what information they provide and how they guide management. Complete pain management strategy with specific drugs, doses, timing, and reassessment plan.`,
  },
  {
    key:"s5", label:"Intra-op Care & Patient Support", weight:25,
    rubric:`Score 1-2: Vague narrative with no actual vital sign values. "BP dropped, we treated it" type entries. No timestamps.
Score 3-4: Some values provided but inconsistently. Timeline loose. Technician's specific actions not clearly described. Problems mentioned but management vague.
Score 5-6: Reasonable timeline with some actual vital sign numbers. Complications described with some management detail. Technician's role partially visible.
Score 7-8: Clear timeline with actual timestamps. All vital signs documented with values (HR, MAP, SpO2, EtCO2, temp). Each problem identified, assessed, and managed with specific interventions and doses. Technician clearly drove decisions. Results of interventions noted.
Score 9-10: Comprehensive play-by-play with timestamps and specific values throughout. Every deviation from the plan is explained. Every intervention includes dose, route, and patient response. Demonstrates the technician independently recognized, diagnosed, and managed complications. Shows critical thinking not just task execution. Discrepancies between plan and reality are explained.`,
  },
  {
    key:"s6", label:"Recovery & Post-op Analgesia", weight:12,
    rubric:`Score 1-2: Recovery barely mentioned. No pain assessment. No post-op analgesic plan.
Score 3-4: Recovery described briefly. Pain mentioned but no scale used or documented. Post-op analgesia listed but incomplete.
Score 5-6: Recovery described with some monitoring detail. Pain scale mentioned. Post-op analgesic plan present with drugs but doses or timing incomplete.
Score 7-8: Recovery monitored with specific parameters and values. Validated pain scale used (e.g. Colorado State, Glasgow) with actual score documented. Complete post-op analgesic plan with drugs, doses, route, and frequency. Patient support described (positioning, warming, oxygen). Quality of recovery assessed.
Score 9-10: All of the above PLUS: pain reassessment timeline established, rescue analgesia plan documented, complications during recovery identified and managed, discharge criteria or handoff instructions to next care team documented.`,
  },
  {
    key:"s7", label:"Professionalism, Terminology & Depth", weight:8,
    rubric:`Score 1-2: Multiple brand drug names used. Doses in mL only. Abbreviations undefined. Spelling/grammar errors that impede reading. Feels like a casual summary not a professional document.
Score 3-4: Some brand names or mL doses. Several abbreviations undefined. Minor grammar issues. Adequate but not polished.
Score 5-6: Mostly generic drug names. Doses in mg/kg. Most abbreviations spelled out on first use. Clean grammar. Professional tone.
Score 7-8: All generic drug names used (only permitted brand exceptions: Telazol, Simbadol, Nocita, Zoletil, Vetstarch, Zorbium, Zenalpha). All doses in mg or mcg/kg. All abbreviations spelled out on first use. No quoted references. Written as explanation to another veterinary professional.
Score 9-10: All of the above with exceptional depth and clinical insight. The report reads as if written by someone with deep expertise. Clinical reasoning is explicit throughout. Appropriate use of advanced terminology. The narrative demonstrates not just what happened but why it mattered clinically.`,
  },
];

// Weighted score → 0–100
function calcWeightedScore(rawScores) {
  let total = 0;
  SECTIONS.forEach((sec, i) => {
    const raw = rawScores[i];
    if (raw !== null && raw !== undefined) {
      total += (raw / 10) * sec.weight;
    }
  });
  return Math.round(total);
}

// ── Case report scoring — deep analysis, plain-text labeled output ─────────────
function buildReportSystem() {
  return `You are a senior AVTAA credentials committee reviewer with extensive experience evaluating VTS (Anesthesia and Analgesia) case reports. You are thorough, precise, and consistent.

BEFORE SCORING: Read the entire document at least twice. Identify the patient, the case complexity, and whether the technician demonstrably drove clinical decisions.

YOUR SCORING PHILOSOPHY:
- Evidence-based only: cite specific text from the report when scoring
- A score of 7 or above requires concrete, specific clinical content — not vague generalities
- A score below 5 means the section is genuinely deficient and would concern the committee
- Do NOT be lenient — the committee is strict and this person needs honest feedback to improve

MANDATORY PRE-SCORE CHECKLIST (check each before scoring):
- Is the patient ASA III, IV, or V (with or without E)? If ASA I or II, flag immediately.
- Are drug doses in mg/kg or mcg/kg? Flag every instance of mL-only dosing.
- Are lab values accompanied by reference ranges? Flag every lab value without a range.
- Is DVM approval of the anesthesia plan explicitly noted?
- Does the technician use first person and active voice showing THEY made decisions?
- Are all abbreviations defined on first use?
- Are only generic drug names used (exceptions: Telazol, Simbadol, Nocita, Zoletil, Vetstarch, Zorbium, Zenalpha)?

SECTION SCORING RUBRICS:

S1 — SIGNALMENT AND PHYSICAL STATUS (weight 10 percent)
What to look for: species, breed, age, sex, weight WITH units; complete physical exam findings relevant to anesthesia; ALL laboratory values WITH reference ranges in parentheses; full medication list; prior anesthetic history or note of none; ASA class III-V with a written paragraph explaining how each specific disease elevates anesthetic risk.

Score 1-2: Multiple required elements absent. ASA not stated or wrong class entirely.
Score 3-4: Most elements present. Labs listed but reference ranges missing from most or all. ASA stated but justification is a single vague phrase like "high risk" or "compromised patient."
Score 5-6: All basic elements present. At least some reference ranges provided but not all. ASA rating with partial explanation that mentions the disease but does not connect it to anesthetic implications.
Score 7-8: Complete, thorough physical status section. Every lab value has a reference range. Every medication listed. Prior anesthetic history addressed. ASA III-V with a clear written explanation connecting EACH specific disease to increased anesthetic risk and influencing drug selection.
Score 9-10: All of 7-8 PLUS: pertinent negatives explicitly noted, organ system review documented, each comorbidity tied to specific drug avoidances or selections with pharmacological reasoning explained.

S2 — REASON FOR ANESTHESIA (weight 5 percent)
Score 1-4: Procedure vague or DVM diagnosis not stated.
Score 5-7: Procedure and DVM diagnosis clearly and concisely stated.
Score 8-10: Procedure, diagnosis, and brief clinical context explaining WHY this case is high risk or complex — setting up why the complications section and plan are what they are.

S3 — ANTICIPATED COMPLICATIONS (weight 18 percent)
What to look for: comorbidity effects on anesthetic choices, drug-specific adverse effects, procedural complications, patient-specific risk factors — AND for every complication a specific if-then contingency plan with named drugs and doses.

Score 1-2: Bare list of complication words with no explanation. Zero contingency plans.
Score 3-4: Some complications mentioned with minimal explanation of why they are expected. No specific if-then plans — just statements like "will monitor" or "will treat if needed."
Score 5-6: Multiple complications identified with explanation of why each is anticipated. At least one if-then contingency plan present with a specific drug, but many others lacking plans.
Score 7-8: All major anticipated complications addressed: how patient comorbidities alter drug selection, expected drug adverse effects, procedural risks specific to this case. Most complications have specific contingency plans (drug name, dose, route, threshold for intervention).
Score 9-10: Every anticipated complication has a complete, dose-ready contingency plan. Proactive reasoning such as "if MAP drops below 60 mmHg I will administer dopamine at 5 mcg/kg/min IV CRI" or "if HR drops below 40 bpm I will give atropine 0.02 mg/kg IV." Drug interactions and emergency drug calculations pre-stated. Demonstrates the technician thought through every scenario before the case began.

S4 — ANESTHESIA PLAN (weight 22 percent)
What to look for: every drug with generic name and mg/kg or mcg/kg dose AND rationale; fluid type, rate (mL/kg/hr), and reasoning; full multimodal pain management for pre-, intra-, and post-operative phases; every monitoring parameter WITH specific target normal ranges; DVM review and approval explicitly documented.

Score 1-2: Drugs named with no doses and no rationale. No monitoring described. No DVM approval.
Score 3-4: Some mg/kg doses present but rationale absent or one-word. Fluid rate missing. Pain management mentions only one drug. DVM approval not noted.
Score 5-6: Most drugs have mg/kg doses and basic rationale. Fluid type and rate included. Some multimodal pain management. Monitoring listed but without target ranges. DVM approval absent or unclear.
Score 7-8: All drugs with mg/kg or mcg/kg doses AND clinical rationale explaining why each was chosen INCLUDING drugs that were intentionally avoided and why. Complete fluid plan with type, rate, and rationale. Full pre/intra/post-op pain management plan with specific drugs, doses, and timing. Every monitoring parameter listed with specific target normal ranges stated. DVM approval of the plan explicitly documented.
Score 9-10: All of 7-8 PLUS: drug interactions acknowledged, CRI calculations shown (loading dose and rate), monitoring parameters explained in terms of what information they provide and how they guide intraoperative decisions, complete pain reassessment plan with rescue analgesia protocol documented.

S5 — INTRA-OPERATIVE CARE AND PATIENT SUPPORT (weight 25 percent)
This is the most heavily weighted section. What to look for: a clear timeline with timestamps; actual measured values for every parameter (not target ranges — actual readings); each complication or event with what the TECHNICIAN did, the specific drug and dose and route, and the measurable patient response; evidence that the technician independently drove decisions.

Score 1-2: Vague narrative with no actual values. Statements like "blood pressure dropped and we treated it." No timestamps. No specific drug doses for interventions.
Score 3-4: Some values provided but inconsistently. Timeline is loose or absent. When problems arose, the response is described in passive voice or without specifics. Technician role is unclear.
Score 5-6: Reasonable timeline with some actual vital sign readings. Complications described with some management detail. Technician role partially visible but some key interventions are undocumented or vague.
Score 7-8: Clear, timestamped narrative throughout. Actual measured values documented for HR, MAP or BP, SpO2, EtCO2, and temperature at regular intervals. Each complication or deviation from plan is described with: what happened (with specific values), what the TECHNICIAN decided to do, the drug or intervention (with dose, route, and timing), and the patient response (with specific values). Technician clearly drove all decisions with active voice.
Score 9-10: Comprehensive, detailed play-by-play. Every deviation from the planned protocol is explained with clinical reasoning. Every intervention is fully documented. Technician demonstrably and independently recognized, diagnosed, and managed all complications. Shows not just what happened but why the technician made the specific choices they did. Discrepancies between plan and actual management are explained.

S6 — RECOVERY AND POST-OPERATIVE ANALGESIA (weight 12 percent)
Score 1-2: Recovery not described or one sentence. No pain assessment. No post-op analgesic plan.
Score 3-4: Brief recovery description. Pain mentioned but no validated scale used and no numeric score documented.
Score 5-6: Recovery described with some monitoring detail. Pain scale named but no actual score. Post-op analgesic drugs listed but doses or frequency incomplete or missing.
Score 7-8: Recovery described with specific monitored values. Validated pain assessment scale used (Colorado State University, Glasgow, or similar) with an actual numeric score documented at a specific time point. Complete post-op analgesic plan with generic drug names, doses in mg/kg, route, and frequency. Patient support during recovery described including positioning, warming, and oxygen supplementation. Quality of recovery assessed with specific observations.
Score 9-10: All of 7-8 PLUS: pain reassessment schedule explicitly established, rescue analgesia plan documented for inadequate pain control, any recovery complications identified and managed, handoff instructions or discharge criteria documented for the next care team.

S7 — PROFESSIONALISM, TERMINOLOGY, AND DEPTH (weight 8 percent)
Score 1-2: Brand drug names throughout. Doses in mL only. Multiple abbreviations undefined. Grammar errors that impede reading. Reads like a casual case note rather than a professional document.
Score 3-4: Some brand names or mL-only doses. Several abbreviations undefined. Minor grammar issues. Adequate professionalism but lacking polish.
Score 5-6: Mostly generic drug names. Doses in mg/kg. Most abbreviations defined on first use. Professional tone throughout. Readable and organized.
Score 7-8: All generic drug names used (permitted brand exceptions only: Telazol, Simbadol, Nocita, Zoletil, Vetstarch, Zorbium, Zenalpha). All doses in mg or mcg/kg. All abbreviations defined on first use with no exceptions. No quoted references from textbooks or journals. Written as an explanation to another veterinary professional — not as a narrative to a layperson. Excellent grammar and spelling throughout.
Score 9-10: All of 7-8 with exceptional clinical depth. The reasoning behind every decision is explicitly stated. Advanced pharmacological and physiological terminology used correctly and naturally. The document reads as if written by a technician with genuine mastery of veterinary anesthesia. Would impress even a DACVAA reviewer.


KNOWN REJECTION PATTERNS FROM THIS APPLICANT'S PREVIOUS CASE REPORTS (Dimari Diaz #2574):
These are the EXACT case report failures from her rejected application. Check every report for these:

REPORT LENGTH AND DEPTH:
- Report 2 was less than 4 pages — reports should fill close to the 5-page maximum to demonstrate adequate depth
- Reports did not demonstrate advanced anesthesia management — must show clinical decision-making not just factual recounting
- Patient history presented as factual information with no assessment — must explain HOW each finding affects the anesthetic plan

ASA CONSISTENCY (caused direct rejection):
- ASA status in case report (IV) did not match case log (III) and anesthetic sheet (III)
- All three documents must show identical ASA status — verify before submitting every report

ANESTHESIA CARE SECTION ERRORS:
- Patient presentation details were repeated in Anesthesia Care and Patient Support section instead of clinical reasoning
- This section must contain active management decisions, interventions with values, and reasoning — not a repeat of the history
- Little to no discussion of disease process and physiology — must explain how each disease alters the anesthetic approach
- No demonstration that the applicant personally assessed and made decisions — write in first person active voice

OVERALL QUALITY ISSUES:
- Reports lacked specific vital sign values and intervention details
- Missing physiological discussion of how patient diseases affected drug choices
- No evidence of advanced anesthesia management or independent clinical reasoning

RESPOND WITH EXACTLY THESE LABELED LINES — nothing before, nothing after, no extra commentary:
SCORE_S1: [integer 1-10]
SCORE_S2: [integer 1-10]
SCORE_S3: [integer 1-10]
SCORE_S4: [integer 1-10]
SCORE_S5: [integer 1-10]
SCORE_S6: [integer 1-10]
SCORE_S7: [integer 1-10]
VERDICT: [Likely Approved or Borderline or Likely Rejected]
VERDICT_REASON: [one sentence with your most important finding]
FIX_1: [most critical issue — be specific, cite the section and what exactly is missing or wrong]
FIX_2: [second most critical issue — specific and actionable]
FIX_3: [third most critical issue — specific and actionable]
STRENGTH_1: [something genuinely done well — be specific]
STRENGTH_2: [another genuine strength]
FB_S1: [detailed feedback for section 1 — cite specific content from the report, explain exactly what raises or lowers the score]
FB_S2: [detailed feedback for section 2]
FB_S3: [detailed feedback for section 3 — this is high weight, be thorough]
FB_S4: [detailed feedback for section 4 — this is high weight, be thorough]
FB_S5: [detailed feedback for section 5 — this is the highest weight, be very thorough]
FB_S6: [detailed feedback for section 6]
FB_S7: [detailed feedback for section 7]`;
}

// Bulletproof plain-text parser — regex line matching, no JSON
function parseScores(rawText) {
  const getLine = (label) => {
    const m = rawText.match(new RegExp("^" + label + ":\\s*(.+)$", "m"));
    return m ? m[1].trim() : "";
  };
  const getScore = (label) => {
    const val = getLine(label);
    const n = parseInt(val, 10);
    return (!isNaN(n) && n >= 1 && n <= 10) ? n : null;
  };
  const scores = [
    getScore("SCORE_S1"), getScore("SCORE_S2"), getScore("SCORE_S3"),
    getScore("SCORE_S4"), getScore("SCORE_S5"), getScore("SCORE_S6"), getScore("SCORE_S7"),
  ];
  const weighted = calcWeightedScore(scores);
  const verdict = getLine("VERDICT") || null;
  const verdictReason = getLine("VERDICT_REASON") || "";
  const criticalFixes = [getLine("FIX_1"), getLine("FIX_2"), getLine("FIX_3")].filter(Boolean);
  const strengths = [getLine("STRENGTH_1"), getLine("STRENGTH_2")].filter(Boolean);
  const sectionFeedback = {
    s1: getLine("FB_S1"), s2: getLine("FB_S2"), s3: getLine("FB_S3"),
    s4: getLine("FB_S4"), s5: getLine("FB_S5"), s6: getLine("FB_S6"), s7: getLine("FB_S7"),
  };
  return { scores, weighted, verdict, verdictReason, criticalFixes, strengths, sectionFeedback, raw: rawText };
}

function ScoreGauge({ score }) {
  const color = score >= 75 ? T.green : score >= 55 ? T.amber : T.red;
  const label = score >= 75 ? "Strong" : score >= 60 ? "Borderline" : score >= 45 ? "Needs Work" : "Likely Rejected";
  const circumference = 2 * Math.PI * 36;
  const dash = (score / 100) * circumference;
  return (
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
      <svg width="96" height="96" viewBox="0 0 96 96">
        <circle cx="48" cy="48" r="36" fill="none" stroke={T.border} strokeWidth="8"/>
        <circle cx="48" cy="48" r="36" fill="none" stroke={color} strokeWidth="8"
          strokeDasharray={`${dash} ${circumference}`} strokeLinecap="round"
          transform="rotate(-90 48 48)" style={{transition:"stroke-dasharray 0.8s ease"}}/>
        <text x="48" y="44" textAnchor="middle" fontSize="18" fontWeight="900" fill={color} fontFamily="DM Sans">{score}</text>
        <text x="48" y="58" textAnchor="middle" fontSize="9" fill={T.muted} fontFamily="DM Sans">/100</text>
      </svg>
      <div style={{fontSize:11,fontWeight:700,color,letterSpacing:0.5}}>{label}</div>
    </div>
  );
}

function SectionBar({ label, score, weight, feedback }) {
  const [open, setOpen] = useState(false);
  const pct = score !== null ? (score / 10) * 100 : 0;
  const color = score >= 8 ? "green" : score >= 6 ? "amber" : "red";
  return (
    <div style={{marginBottom:10}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3,cursor:feedback?"pointer":"default"}} onClick={()=>feedback&&setOpen(o=>!o)}>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:12,color:T.sub}}>{label}</span>
          <span style={{fontSize:10,color:T.muted,background:T.panel,padding:"1px 6px",borderRadius:10}}>{weight}%</span>
          {feedback && <span style={{fontSize:10,color:T.muted}}>{open?"▲":"▼"}</span>}
        </div>
        <span style={{fontSize:13,fontWeight:800,color:color==="green"?T.green:color==="amber"?T.amber:T.red,minWidth:36,textAlign:"right"}}>
          {score !== null ? `${score}/10` : "—"}
        </span>
      </div>
      <div style={S.scoreBar}><div style={S.scoreFill(pct, color)}/></div>
      {open && feedback && (
        <div style={{marginTop:8,padding:"10px 12px",background:"#0e0c18",border:`1px solid ${T.border}`,borderLeft:`3px solid ${color==="green"?T.green:color==="amber"?T.amber:T.red}`,borderRadius:"0 8px 8px 0",fontSize:12,color:T.sub,lineHeight:1.65}}>
          {feedback}
        </div>
      )}
    </div>
  );
}

function CaseReportTab({reportScores, setReportScores, setReportUploadCount}) {
  // All state per report in one object — no stale closures from split updates
  const blank = (id) => ({
    id, label:`Case Report ${id}`,
    file:null,
    phase:"idle",    // idle | ready | running | done | error
    step:"",         // human-readable current step shown in UI
    text:"",
    formatResults:null,
    parsed:null,
    errMsg:"",
    showReplace:false,
  });
  const [reports, setReports] = useState([1,2,3,4].map(blank));

  // Atomic patch — always uses functional updater so closures never go stale
  const patch = (id, obj) =>
    setReports(prev => prev.map(r => r.id===id ? {...r,...obj} : r));

  // Step 1: file drop — just register it, show the button
  const onFile = (id, file) => {
    patch(id, {...blank(id), file, phase:"ready"});
    // Update upload count in parent — count all reports that have a file
    setReports(prev => {
      const updated = prev.map(r => r.id===id ? {...blank(id), file, phase:"ready"} : r);
      setReportUploadCount(updated.filter(r=>r.file||r.phase==="done").length);
      return updated;
    });
  };

  // Step 2: single button — read file, check format, score content, show everything
  const run = async (id) => {
    const file = reports.find(r=>r.id===id)?.file;
    if (!file) return;

    patch(id, {phase:"running", step:"Reading file...", parsed:null, formatResults:null, errMsg:"", text:""});

    try {
      // ── Read text + format check ────────────────────────────────────────────
      let text = "";
      let formatResults = null;

      if (file.name.endsWith(".docx")) {
        patch(id, {step:"Checking font, spacing and margins..."});
        // checkDocxFormatting clones the buffer internally — safe to call first
        const buf = await file.arrayBuffer();
        // Pass a fresh clone to format checker; keep another for mammoth
        formatResults = await checkDocxFormatting(buf.slice(0));
        const mm = await mammoth.extractRawText({arrayBuffer: buf.slice(0)});
        text = mm.value || "";

      } else if (file.name.endsWith(".pdf")) {
        patch(id, {step:"Extracting text from PDF..."});
        text = await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onerror = () => reject(new Error("Could not read PDF file"));
          reader.onload = async (e) => {
            try {
              const base64 = e.target.result.split(",")[1];
              const res = await fetch("https://api.anthropic.com/v1/messages", {
                method:"POST", headers:{"Content-Type":"application/json"},
                body: JSON.stringify({
                  model:"claude-sonnet-4-20250514", max_tokens:1200,
                  messages:[{role:"user", content:[
                    {type:"document", source:{type:"base64", media_type:"application/pdf", data:base64}},
                    {type:"text", text:"Extract ALL text from this document exactly as written. Return only the document text, preserving section headers."}
                  ]}]
                })
              });
              const d = await res.json();
              resolve(d.content?.map(b=>b.text||"").join("") || "");
            } catch(e) { reject(e); }
          };
          reader.readAsDataURL(file);
        });

      } else {
        text = await file.text();
      }

      if (!text.trim()) throw new Error("No text extracted. Try re-saving the file as .docx or .txt.");

      // ── AI scoring ──────────────────────────────────────────────────────────
      patch(id, {step:"Scoring against AVTAA rubric — this takes about 15 seconds...", text, formatResults});

      const raw = await callClaude(
        buildReportSystem(),
        `You are performing a detailed AVTAA VTS case report scoring. Read every word carefully before scoring. Respond with ONLY the labeled lines as specified:\n\n${text.slice(0,12000)}`,
        2200
      );

      const parsed = parseScores(raw);

      // Update the shared readiness tracker
      setReportScores(prev => {
        const ns = [...prev];
        ns[id-1] = parsed.weighted;
        return ns;
      });

      patch(id, {phase:"done", step:"", parsed, text, formatResults});
      // Recount uploads after scoring
      setReports(prev => { setReportUploadCount(prev.filter(r=>r.file||r.phase==="done").length); return prev; });

    } catch(err) {
      patch(id, {phase:"error", step:"", errMsg: err.message || "Something went wrong"});
    }
  };

  const [activeReport, setActiveReport] = useState(null); // null = show all, id = scroll to / highlight

  return (
    <div>
      <div style={S.sectionTitle}>Case Report Reviewer</div>
      <div style={S.sectionSub}>
        Drop a file → hit <strong style={{color:T.violetMid}}>Analyze &amp; Score</strong>.
        One button runs the format check and AI scoring together and shows you both results when complete.
      </div>

      {/* Quick-access dropdown — always visible */}
      <div style={{...S.card, padding:"14px 20px", marginBottom:20, display:"flex", alignItems:"center", gap:14, flexWrap:"wrap"}}>
        <div style={{fontSize:12, color:T.muted, fontWeight:700, whiteSpace:"nowrap"}}>🔖 Jump to Report:</div>
        <select
          style={{...S.input, flex:1, cursor:"pointer", minWidth:200}}
          value={activeReport||""}
          onChange={e => {
            const id = parseInt(e.target.value);
            setActiveReport(id||null);
            if (id) {
              setTimeout(()=>{
                document.getElementById(`case-report-card-${id}`)?.scrollIntoView({behavior:"smooth", block:"start"});
              }, 50);
            }
          }}
        >
          <option value="">— Select a report —</option>
          {reports.map(r=>{
            const p = r.parsed;
            const scored = r.phase==="done" && p;
            const verdict = p?.verdict?.includes("Approved")?"✅":p?.verdict?.includes("Borderline")?"⚠️":scored?"❌":"";
            return (
              <option key={r.id} value={r.id}>
                {scored
                  ? `${verdict} ${r.label} · ${p?.weighted??0}/100 · ${p?.verdict||"—"} · ${r.file?.name||""}`
                  : `${r.label}${r.phase==="idle"?" · Not yet uploaded":r.phase==="ready"?" · Ready to score":r.phase==="running"?" · Analyzing...":""}`
                }
              </option>
            );
          })}
        </select>
        <div style={{fontSize:11, color:T.muted}}>
          {reports.filter(r=>r.phase==="done"&&r.parsed).length} of 4 scored
        </div>
      </div>

      <div style={{...S.card, background:T.amberSoft, border:`1px solid ${T.amber}44`, marginBottom:24}}>
        <div style={{fontSize:13, color:T.amber, fontWeight:700, marginBottom:6}}>⚡ Auto-Rejected If Format Wrong</div>
        <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:"4px 20px", fontSize:12, color:T.sub}}>
          <div>📝 Font: Times New Roman ONLY</div><div>📐 Margins: 0.5" all sides</div>
          <div>📏 Size: 10.5pt</div><div>📄 Max: 5 pages</div>
          <div>↕️ Spacing: 1.5 lines</div><div>💾 Submit as .docx (not .pdf)</div>
        </div>
      </div>

      <div style={{...S.card, padding:"12px 20px", marginBottom:20, display:"flex", gap:20, flexWrap:"wrap", alignItems:"center"}}>
        <div style={{fontSize:11, color:T.muted, fontWeight:700, letterSpacing:0.5}}>SCORE GUIDE</div>
        {[["75–100","Strong — likely approved",T.green],["55–74","Borderline",T.amber],["0–54","Likely rejected",T.red]].map(([range,lbl,c])=>(
          <div key={range} style={{display:"flex", alignItems:"center", gap:6, fontSize:12}}>
            <div style={{width:9, height:9, borderRadius:"50%", background:c, flexShrink:0}}/>
            <span style={{color:c, fontWeight:700}}>{range}</span>
            <span style={{color:T.muted}}>{lbl}</span>
          </div>
        ))}
        <div style={{fontSize:11, color:T.muted, marginLeft:"auto"}}>Tap section bars to expand feedback</div>
      </div>

      {reports.map(rpt => {
        const p = rpt.parsed;
        const vc = p?.verdict?.includes("Approved")?"green":p?.verdict?.includes("Borderline")?"amber":p?.verdict?"red":"violet";
        const running = rpt.phase==="running";
        const done    = rpt.phase==="done";
        const errored = rpt.phase==="error";
        const ready   = rpt.phase==="ready";

        return (
          <div key={rpt.id} id={`case-report-card-${rpt.id}`}
          style={{...S.card, marginBottom:20,
            outline: activeReport===rpt.id ? `2px solid ${T.violet}` : "none",
            transition:"outline 0.2s"}}>

            {/* Header */}
            <div style={{display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:16, flexWrap:"wrap", gap:12}}>
              <div>
                <div style={{fontWeight:800, fontSize:16, marginBottom:4}}>{rpt.label}</div>
                {p?.verdict && <span style={S.chip(vc)}>{p.verdict}</span>}
                {p?.verdictReason && <div style={{fontSize:12, color:T.muted, marginTop:6, maxWidth:500, lineHeight:1.5}}>{p.verdictReason}</div>}
              </div>
              {done && p && <ScoreGauge score={p.weighted}/>}
            </div>

            {/* Section score bars */}
            {done && p?.scores?.length > 0 && (
              <div style={{marginBottom:16}}>
                <div style={{fontSize:11, color:T.muted, fontWeight:700, letterSpacing:0.5, marginBottom:10}}>
                  SECTION BREAKDOWN — tap any bar to expand feedback
                </div>
                {SECTIONS.map((sec,i) => (
                  <SectionBar key={sec.key} label={sec.label} score={p.scores[i]??null} weight={sec.weight} feedback={p.sectionFeedback?.[sec.key]}/>
                ))}
              </div>
            )}

            {/* Critical fixes + Strengths */}
            {done && (p?.criticalFixes?.length>0 || p?.strengths?.length>0) && (
              <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:16}}>
                <div style={{background:T.redSoft, border:`1px solid ${T.red}44`, borderRadius:10, padding:14}}>
                  <div style={{fontSize:11, color:T.red, fontWeight:700, marginBottom:8, letterSpacing:0.5}}>🔴 CRITICAL FIXES</div>
                  {p.criticalFixes?.length > 0
                    ? p.criticalFixes.map((f,i)=>(
                        <div key={i} style={{fontSize:12, color:T.sub, marginBottom:6, paddingLeft:8, borderLeft:`2px solid ${T.red}55`, lineHeight:1.5}}>{f}</div>
                      ))
                    : <div style={{fontSize:12, color:T.muted, fontStyle:"italic"}}>No critical issues found</div>
                  }
                </div>
                <div style={{background:T.greenSoft, border:`1px solid ${T.green}44`, borderRadius:10, padding:14}}>
                  <div style={{fontSize:11, color:T.green, fontWeight:700, marginBottom:8, letterSpacing:0.5}}>✅ STRENGTHS</div>
                  {p.strengths?.length > 0
                    ? p.strengths.map((s,i)=>(
                        <div key={i} style={{fontSize:12, color:T.sub, marginBottom:6, paddingLeft:8, borderLeft:`2px solid ${T.green}55`, lineHeight:1.5}}>{s}</div>
                      ))
                    : <div style={{fontSize:12, color:T.muted, fontStyle:"italic"}}>None noted</div>
                  }
                </div>
              </div>
            )}

            {/* Format compliance (shown after done, only for .docx) */}
            {done && rpt.formatResults?.passes && (
              <div style={{marginBottom:16}}>
                <div style={{fontSize:11, color:T.muted, fontWeight:700, letterSpacing:0.5, marginBottom:8}}>FORMAT COMPLIANCE</div>
                <FormatResults results={rpt.formatResults.passes}/>
              </div>
            )}

            {/* File area — collapsed pill when done, full drop zone otherwise */}
            {done ? (
              <div style={{marginTop:16}}>
                {/* Current file pill */}
                <div style={{display:"flex", alignItems:"center", justifyContent:"space-between", padding:"10px 14px", background:T.panel, border:`1px solid ${T.border}`, borderRadius:10, marginBottom: rpt.showReplace ? 10 : 0}}>
                  <div style={{display:"flex", alignItems:"center", gap:10}}>
                    <span style={{fontSize:16}}>📄</span>
                    <div>
                      <div style={{fontSize:13, fontWeight:600, color:T.green}}>{rpt.file?.name}</div>
                      <div style={{fontSize:11, color:T.muted}}>{rpt.text?.split(/\s+/).filter(Boolean).length.toLocaleString()} words scored</div>
                    </div>
                  </div>
                  <button
                    style={{...S.btn("ghost"), fontSize:11, padding:"5px 12px", border:`1px solid ${T.border}`, color:T.sub}}
                    onClick={() => patch(rpt.id, {showReplace: !rpt.showReplace})}
                  >
                    {rpt.showReplace ? "✕ Cancel" : "📂 Upload New File"}
                  </button>
                </div>

                {/* New file drop zone — only shown when she explicitly clicks Upload New */}
                {rpt.showReplace && (
                  <div style={{border:`1px solid ${T.amber}44`, borderRadius:10, padding:14, background:T.amberSoft}}>
                    <div style={{fontSize:12, color:T.amber, fontWeight:600, marginBottom:10}}>
                      ⚠️ Uploading a new file will replace the current results for {rpt.label}
                    </div>
                    <DropZone
                      accept=".docx,.pdf,.txt"
                      label="Drop new file here"
                      sublabel=".docx preferred · drag & drop or click"
                      onFile={f => { onFile(rpt.id, f); patch(rpt.id, {showReplace: false}); }}
                      fileName={null}
                    />
                  </div>
                )}
              </div>
            ) : (
              <div>
                <label style={{...S.label, marginTop:0}}>Upload Case Report</label>
                <DropZone
                  accept=".docx,.pdf,.txt"
                  label={`Drop ${rpt.label} here`}
                  sublabel="Best: .docx (required by AVTAA) · Also accepts .pdf or .txt"
                  onFile={f => onFile(rpt.id, f)}
                  fileName={rpt.file?.name}
                />
              </div>
            )}

            {/* Non-docx warning */}
            {rpt.file && !rpt.file.name.endsWith(".docx") && !running && (
              <div style={{marginTop:8, padding:"7px 12px", background:T.amberSoft, border:`1px solid ${T.amber}44`, borderRadius:8, fontSize:12, color:T.amber}}>
                ⚠️ Content will be reviewed — but AVTAA requires .docx for actual submission.
              </div>
            )}

            {/* Progress indicator */}
            {running && (
              <div style={{marginTop:14, padding:"13px 16px", background:T.violetSoft, border:`1px solid ${T.violet}44`, borderRadius:10, display:"flex", alignItems:"center", gap:10}}>
                <span style={S.spinner}/>
                <span style={{fontSize:13, color:T.violetMid, fontWeight:600}}>{rpt.step}</span>
              </div>
            )}

            {/* Error */}
            {errored && (
              <div style={{marginTop:14, padding:"13px 16px", background:T.redSoft, border:`1px solid ${T.red}44`, borderRadius:10}}>
                <div style={{fontSize:13, color:T.red, marginBottom:8}}>❌ {rpt.errMsg}</div>
                <button style={{...S.btn("ghost"), fontSize:12, padding:"5px 12px"}}
                  onClick={() => patch(rpt.id, {phase:"ready", errMsg:""})}>
                  Try Again
                </button>
              </div>
            )}

            {/* THE single action button */}
            {(ready || done) && !running && (
              <div style={{marginTop:14, display:"flex", alignItems:"center", gap:14, flexWrap:"wrap"}}>
                <button style={{...S.btn("primary"), padding:"12px 28px", fontSize:14}}
                  onClick={() => run(rpt.id)}>
                  {done ? "🔄 Re-Analyze & Re-Score" : "🎯 Analyze & Score"}
                </button>
                {done && rpt.text && (
                  <span style={{fontSize:12, color:T.muted}}>
                    {rpt.text.split(/\s+/).filter(Boolean).length.toLocaleString()} words scored
                  </span>
                )}
              </div>
            )}

          </div>
        );
      })}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// TAB 6 — CASE REPORT BUILDER
// ═══════════════════════════════════════════════════════════════════════════════
const BUILDER_SECTIONS = [
  {
    id:"signalment", label:"1. Patient Information & Physical Status",
    hint:"AVTAA needs: species, breed, age, sex, weight, key physical exam findings, all lab values WITH reference ranges, current medications, past anesthetic history, and your ASA rating (III–V) with a written explanation of why.",
    prompts:["What species, breed, age, sex, and weight is the patient?","What were the key physical exam findings relevant to anesthesia?","What lab values did you have? List them with normal ranges in parentheses.","What medications is the patient currently on?","Any prior anesthetic complications or relevant history?","What ASA status did you assign (must be III, IV, or V) and why?"],
    placeholder:"Example: 8yr MN Golden Retriever, 32kg. PE: grade III/VI left systolic heart murmur, mild tachycardia HR 128bpm, mild exercise intolerance reported by owner. CBC/Chem WNL except BNP elevated at 320pmol/L (ref: <200). Current medications: enalapril 5mg PO BID, furosemide 20mg PO BID. No prior anesthetic complications. ASA IV — uncompensated cardiac disease with clinical signs affecting daily function.",
  },
  {
    id:"reason", label:"2. Reason for Anesthesia",
    hint:"State the procedure clearly. Include the diagnosis made by the DVM. Keep it concise.",
    prompts:["What procedure was performed?","What was the DVM's diagnosis?"],
    placeholder:"Example: General anesthesia for right lateral thoracotomy and pericardial window procedure. Diagnosis: pericardial effusion with early cardiac tamponade secondary to suspected idiopathic cause vs. neoplasia.",
  },
  {
    id:"anticipated", label:"3. Anticipated Complications",
    hint:"This is a high-scoring section. AVTAA wants to see that YOU thought ahead. Cover: (a) how the patient's diseases will affect anesthesia, (b) drug-related complications you anticipated, (c) procedural complications, (d) your contingency plan for each.",
    prompts:["How will the patient's cardiac condition affect your anesthetic choices?","What complications from your chosen drugs did you anticipate?","What procedural complications were possible?","What was your contingency plan if hypotension occurred?","What was your contingency plan if arrhythmias occurred?"],
    placeholder:"Example: Anticipated hypotension from inhalant anesthesia — planned to keep isoflurane at lowest effective concentration, have dopamine CRI prepared at 5mcg/kg/min. Anticipated arrhythmias given existing cardiac disease — ECG monitoring throughout, lidocaine bolus 2mg/kg IV prepared. Open chest will eliminate normal intrathoracic pressure — planned IPPV throughout. Drug-related: acepromazine avoided due to vasodilation risk in cardiac patient...",
  },
  {
    id:"plan", label:"4. Anesthesia Plan",
    hint:"State EXACTLY what you planned BEFORE the case. Include every drug with mg/kg dose and rationale, fluid plan, full pain management strategy (pre/intra/post-op), monitoring equipment with normal ranges, and note that the DVM approved the plan.",
    prompts:["What premeds did you plan and what dose (mg/kg)? Why those drugs?","What induction agent and dose (mg/kg)?","What maintenance agent and expected %?","What fluid type and rate (mL/kg/hr)?","What was your pain management plan for pre-op, during, and after?","What monitoring equipment and what were your target normal ranges?","Did the overseeing DVM approve this plan? Any changes?"],
    placeholder:"Example: Premed: hydromorphone 0.1mg/kg IV for pre-op analgesia — opioid chosen for minimal cardiac effects, avoidance of acepromazine due to vasodilatory risk in compensated cardiac patient. Induction: propofol titrated IV to effect (approximately 2-4mg/kg) — chosen for rapid onset and smooth induction with ability to titrate... Monitoring: ECG (target NSR, HR 70-120bpm), direct arterial BP (MAP target >65mmHg), SpO2 (target >95%), capnography (EtCO2 target 35-45mmHg), esophageal temperature (target >36°C)... DVM approved plan, addition of dobutamine as backup if dopamine insufficient.",
  },
  {
    id:"intraop", label:"5. What Actually Happened During Anesthesia",
    hint:"Walk through the ACTUAL case with a timeline. Include real vital signs, any deviations from your plan, problems that came up, and exactly what YOU did to manage them. This is where you show you drove the case.",
    prompts:["Walk through the induction — what happened, what were the vitals?","What happened during maintenance? Any complications?","If vitals went abnormal — what exactly happened and what did YOU do to fix it?","Were there any deviations from your original plan? Why?","Provide a timeline with actual times and vital sign values."],
    placeholder:"Example: 08:15 — premed administered, patient mildly sedated, IV catheter placed without difficulty. 08:45 — propofol 2.1mg/kg IV administered over 60 sec, smooth induction, intubated with 10mm ETT on first attempt, cuff inflated to 20cmH2O. 08:50 — isoflurane initiated at 2%, oxygen flow 2L/min, IPPV initiated at 10bpm tidal volume 10mL/kg. 09:00 — MAP dropped to 52mmHg (target >65), HR 118bpm NSR. Isoflurane reduced to 1.3%, dopamine CRI started at 5mcg/kg/min. 09:08 — MAP recovered to 71mmHg...",
  },
  {
    id:"recovery", label:"6. Recovery & Post-Anesthesia Care",
    hint:"Describe how you managed the recovery. Include pain scoring method, post-op analgesic plan, how you supported the patient, quality of recovery, and any complications.",
    prompts:["How did you assess pain post-op? What scale?","What post-op analgesic plan did you implement?","How did you support the patient through recovery (temperature, positioning, monitoring)?","What was the quality of recovery? Any complications?"],
    placeholder:"Example: Patient transferred to recovery with oxygen supplementation. Pain assessed using Colorado State University Acute Pain Scale — score 2/4 at 30 min post-op. Hydromorphone 0.05mg/kg IV administered for pain management. Patient maintained in sternal recumbency on warm water blanket, temperature monitored q15min. Smooth recovery, extubated at swallowing reflex, HR 88bpm, MAP 74mmHg on final check. Discharged to ICU team with pain reassessment order q4h and fentanyl CRI instructions at 2mcg/kg/hr.",
  },
  {
    id:"reflection", label:"7. Case Reflection (Optional — only if under 5 pages)",
    hint:"Optional section. Only include if you have room under the 5-page limit. Briefly reflect: what would you do differently, what did you learn?",
    prompts:["Is there anything you would do differently next time with a similar patient?","What was the most valuable thing you learned from this case?"],
    placeholder:"Example: In retrospect, initiating the dopamine CRI prophylactically at induction rather than reactively may have prevented the transient hypotensive episode. I would also consider an arterial catheter placed pre-induction for continuous blood pressure monitoring during induction in future high-risk cardiac patients.",
  },
];

const BUILD_SYSTEM=`You are an AVTAA VTS (Anesthesia & Analgesia) application expert helping a veterinary technician write a case report section.

Your job is to take their raw clinical notes for one section and turn it into polished, professional prose that:
- Meets AVTAA case report standards for that specific section
- Uses proper medical terminology throughout
- Uses generic drug names only (except Telazol, Simbadol, Nocita, Zoletil, Vetstarch, Zorbium, Zenalpha)
- Lists all doses as mg or mcg/kg (never mL)
- Shows the TECHNICIAN drove decisions and reasoning, not just recorded facts
- Spells out all abbreviations on first use
- Is written as if explaining to another veterinary professional
- Does NOT quote references
- Flows naturally and professionally

Return ONLY the polished prose for that section — no meta-commentary, no headers, just the text ready to paste into the Word document.`;

function CaseReportBuilderTab({ builderSections, setBuilderSections }) {
  // builderSections lifted to App for persistence
  const [activeSection, setActiveSection] = useState("signalment");
  const [showPreview, setShowPreview]     = useState(false);
  const [downloading, setDownloading]     = useState(false);

  const update = (id, field, val) =>
    setBuilderSections(p => ({...p, [id]: {...p[id], [field]: val}}));

  const current = BUILDER_SECTIONS.find(s => s.id === activeSection);
  const sec     = builderSections[activeSection];

  const polish = async () => {
    if (!sec.notes.trim()) return;
    update(activeSection, "loading", true);
    const r = await callClaude(
      BUILD_SYSTEM,
      `Section: ${current.label}\nAVTAA requirements: ${current.hint}\n\nRaw clinical notes:\n${sec.notes}`,
      1000
    );
    update(activeSection, "polished", r);
    update(activeSection, "loading", false);
  };

  const completedSections = BUILDER_SECTIONS.filter(s => builderSections[s.id]?.polished);
  const completedCount    = completedSections.length;
  const fullDraft         = completedSections
    .map(s => `${s.label.replace(/^\d+\.\s*/, "")}\n\n${builderSections[s.id].polished}`)
    .join("\n\n---\n\n");

  // ── Download as .docx using raw Open XML ──────────────────────────────────
  const downloadDocx = async () => {
    if (!fullDraft) return;
    setDownloading(true);
    try {
      const JSZip = await loadJSZip();

      // Build Open XML document with AVTAA-required formatting:
      // Font: Times New Roman 10.5pt (21 half-pts)
      // Spacing: 1.5 lines (360 twips)
      // Margins: 0.5" all sides (720 twips)
      const escXml = (s) => s
        .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
        .replace(/"/g,"&quot;").replace(/'/g,"&apos;");

      const sectionXml = completedSections.map(s => {
        const heading = escXml(s.label.replace(/^\d+\.\s*/, "").toUpperCase());
        const body    = builderSections[s.id].polished
          .split(/\n+/)
          .filter(Boolean)
          .map(line => `
          <w:p>
            <w:pPr>
              <w:spacing w:line="360" w:lineRule="auto" w:before="0" w:after="160"/>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
                <w:sz w:val="21"/><w:szCs w:val="21"/>
              </w:rPr>
              <w:t xml:space="preserve">${escXml(line)}</w:t>
            </w:r>
          </w:p>`)
          .join("\n");
        return `
        <w:p>
          <w:pPr>
            <w:spacing w:line="360" w:lineRule="auto" w:before="200" w:after="80"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
              <w:sz w:val="21"/><w:szCs w:val="21"/>
              <w:b/><w:color w:val="2D1F5E"/>
            </w:rPr>
            <w:t>${heading}</w:t>
          </w:r>
        </w:p>
        ${body}`;
      }).join("\n");

      const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  mc:Ignorable="w14 wp14">
<w:body>
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:line="360" w:lineRule="auto" w:before="0" w:after="200"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="28"/><w:szCs w:val="28"/><w:b/>
      </w:rPr>
      <w:t>AVTAA VTS Case Report</w:t>
    </w:r>
  </w:p>
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:line="360" w:lineRule="auto" w:before="0" w:after="360"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="21"/><w:szCs w:val="21"/><w:color w:val="6B5F8A"/>
      </w:rPr>
      <w:t>Draft — Built with VTS Compass</w:t>
    </w:r>
  </w:p>
  ${sectionXml}
  <w:sectPr>
    <w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"
             w:header="708" w:footer="708" w:gutter="0"/>
  </w:sectPr>
</w:body>
</w:document>`;

      const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

      const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="21"/><w:szCs w:val="21"/>
        <w:lang w:val="en-US"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:line="360" w:lineRule="auto" w:before="0" w:after="160"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
</w:styles>`;

      const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

      const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

      const zip = new JSZip();
      zip.file("[Content_Types].xml", contentTypes);
      zip.file("_rels/.rels", rootRels);
      zip.file("word/document.xml", docXml);
      zip.file("word/styles.xml", stylesXml);
      zip.file("word/_rels/document.xml.rels", relsXml);

      const blob = await zip.generateAsync({type:"blob", mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement("a");
      a.href     = url;
      a.download = "vts-case-report-draft.docx";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch(err) {
      alert("Download failed: " + err.message);
    }
    setDownloading(false);
  };

  return (
    <div>
      <div style={S.sectionTitle}>Case Report Builder</div>
      <div style={S.sectionSub}>Fill in raw clinical notes section by section — Claude polishes them into AVTAA-ready prose. All notes auto-save so you can pick up where you left off.</div>

      {/* Section tabs */}
      <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:20}}>
        {BUILDER_SECTIONS.map(s=>{
          const done   = !!builderSections[s.id]?.polished;
          const active = activeSection === s.id;
          return (
            <button key={s.id} onClick={()=>setActiveSection(s.id)} style={{padding:"8px 14px",borderRadius:8,border:`1px solid ${active?T.violet:done?T.green:T.border}`,background:active?T.violetSoft:done?T.greenSoft:"transparent",color:active?T.violetMid:done?T.green:T.muted,fontSize:12,fontWeight:700,cursor:"pointer",transition:"all 0.15s"}}>
              {done?"✓ ":""}{s.label.replace(/^\d+\.\s*/,"").split("&")[0].trim()}
            </button>
          );
        })}
      </div>

      {/* Progress + action bar */}
      {completedCount > 0 && (
        <div style={{...S.card,background:T.greenSoft,border:`1px solid ${T.green}44`,marginBottom:20,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
          <div>
            <div style={{fontSize:13,color:T.green,fontWeight:700,marginBottom:2}}>{completedCount} of {BUILDER_SECTIONS.length} sections completed</div>
            <div style={{fontSize:11,color:T.sub}}>All notes saved automatically — closes and reopens where you left off</div>
          </div>
          <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
            <button style={{...S.btn("ghost"),fontSize:12,padding:"8px 16px"}} onClick={()=>setShowPreview(true)}>
              👁 Preview Full Draft
            </button>
            <button style={{...S.btn("teal"),fontSize:12,padding:"8px 16px"}} onClick={downloadDocx} disabled={downloading}>
              {downloading?<><span style={S.spinner}/>Building .docx...</>:"⬇️ Download as .docx"}
            </button>
          </div>
        </div>
      )}

      {/* Two-column editor */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
        <div>
          <div style={S.card}>
            <div style={{fontWeight:700,fontSize:15,marginBottom:4}}>{current.label}</div>
            <div style={{fontSize:12,color:T.muted,marginBottom:14,lineHeight:1.6}}>{current.hint}</div>
            <div style={{...S.card,background:T.violetSoft,border:`1px solid ${T.violet}44`,marginBottom:16,padding:14}}>
              <div style={{fontSize:11,color:T.violetMid,fontWeight:700,marginBottom:8,letterSpacing:0.5}}>NOTATION PROMPTS</div>
              {current.prompts.map((p,i)=>(
                <div key={i} style={{fontSize:12,color:T.sub,marginBottom:6,paddingLeft:8,borderLeft:`2px solid ${T.violet}44`,lineHeight:1.5}}>{i+1}. {p}</div>
              ))}
            </div>
            <label style={S.label}>Your Raw Clinical Notes</label>
            <textarea
              style={{...S.textarea,minHeight:220}}
              placeholder={current.placeholder}
              value={sec?.notes||""}
              onChange={e=>update(activeSection,"notes",e.target.value)}
            />
            <div style={{marginTop:6,fontSize:11,color:T.muted}}>{(sec?.notes||"").length} characters</div>
            <div style={{marginTop:14}}>
              <button style={S.btn("primary")} onClick={polish} disabled={sec?.loading||!sec?.notes?.trim()}>
                {sec?.loading?<><span style={S.spinner}/>Writing Section...</>:"✨ Polish This Section"}
              </button>
            </div>
          </div>
        </div>

        <div>
          <div style={S.card}>
            <div style={{fontWeight:700,fontSize:15,marginBottom:4}}>Polished Output</div>
            <div style={{fontSize:12,color:T.muted,marginBottom:14}}>Copy this into your Word document, or use the buttons below to preview or download the full draft.</div>
            {sec?.polished?(
              <>
                <div style={{background:"#0e0c18",border:`1px solid ${T.green}44`,borderRadius:10,padding:16,fontSize:13,lineHeight:1.75,color:T.text,whiteSpace:"pre-wrap",minHeight:200,maxHeight:400,overflowY:"auto"}}>{sec.polished}</div>
                <button style={{...S.btn("ghost"),marginTop:10,fontSize:12}} onClick={()=>navigator.clipboard?.writeText(sec.polished)}>📋 Copy Section</button>
              </>
            ):(
              <div style={{background:"#0e0c18",border:`1px solid ${T.border}`,borderRadius:10,padding:16,minHeight:200,display:"flex",alignItems:"center",justifyContent:"center",color:T.muted,fontSize:13}}>
                {sec?.loading?<><span style={S.spinner}/>Writing your polished section...</>:"Fill in your notes and click Polish"}
              </div>
            )}

            {/* Always-visible Preview + Download row, centered under output */}
            <div style={{marginTop:18,paddingTop:16,borderTop:`1px solid ${T.border}`,display:"flex",gap:10,justifyContent:"center",flexWrap:"wrap"}}>
              <button
                style={{...S.btn("ghost"),fontSize:13,padding:"10px 22px",border:`1px solid ${T.violet}55`,color:T.violetMid}}
                onClick={()=>setShowPreview(true)}
              >
                👁 Preview Full Draft
              </button>
              <button
                style={{...S.btn("teal"),fontSize:13,padding:"10px 22px"}}
                onClick={downloadDocx}
                disabled={downloading||completedCount===0}
                title={completedCount===0?"Polish at least one section first":"Download full draft as .docx"}
              >
                {downloading?<><span style={S.spinner}/>Building...</>:"⬇️ Download .docx"}
              </button>
            </div>
            {completedCount===0&&(
              <div style={{textAlign:"center",fontSize:11,color:T.muted,marginTop:8}}>
                Polish at least one section to enable download
              </div>
            )}
          </div>

          {/* Download detail card — only when sections completed */}
          {completedCount > 0 && (
            <div style={{...S.card,background:T.tealSoft,border:`1px solid ${T.teal}44`,marginTop:0}}>
              <div style={{fontSize:12,color:T.sub,lineHeight:1.6}}>
                <strong style={{color:T.teal}}>✓ {completedCount} section{completedCount!==1?"s":""} ready.</strong> Downloads as .docx with Times New Roman 10.5pt, 1.5 spacing, 0.5" margins — AVTAA-compliant formatting applied automatically.
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Preview modal */}
      {showPreview && (
        <div style={{position:"fixed",inset:0,zIndex:900,background:"rgba(0,0,0,0.8)",backdropFilter:"blur(4px)",display:"flex",alignItems:"center",justifyContent:"center",padding:20}} onClick={e=>{if(e.target===e.currentTarget)setShowPreview(false);}}>
          <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:20,padding:32,maxWidth:760,width:"100%",maxHeight:"85vh",display:"flex",flexDirection:"column",gap:16}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <div style={{fontWeight:800,fontSize:18,color:T.text}}>Full Draft Preview</div>
                <div style={{fontSize:12,color:T.muted,marginTop:3}}>{completedCount} section{completedCount!==1?"s":""} · Simulates Times New Roman 10.5pt formatting</div>
              </div>
              <div style={{display:"flex",gap:10}}>
                <button style={{...S.btn("teal"),fontSize:12}} onClick={()=>{setShowPreview(false);downloadDocx();}}>⬇️ Download .docx</button>
                <button style={{...S.btn("ghost"),fontSize:13,padding:"8px 14px"}} onClick={()=>setShowPreview(false)}>✕ Close</button>
              </div>
            </div>
            <div style={{overflowY:"auto",flex:1,background:"#fff",borderRadius:12,padding:"48px 52px",color:"#1a1a1a",fontFamily:"'Times New Roman',Times,serif",fontSize:"10.5pt",lineHeight:1.5}}>
              <div style={{textAlign:"center",marginBottom:32}}>
                <div style={{fontSize:16,fontWeight:700,marginBottom:4}}>AVTAA VTS Case Report</div>
                <div style={{fontSize:11,color:"#888"}}>Draft — Built with VTS Compass</div>
              </div>
              {completedSections.map(s=>(
                <div key={s.id} style={{marginBottom:28}}>
                  <div style={{fontWeight:700,fontSize:"10.5pt",marginBottom:10,textTransform:"uppercase",borderBottom:"1px solid #ddd",paddingBottom:6,color:"#2D1F5E"}}>
                    {s.label.replace(/^\d+\.\s*/,"")}
                  </div>
                  <div style={{whiteSpace:"pre-wrap",lineHeight:1.5,fontSize:"10.5pt"}}>
                    {builderSections[s.id].polished}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// TAB 7 — SKILLS MAP
// ═══════════════════════════════════════════════════════════════════════════════
const CORE_SKILLS=["Pre-anesthetic patient assessment","Physical examination for anesthesia planning","Interpretation of diagnostic/lab values for ASA rating","IV catheter placement","Inhalant anesthesia induction & maintenance","Endotracheal intubation — routine","Endotracheal intubation — difficult airway","Oxygen supplementation & airway management","Mechanical ventilation (IPPV)","Drug protocol development with veterinarian","Premed administration & sedation monitoring","IV induction agent administration","CRI calculation & administration","Epidural injection","Local/regional nerve block","Arterial catheter placement","Direct arterial blood pressure monitoring","Indirect BP monitoring (Doppler/oscillometric)","ECG monitoring & interpretation","Pulse oximetry monitoring","Capnography monitoring & interpretation","Fluid therapy management (rate/type selection)","Colloid/blood product administration","Pre-op pain assessment & multimodal analgesia planning","Intra-op pain assessment & adjustments","Post-op pain assessment & management","Recovery monitoring & support","Emergency drug calculation","CPR/CPCR response","Anesthesia machine setup & leak testing","Troubleshooting anesthesia equipment","Documentation on anesthesia record"];
const SUPP_SKILLS=["Transtracheal wash / BAL under anesthesia","Jugular catheter placement","Central venous pressure monitoring","Arterial blood gas sampling & interpretation","Thoracocentesis/pericardiocentesis assistance","Neuromuscular blockade monitoring","Temperature management (active warming/cooling)","Anesthesia for MRI / special imaging","Anesthesia for field procedures / non-hospital","Exotic species anesthesia","Ophthalmic anesthesia considerations","Thoracotomy anesthesia management","Hepatic disease anesthesia considerations","Renal disease anesthesia considerations","Cardiovascular disease anesthesia management","Neurological patient anesthesia"];

function SkillsTab({coreSkillsPct,setCoreSkillsPct,suppSkillsPct,setSuppSkillsPct}) {
  const [cs,setCs]=useState(Object.fromEntries(CORE_SKILLS.map(s=>[s,{done:false,caseNum:""}])));
  const [ss,setSs]=useState(Object.fromEntries(SUPP_SKILLS.map(s=>[s,{done:false,caseNum:""}])));
  const tc=s=>setCs(p=>{const n={...p,[s]:{...p[s],done:!p[s].done}};const d=Object.values(n).filter(v=>v.done).length;setCoreSkillsPct(Math.round(d/CORE_SKILLS.length*100));return n;});
  const ts=s=>setSs(p=>{const n={...p,[s]:{...p[s],done:!p[s].done}};const d=Object.values(n).filter(v=>v.done).length;setSuppSkillsPct(Math.round(d/SUPP_SKILLS.length*100));return n;});
  const cd=Object.values(cs).filter(v=>v.done).length,sd=Object.values(ss).filter(v=>v.done).length;
  const SR=({skill,status,toggle,setCN})=>(
    <div style={S.checkRow}>
      <div style={S.checkbox(status.done)} onClick={()=>toggle(skill)}>{status.done?"✓":""}</div>
      <div style={{flex:1,fontSize:13,color:status.done?T.muted:T.sub,textDecoration:status.done?"line-through":"none"}} onClick={()=>toggle(skill)}>{skill}</div>
      {status.done&&<input type="text" placeholder="Case #" style={{...S.input,width:65,padding:"4px 8px",fontSize:11}} value={status.caseNum} onChange={e=>setCN(skill,e.target.value)} onClick={e=>e.stopPropagation()}/>}
    </div>
  );
  return (
    <div>
      <div style={S.sectionTitle}>Skills Coverage Map</div>
      <div style={S.sectionSub}>Check off mastered skills and enter the case log number that demonstrates each. Need 90% core, 50% supplemental.</div>
      <div style={S.grid2}>
        <div style={S.statCard(coreSkillsPct>=90?"green":"red")}>
          <div style={{fontSize:11,color:T.muted,fontWeight:700,marginBottom:4}}>CORE SKILLS</div>
          <div style={{fontSize:28,fontWeight:900,color:coreSkillsPct>=90?T.green:T.red}}>{cd}/{CORE_SKILLS.length}</div>
          <div style={S.scoreBar}><div style={S.scoreFill(coreSkillsPct,coreSkillsPct>=90?"green":"red")}/></div>
          <div style={{fontSize:11,color:T.muted,marginTop:4}}>{coreSkillsPct}% — Need 90%+</div>
          {coreSkillsPct<90&&<div style={{fontSize:11,color:T.red,marginTop:4}}>Need {Math.ceil(CORE_SKILLS.length*.9)-cd} more</div>}
        </div>
        <div style={S.statCard(suppSkillsPct>=50?"green":"amber")}>
          <div style={{fontSize:11,color:T.muted,fontWeight:700,marginBottom:4}}>SUPPLEMENTAL SKILLS</div>
          <div style={{fontSize:28,fontWeight:900,color:suppSkillsPct>=50?T.green:T.amber}}>{sd}/{SUPP_SKILLS.length}</div>
          <div style={S.scoreBar}><div style={S.scoreFill(suppSkillsPct,suppSkillsPct>=50?"green":"amber")}/></div>
          <div style={{fontSize:11,color:T.muted,marginTop:4}}>{suppSkillsPct}% — Need 50%+</div>
          {suppSkillsPct<50&&<div style={{fontSize:11,color:T.amber,marginTop:4}}>Need {Math.ceil(SUPP_SKILLS.length*.5)-sd} more</div>}
        </div>
      </div>
      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:14,marginBottom:4}}>Core Skills — Need ≥{Math.ceil(CORE_SKILLS.length*.9)} of {CORE_SKILLS.length}</div>
        <div style={{fontSize:12,color:T.muted,marginBottom:14}}>Each skill must be described IN CONTEXT in a case log — not just named. Enter the case # for each.</div>
        {CORE_SKILLS.map(s=><SR key={s} skill={s} status={cs[s]} toggle={tc} setCN={(sk,v)=>setCs(p=>({...p,[sk]:{...p[sk],caseNum:v}}))}/>)}
      </div>
      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:14,marginBottom:4}}>Supplemental Skills — Need ≥{Math.ceil(SUPP_SKILLS.length*.5)} of {SUPP_SKILLS.length}</div>
        <div style={{fontSize:12,color:T.muted,marginBottom:14}}>Same rules apply — must be shown in context in your case logs.</div>
        {SUPP_SKILLS.map(s=><SR key={s} skill={s} status={ss[s]} toggle={ts} setCN={(sk,v)=>setSs(p=>({...p,[sk]:{...p[sk],caseNum:v}}))}/>)}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// TAB 8 — CE AUDITOR
// ═══════════════════════════════════════════════════════════════════════════════
const CE_SYSTEM=`You are an AVTAA credentials committee reviewer checking CE submissions.
Rules: Min 40 hours. Acceptable: DACVAA, DACVS, DACVECC, VTS (NAVTA), board-eligible DVM (3yr residency), FANZCVS, boarded human specialist at vet conf.
NOT acceptable: DVM only, MRCVS, MANZCVS, DAAPM, CVPP, CCRP, LVMT, LVT, RVT, CVT.
Dates: Jan 1 2020–Dec 31 2024. In-house max 10hrs. Externship max 10hrs + AVTAA pre-approval needed. Journal articles max 3hrs (0.25 each). Content must directly relate to anesthesia or peri-op analgesia.
Output: TOTAL QUALIFYING HOURS: X | HOURS AT RISK: X | FLAGS (numbered) | CLEAN ENTRIES | RECOMMENDATION`;

function CETab({ceHours,setCeHours}) {
  const [ceText,setCeText]=useState("");
  const [result,setResult]=useState("");
  const [loading,setLoading]=useState(false);
  const review=async()=>{
    if(!ceText.trim())return;
    setLoading(true);
    const r=await callClaude(CE_SYSTEM,`Review my AVTAA CE list:\n\n${ceText}`);
    setResult(r);
    const m=r.match(/TOTAL QUALIFYING HOURS:\s*([\d.]+)/);
    if(m)setCeHours(parseFloat(m[1]));
    setLoading(false);
  };
  return (
    <div>
      <div style={S.sectionTitle}>CE Hours Auditor</div>
      <div style={S.sectionSub}>Paste your CE list. Claude flags presenter credential problems, content issues, and category hour limits before you submit.</div>
      <div style={{...S.card,background:T.tealSoft,border:`1px solid ${T.teal}44`,marginBottom:20}}>
        <div style={{fontSize:13,color:T.teal,fontWeight:700,marginBottom:8}}>📋 Quick Rules</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"5px 20px",fontSize:12,color:T.sub}}>
          <div>✅ DACVAA, DACVS, DACVECC</div><div>❌ DVM only — rejected</div>
          <div>✅ VTS (any NAVTA academy)</div><div>❌ CVT/RVT/LVT — rejected</div>
          <div>✅ Board-eligible DVM (3yr residency)</div><div>❌ DAAPM, CVPP, CCRP — rejected</div>
          <div>⚠️ In-house: max 10 hrs</div><div>⚠️ Externship: max 10 hrs + pre-approval</div>
          <div>⚠️ Journal articles: max 3 hrs</div><div>📅 Jan 2020 – Dec 2024 only</div>
        </div>
      </div>
      <div style={S.card}>
        <label style={S.label}>Paste CE List</label>
        <div style={{fontSize:12,color:T.muted,marginBottom:10}}>Format: Conference | Date | Presenter | Credential | Topic | Hours</div>
        <textarea style={{...S.textarea,minHeight:220}} placeholder={"Western Vet Conference | March 2023 | Dr. Jane Smith | DACVAA | Pain Management in the Critical Patient | 1.5 hrs\nVETgirl Online | June 2022 | Dr. Mark Jones | DACVECC | Ventilator Management | 1.0 hr\nIn-house CE | Jan 2024 | Dr. R. Williams | DACVS | Regional Blocks for Orthopedics | 2.0 hrs"} value={ceText} onChange={e=>setCeText(e.target.value)}/>
        <div style={{marginTop:14}}>
          <button style={S.btn("primary")} onClick={review} disabled={loading||!ceText.trim()}>
            {loading?<><span style={S.spinner}/>Auditing CE...</>:"🎓 Audit My CE Hours"}
          </button>
        </div>
        {result&&<AIResult result={result}/>}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// FEATURE REQUEST MODAL
// ═══════════════════════════════════════════════════════════════════════════════
const FEATURE_IDEAS = [
  "Add a timer / countdown to the December 31 deadline",
  "Add an anesthesia record upload and review section",
  "Add a glossary of AVTAA terminology and abbreviations",
  "Add a drug dose calculator (mg/kg converter)",
  "Add a way to export all my progress as a summary PDF",
  "Add a side-by-side comparison of two case report scores",
  "Add a notes/journal section for tracking my progress",
  "Something else (I'll describe it below)",
];

function FeatureRequestModal({ onClose }) {
  const [selected, setSelected] = useState("");
  const [custom, setCustom] = useState("");
  const [copied, setCopied] = useState(false);

  const featureText = selected === FEATURE_IDEAS[FEATURE_IDEAS.length - 1]
    ? custom.trim()
    : selected;

  const prompt = `Hi! I'm using the AVTAA VTS Anesthesia Application Suite dashboard — a React app built to help veterinary technicians prepare their AVTAA VTS (Anesthesia & Analgesia) specialization application.

I'd like to request a new feature be added to the app:

"${featureText}"

The app is a multi-tab React dashboard with the following tabs:
- Readiness Score (pulls data from all tabs into one % score + to-do list)
- Checklist (every required AVTAA document tracked with checkboxes)
- Rejection Analysis (paste AVTAA rejection letter, get a fix plan)
- Case Log Auditor (upload case log PDF, AI reviews every entry)
- Case Report Reviewer (upload .docx reports, format check + 7-section scoring)
- Case Report Builder (guided section-by-section builder with notation prompts)
- Skills Map (core/supplemental skill checklist with case log # tracking)
- CE Auditor (paste CE list, flags presenter credential issues)

It uses the Anthropic Claude API for AI features, mammoth for .docx parsing, JSZip for XML format checking, and Supabase for cloud auto-save.

Please add the requested feature to the existing code and return the full updated .jsx file.`;

  const handleCopy = () => {
    navigator.clipboard?.writeText(prompt).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    });
  };

  return (
    <div style={{position:"fixed",inset:0,zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.7)",backdropFilter:"blur(4px)"}}
      onClick={e=>{ if(e.target===e.currentTarget) onClose(); }}>
      <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:20,padding:32,maxWidth:600,width:"90%",maxHeight:"85vh",overflowY:"auto",boxShadow:`0 0 60px ${T.violet}33`}}>

        {/* Modal header */}
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:20}}>
          <div>
            <div style={{fontSize:18,fontWeight:800,color:T.text,letterSpacing:-0.5}}>✨ Request a Feature</div>
            <div style={{fontSize:12,color:T.muted,marginTop:4}}>Pick what you want added — we'll build the prompt for you to paste into a new Claude chat.</div>
          </div>
          <button onClick={onClose} style={{background:"transparent",border:"none",color:T.muted,fontSize:20,cursor:"pointer",padding:"0 4px",lineHeight:1}}>✕</button>
        </div>

        {/* Feature picker */}
        <label style={S.label}>What would you like added?</label>
        <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:20}}>
          {FEATURE_IDEAS.map(idea => (
            <div key={idea} onClick={()=>setSelected(idea)}
              style={{padding:"10px 14px",borderRadius:10,border:`1px solid ${selected===idea?T.violet:T.border}`,background:selected===idea?T.violetSoft:"transparent",cursor:"pointer",fontSize:13,color:selected===idea?T.violetMid:T.sub,transition:"all 0.15s",display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:16,height:16,borderRadius:"50%",border:`2px solid ${selected===idea?T.violet:T.border}`,background:selected===idea?T.violet:"transparent",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
                {selected===idea && <div style={{width:6,height:6,borderRadius:"50%",background:"#fff"}}/>}
              </div>
              {idea}
            </div>
          ))}
        </div>

        {/* Custom description if "something else" */}
        {selected === FEATURE_IDEAS[FEATURE_IDEAS.length - 1] && (
          <div style={{marginBottom:20}}>
            <label style={S.label}>Describe the feature you want</label>
            <textarea
              style={{...S.textarea, minHeight:80}}
              placeholder="e.g. Add a section where I can upload my anesthesia records and have Claude check them for completeness..."
              value={custom}
              onChange={e=>setCustom(e.target.value)}
              autoFocus
            />
          </div>
        )}

        {/* Generated prompt preview */}
        {featureText && (
          <div style={{marginBottom:20}}>
            <label style={S.label}>Your ready-to-paste Claude prompt</label>
            <div style={{background:"#0e0c18",border:`1px solid ${T.border}`,borderRadius:10,padding:14,fontSize:12,color:T.sub,lineHeight:1.7,maxHeight:200,overflowY:"auto",whiteSpace:"pre-wrap",fontFamily:"monospace"}}>
              {prompt}
            </div>
          </div>
        )}

        {/* How to use */}
        <div style={{background:T.violetSoft,border:`1px solid ${T.violet}44`,borderRadius:10,padding:14,marginBottom:20,fontSize:12,color:T.sub,lineHeight:1.7}}>
          <div style={{fontWeight:700,color:T.violetMid,marginBottom:6}}>How to use this:</div>
          1. Copy the prompt below<br/>
          2. Open a <strong style={{color:T.text}}>new Claude chat</strong> at claude.ai<br/>
          3. Paste the prompt and send it<br/>
          4. Claude will return an updated .jsx file<br/>
          5. Download it and drag it back into Claude to use
        </div>

        {/* Action buttons */}
        <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
          <button
            style={{...S.btn("primary"), opacity: featureText?1:0.5, flex:1}}
            onClick={handleCopy}
            disabled={!featureText}
          >
            {copied ? "✓ Copied to Clipboard!" : "📋 Copy Prompt"}
          </button>
          <button style={{...S.btn("ghost")}} onClick={onClose}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// LOGIN SCREEN
// ═══════════════════════════════════════════════════════════════════════════════
// Password is hashed so it's not visible in plain text in the source.
// ─── AUTH ─────────────────────────────────────────────────────────────────────
// Simple session auth — password gates entry, persists for the tab session.
// localStorage is restricted in the sandbox, auth uses module-level state,
// so we use a module-level variable that survives React re-renders.
let _sessionAuthed = false;

function checkAuth()  { return _sessionAuthed; }
function grantAuth()  { _sessionAuthed = true; }
function revokeAuth() { _sessionAuthed = false; }

function LoginScreen({ onAuth }) {
  const [pw, setPw]           = useState("");
  const [show, setShow]       = useState(false);
  const [error, setError]     = useState("");
  const [checking, setChecking] = useState(false);
  const [shake, setShake]     = useState(false);
  const inputRef              = useRef();

  useEffect(() => { inputRef.current?.focus(); }, []);

  const attempt = () => {
    if (!pw.trim()) return;
    setChecking(true);
    setError("");
    try {
      const PLAIN = "rocky";
      if (pw.trim().toLowerCase() === PLAIN) {
        grantAuth();
        onAuth();
      } else {
        setError("Incorrect password. Try again.");
        setShake(true);
        setPw("");
        setTimeout(() => setShake(false), 600);
        inputRef.current?.focus();
        setChecking(false);
      }
    } catch(_) {
      setError("Something went wrong. Try again.");
      setChecking(false);
    }
  };

  const onKey = (e) => { if (e.key === "Enter") attempt(); };

  return (
    <div style={{
      position:"fixed", inset:0, zIndex:10000,
      background:`radial-gradient(ellipse at 45% 38%, #1c0f45 0%, #0d0a1c 55%, #080612 100%)`,
      display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center",
      fontFamily:"'DM Sans','Helvetica Neue',sans-serif",
    }}>
      <style>{`
        @keyframes loginShake {
          0%,100%{transform:translateX(0)}
          20%{transform:translateX(-10px)}
          40%{transform:translateX(10px)}
          60%{transform:translateX(-6px)}
          80%{transform:translateX(6px)}
        }
        @keyframes loginFadeIn { from{opacity:0;transform:translateY(20px)} to{opacity:1;transform:translateY(0)} }
      `}</style>

      {/* Ambient glow */}
      <div style={{position:"absolute",width:500,height:500,borderRadius:"50%",background:`radial-gradient(circle,${T.violet}12 0%,transparent 70%)`,pointerEvents:"none"}}/>

      {/* Compass logo */}
      <div style={{marginBottom:24,filter:`drop-shadow(0 0 20px ${T.violet}88)`,animation:"loginFadeIn 0.5s ease both"}}>
        <svg viewBox="0 0 80 80" xmlns="http://www.w3.org/2000/svg" width="80" height="80">
          <defs>
            <linearGradient id="lg1" x1="0%" y1="0%" x2="100%" y2="100%">
              <stop offset="0%" stopColor="#7c3aed"/><stop offset="100%" stopColor="#a855f7"/>
            </linearGradient>
          </defs>
          <circle cx="40" cy="40" r="37" fill="none" stroke="url(#lg1)" strokeWidth="1.5" opacity="0.5"/>
          {[0,90,180,270].map((deg,i)=>{
            const r1=30,r2=35,rad=deg*Math.PI/180;
            return <line key={i} x1={40+r1*Math.sin(rad)} y1={40-r1*Math.cos(rad)} x2={40+r2*Math.sin(rad)} y2={40-r2*Math.cos(rad)} stroke="url(#lg1)" strokeWidth="1.5" opacity="0.6"/>;
          })}
          <polygon points="40,7 36,40 40,35 44,40" fill="url(#lg1)"/>
          <polygon points="40,73 44,40 40,45 36,40" fill="#2d1f5e"/>
          <polygon points="73,40 40,36 45,40 40,44" fill="#a855f7" opacity="0.7"/>
          <polygon points="7,40 40,44 35,40 40,36" fill="#2d1f5e" opacity="0.9"/>
          <circle cx="40" cy="40" r="4" fill="url(#lg1)"/>
          <circle cx="40" cy="40" r="1.5" fill="#fff"/>
        </svg>
      </div>

      {/* Title */}
      <div style={{animation:"loginFadeIn 0.5s ease 0.1s both", textAlign:"center", marginBottom:36}}>
        <div style={{fontSize:32,fontWeight:900,color:T.text,letterSpacing:-1}}>VTS Compass</div>
        <div style={{fontSize:12,color:T.muted,marginTop:4,letterSpacing:1}}>POWERED BY CLAUDE</div>
      </div>

      {/* Login card */}
      <div style={{
        animation:`loginFadeIn 0.5s ease 0.2s both, ${shake?"loginShake 0.5s ease":"none"}`,
        background:`linear-gradient(135deg,#1e1040,#12102a)`,
        border:`1px solid ${T.violet}55`,
        borderRadius:20, padding:"32px 40px",
        width:"100%", maxWidth:380,
        boxShadow:`0 0 50px ${T.violet}22, inset 0 1px 0 ${T.violet}33`,
      }}>
        <div style={{fontSize:13,color:T.muted,textAlign:"center",marginBottom:24,letterSpacing:0.5}}>
          Enter your password to continue
        </div>

        <div style={{position:"relative",marginBottom:16}}>
          <input
            ref={inputRef}
            type={show?"text":"password"}
            value={pw}
            onChange={e=>{setPw(e.target.value);setError("");}}
            onKeyDown={onKey}
            placeholder="Password"
            style={{
              width:"100%", padding:"14px 48px 14px 16px",
              background:"#0e0c1a", border:`1px solid ${error?T.red:T.border}`,
              borderRadius:10, color:T.text, fontSize:15, outline:"none",
              fontFamily:"inherit", boxSizing:"border-box",
              transition:"border-color 0.15s",
            }}
          />
          <button
            onClick={()=>setShow(s=>!s)}
            style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",padding:4,opacity:0.7,display:"flex",alignItems:"center",justifyContent:"center"}}
            title={show?"Hide password":"Show password"}
          >
            {show ? (
              // Mini compass (password visible state)
              <svg viewBox="0 0 20 20" width="18" height="18" xmlns="http://www.w3.org/2000/svg">
                <circle cx="10" cy="10" r="8.5" fill="none" stroke={T.violetMid} strokeWidth="1.2" opacity="0.8"/>
                <polygon points="10,2 8.5,10 10,8.5 11.5,10" fill={T.violetMid}/>
                <polygon points="10,18 11.5,10 10,11.5 8.5,10" fill={T.violetSoft}/>
                <polygon points="18,10 10,8.5 11.5,10 10,11.5" fill={T.violetMid} opacity="0.6"/>
                <polygon points="2,10 10,11.5 8.5,10 10,8.5" fill={T.violetSoft}/>
                <circle cx="10" cy="10" r="1.5" fill={T.violetMid}/>
              </svg>
            ) : (
              // Eye icon (password hidden state)
              <svg viewBox="0 0 20 20" width="18" height="18" xmlns="http://www.w3.org/2000/svg" fill="none">
                <path d="M1 10s3-6 9-6 9 6 9 6-3 6-9 6-9-6-9-6z" stroke={T.muted} strokeWidth="1.3"/>
                <circle cx="10" cy="10" r="2.5" stroke={T.muted} strokeWidth="1.3"/>
              </svg>
            )}
          </button>
        </div>

        {error && (
          <div style={{fontSize:12,color:T.red,marginBottom:14,textAlign:"center",fontWeight:600}}>
            ❌ {error}
          </div>
        )}

        <button
          style={{
            width:"100%", padding:"13px",
            background:`linear-gradient(135deg,${T.violet},${T.violetMid})`,
            border:"none", borderRadius:10, color:"#fff",
            fontSize:15, fontWeight:800, cursor:"pointer",
            letterSpacing:0.3, transition:"opacity 0.15s",
            opacity: checking ? 0.7 : 1,
          }}
          onClick={attempt}
          disabled={checking||!pw.trim()}
        >
          {checking ? "Checking..." : "Enter →"}
        </button>

        <div style={{fontSize:11,color:T.muted,textAlign:"center",marginTop:18,lineHeight:1.6}}>
          VTS Compass is a private tool.<br/>Contact the owner for access.
        </div>
      </div>
    </div>
  );
}
function SplashScreen({ onDone }) {
  const [fadeOut, setFadeOut] = useState(false);
  useEffect(() => {
    const t1 = setTimeout(() => setFadeOut(true), 4500);
    const t2 = setTimeout(() => onDone(), 5000);
    return () => { clearTimeout(t1); clearTimeout(t2); };
  }, []);

  const CompassLarge = () => (
    <svg viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg" width="130" height="130">
      <defs>
        <linearGradient id="sp1" x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stopColor="#7c3aed"/><stop offset="100%" stopColor="#a855f7"/>
        </linearGradient>
        <linearGradient id="sp2" x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stopColor="#a855f7"/><stop offset="100%" stopColor="#c084fc"/>
        </linearGradient>
        <filter id="spglow">
          <feGaussianBlur stdDeviation="3" result="blur"/>
          <feMerge><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge>
        </filter>
      </defs>
      <circle cx="60" cy="60" r="57" fill="none" stroke="url(#sp1)" strokeWidth="1.5" opacity="0.5"/>
      <circle cx="60" cy="60" r="49" fill="none" stroke="url(#sp1)" strokeWidth="0.6" opacity="0.25"/>
      {[0,45,90,135,180,225,270,315].map((deg,i)=>{
        const r1=i%2===0?42:46, r2=52, rad=deg*Math.PI/180;
        return <line key={deg}
          x1={60+r1*Math.sin(rad)} y1={60-r1*Math.cos(rad)}
          x2={60+r2*Math.sin(rad)} y2={60-r2*Math.cos(rad)}
          stroke="url(#sp1)" strokeWidth={i%2===0?"2":"1"} opacity="0.65"/>;
      })}
      {[["N",60,13],["S",60,111],["E",109,65],["W",11,65]].map(([l,x,y])=>(
        <text key={l} x={x} y={y} textAnchor="middle" fontSize="11" fontWeight="900"
          fill="#a855f7" fontFamily="DM Sans,sans-serif" opacity="0.9">{l}</text>
      ))}
      <polygon points="60,9 54,60 60,51 66,60"  fill="url(#sp1)" filter="url(#spglow)"/>
      <polygon points="60,111 66,60 60,69 54,60" fill="#2d1f5e"/>
      <polygon points="111,60 60,54 69,60 60,66" fill="url(#sp2)" opacity="0.75"/>
      <polygon points="9,60 60,66 51,60 60,54"   fill="#2d1f5e" opacity="0.9"/>
      <circle cx="60" cy="60" r="7" fill="url(#sp1)" filter="url(#spglow)"/>
      <circle cx="60" cy="60" r="2.8" fill="#fff"/>
    </svg>
  );

  return (
    <div style={{
      position:"fixed", inset:0, zIndex:9999,
      background:`radial-gradient(ellipse at 45% 38%, #1c0f45 0%, #0d0a1c 55%, #080612 100%)`,
      display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center",
      fontFamily:"'DM Sans','Helvetica Neue',sans-serif",
      transition:"opacity 0.5s ease",
      opacity: fadeOut ? 0 : 1,
    }}>
      <style>{`
        @keyframes splashSpin { from{transform:rotate(0deg)} to{transform:rotate(360deg)} }
        @keyframes splashPulse { 0%,100%{opacity:0.6} 50%{opacity:1} }
        @keyframes splashBar { from{width:0%} to{width:100%} }
        @keyframes splashFadeUp { from{opacity:0;transform:translateY(16px)} to{opacity:1;transform:translateY(0)} }
      `}</style>

      {/* Ambient glow */}
      <div style={{position:"absolute",width:600,height:600,borderRadius:"50%",background:`radial-gradient(circle, ${T.violet}14 0%, transparent 68%)`,pointerEvents:"none"}}/>

      {/* Spinning compass */}
      <div style={{animation:"splashSpin 12s linear infinite", filter:`drop-shadow(0 0 28px ${T.violet}99)`, marginBottom:30}}>
        <CompassLarge/>
      </div>

      {/* Title */}
      <div style={{animation:"splashFadeUp 0.6s ease 0.3s both"}}>
        <div style={{fontSize:46,fontWeight:900,color:T.text,letterSpacing:-2,textAlign:"center",lineHeight:1}}>
          VTS Compass
        </div>
        <div style={{fontSize:12,color:T.muted,textAlign:"center",letterSpacing:2,marginTop:8,textTransform:"uppercase"}}>
          Powered by Claude
        </div>
      </div>

      {/* Welcome card */}
      <div style={{
        animation:"splashFadeUp 0.6s ease 0.6s both",
        marginTop:40,
        background:`linear-gradient(135deg, #1e1040 0%, #12102a 100%)`,
        border:`1px solid ${T.violet}55`,
        borderRadius:20, padding:"22px 48px", textAlign:"center",
        boxShadow:`0 0 50px ${T.violet}22, inset 0 1px 0 ${T.violet}33`,
      }}>
        <div style={{fontSize:11,color:T.muted,letterSpacing:2,marginBottom:8,textTransform:"uppercase"}}>Welcome back</div>
        <div style={{fontSize:32,fontWeight:900,color:T.text,letterSpacing:-1}}>
          Sunny Dee! 👋
        </div>
        <div style={{fontSize:13,color:T.sub,marginTop:10,animation:"splashPulse 1.5s ease infinite"}}>
          Loading your dashboard...
        </div>
      </div>

      {/* Progress bar */}
      <div style={{marginTop:32,width:220,height:3,background:T.border,borderRadius:2,overflow:"hidden",animation:"splashFadeUp 0.4s ease 0.8s both"}}>
        <div style={{height:"100%",background:`linear-gradient(90deg,${T.violet},${T.violetMid},#c084fc)`,borderRadius:2,animation:"splashBar 4.5s cubic-bezier(0.4,0,0.2,1) forwards"}}/>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════════════════════════
const TABS=[
  {id:"readiness",label:"🎯 Readiness"},
  {id:"checklist",label:"📋 Checklist"},
  {id:"reports",label:"📝 Case Reports"},
  {id:"builder",label:"✨ Report Builder"},
  {id:"caselogs",label:"📊 Case Logs"},
  {id:"skills",label:"🛠 Skills Map"},
  {id:"rejection",label:"🔍 Rejection Analysis"},
  {id:"ce",label:"🎓 CE Auditor"},
];

export default function App() {
  const [authed, setAuthed] = useState(() => checkAuth());
  const [showSplash, setShowSplash] = useState(true);
  const [tab,setTab]=useState("readiness");
  const [savedIndicator,setSavedIndicator]=useState(true);
  const [showFeatureModal,setShowFeatureModal]=useState(false);

  // Preload JSZip on mount so it's ready when first .docx is dropped
  useEffect(() => { loadJSZip().catch(()=>{}); }, []);

  // shared state
  const [checked,setChecked]=useState(()=>{
    const allIds=Object.values(CHECKLIST_ITEMS).flat().map(i=>i.id);
    return Object.fromEntries(allIds.map(id=>[id,id.startsWith("pa")]));
  });
  const [caseLogCount,setCaseLogCount]=useState(0);
  const [asaCounts,setAsaCounts]=useState({I:0,II:0,III:0,IV:0,V:0,E:0});
  const [sedationOnly,setSedationOnly]=useState(0);
  const [reportScores,setReportScores]=useState([null,null,null,null]);
  const [reportUploadCount,setReportUploadCount]=useState(0);
  const [ceHours,setCeHours]=useState(0);
  const [coreSkillsPct,setCoreSkillsPct]=useState(0);
  const [suppSkillsPct,setSuppSkillsPct]=useState(0);
  const [caseLogRecords,setCaseLogRecords]=useState([]);
  const [builderSections,setBuilderSections]=useState(
    Object.fromEntries(BUILDER_SECTIONS.map(s=>[s.id,{notes:"",polished:"",loading:false}]))
  );

  // auto-save
  const saveData=useCallback(async()=>{
    setSavedIndicator(false);
    // strip loading flags before saving — they're transient
    const builderToSave = Object.fromEntries(
      Object.entries(builderSections).map(([k,v])=>[k,{notes:v.notes,polished:v.polished}])
    );
    await storageSave({checked,caseLogCount,asaCounts,sedationOnly,reportScores,ceHours,coreSkillsPct,suppSkillsPct,builderToSave,caseLogRecords,savedAt:Date.now()});
    setSavedIndicator(true);
  },[checked,caseLogCount,asaCounts,sedationOnly,reportScores,ceHours,coreSkillsPct,suppSkillsPct,builderSections,caseLogRecords]);

  // load on mount — pulls from Supabase cloud
  useEffect(()=>{
    setSavedIndicator(false);
    storageLoad().then(d=>{
      if(!d){ setSavedIndicator(true); return; }
      if(d.checked)setChecked(d.checked);
      if(d.caseLogCount)setCaseLogCount(d.caseLogCount);
      if(d.asaCounts)setAsaCounts(d.asaCounts);
      if(d.sedationOnly!=null)setSedationOnly(d.sedationOnly);
      if(d.reportScores)setReportScores(d.reportScores);
      if(d.ceHours)setCeHours(d.ceHours);
      if(d.coreSkillsPct!=null)setCoreSkillsPct(d.coreSkillsPct);
      if(d.suppSkillsPct!=null)setSuppSkillsPct(d.suppSkillsPct);
      if(d.caseLogRecords) setCaseLogRecords(d.caseLogRecords);
      if(d.builderToSave){
        setBuilderSections(prev=>{
          const merged={...prev};
          Object.entries(d.builderToSave).forEach(([k,v])=>{
            if(merged[k])merged[k]={...merged[k],notes:v.notes||"",polished:v.polished||""};
          });
          return merged;
        });
      }
      setSavedIndicator(true);
    });
  },[]);

  // save on change
  useEffect(()=>{const t=setTimeout(saveData,1200);return()=>clearTimeout(t);},[saveData]);

  // ── Export progress as JSON file ──────────────────────────────────────────
  const exportProgress = () => {
    const builderToSave = Object.fromEntries(
      Object.entries(builderSections).map(([k,v])=>[k,{notes:v.notes,polished:v.polished}])
    );
    const data = {
      checked, caseLogCount, asaCounts, sedationOnly,
      reportScores, ceHours, coreSkillsPct, suppSkillsPct,
      builderToSave, caseLogRecords, savedAt: Date.now(), version: "vtsc_v1",
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], {type:"application/json"});
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href     = url;
    a.download = `vts-compass-progress-${new Date().toISOString().slice(0,10)}.json`;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
  };

  // ── Import progress from JSON file ────────────────────────────────────────
  const importProgress = () => {
    const input = document.createElement("input");
    input.type  = "file";
    input.accept = ".json";
    input.onchange = (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        try {
          const d = JSON.parse(ev.target.result);
          if (d.version !== "vtsc_v1") { alert("This file doesn't look like a VTS Compass progress file."); return; }
          if (d.checked)              setChecked(d.checked);
          if (d.caseLogCount)         setCaseLogCount(d.caseLogCount);
          if (d.asaCounts)            setAsaCounts(d.asaCounts);
          if (d.sedationOnly != null) setSedationOnly(d.sedationOnly);
          if (d.reportScores)         setReportScores(d.reportScores);
          if (d.ceHours)              setCeHours(d.ceHours);
          if (d.coreSkillsPct != null) setCoreSkillsPct(d.coreSkillsPct);
          if (d.suppSkillsPct != null) setSuppSkillsPct(d.suppSkillsPct);
          if (d.caseLogRecords) setCaseLogRecords(d.caseLogRecords);
          if (d.builderToSave) {
            setBuilderSections(prev => {
              const merged = {...prev};
              Object.entries(d.builderToSave).forEach(([k,v]) => {
                if (merged[k]) merged[k] = {...merged[k], notes:v.notes||"", polished:v.polished||""};
              });
              return merged;
            });
          }
          alert("✓ Progress loaded successfully!");
        } catch(_) { alert("Could not read this file. Make sure it's a VTS Compass progress file."); }
      };
      reader.readAsText(file);
    };
    document.body.appendChild(input); input.click(); document.body.removeChild(input);
  };

  return (
    <div style={{minHeight:"100vh",background:T.bg,fontFamily:"'DM Sans','Helvetica Neue',sans-serif",color:T.text}}>
      {!authed && <LoginScreen onAuth={() => setAuthed(true)}/>}
      {authed && showSplash && <SplashScreen onDone={() => setShowSplash(false)}/>}
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700;800;900&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}
        @keyframes spin{to{transform:rotate(360deg)}}
        button:hover:not(:disabled){opacity:0.85;transform:translateY(-1px)}
        button:disabled{opacity:0.45;cursor:not-allowed}
        textarea:focus,input:focus{border-color:${T.violet}!important;box-shadow:0 0 0 2px ${T.violet}22}
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-track{background:${T.panel}}
        ::-webkit-scrollbar-thumb{background:${T.border};border-radius:3px}
      `}</style>

      {/* Header */}
      <div style={{background:`linear-gradient(135deg,#0e0b1a 0%,#1a1030 50%,#0e0b1a 100%)`,borderBottom:`1px solid ${T.border}`,padding:"20px 40px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div style={{display:"flex",alignItems:"center",gap:14}}>
          <div style={{width:50,height:50,flexShrink:0}}>
            <svg viewBox="0 0 50 50" xmlns="http://www.w3.org/2000/svg" width="50" height="50">
              <defs>
                <linearGradient id="cg1" x1="0%" y1="0%" x2="100%" y2="100%">
                  <stop offset="0%" stopColor="#7c3aed"/>
                  <stop offset="100%" stopColor="#a855f7"/>
                </linearGradient>
                <linearGradient id="cg2" x1="0%" y1="0%" x2="100%" y2="100%">
                  <stop offset="0%" stopColor="#a855f7"/>
                  <stop offset="100%" stopColor="#c084fc"/>
                </linearGradient>
                <filter id="cglow">
                  <feGaussianBlur stdDeviation="1.5" result="blur"/>
                  <feMerge><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge>
                </filter>
              </defs>
              {/* Outer ring */}
              <circle cx="25" cy="25" r="23" fill="none" stroke="url(#cg1)" strokeWidth="1.5" opacity="0.6"/>
              <circle cx="25" cy="25" r="19" fill="none" stroke="url(#cg1)" strokeWidth="0.5" opacity="0.3"/>
              {/* Cardinal tick marks */}
              {[0,45,90,135,180,225,270,315].map((deg,i)=>{
                const r1 = i%2===0 ? 17 : 18.5;
                const r2 = 21;
                const rad = deg * Math.PI / 180;
                return <line key={deg}
                  x1={25+r1*Math.sin(rad)} y1={25-r1*Math.cos(rad)}
                  x2={25+r2*Math.sin(rad)} y2={25-r2*Math.cos(rad)}
                  stroke="url(#cg1)" strokeWidth={i%2===0?"1.5":"0.8"} opacity="0.7"/>;
              })}
              {/* North arrow — filled violet */}
              <polygon points="25,4 22,25 25,22 28,25" fill="url(#cg1)" filter="url(#cglow)"/>
              {/* South arrow — muted */}
              <polygon points="25,46 28,25 25,28 22,25" fill="#2d1f5e"/>
              {/* East arrow — filled violet */}
              <polygon points="46,25 25,22 28,25 25,28" fill="url(#cg2)" opacity="0.75"/>
              {/* West arrow — muted */}
              <polygon points="4,25 25,28 22,25 25,22" fill="#2d1f5e" opacity="0.9"/>
              {/* Center dot */}
              <circle cx="25" cy="25" r="2.5" fill="url(#cg1)" filter="url(#cglow)"/>
              <circle cx="25" cy="25" r="1" fill="#fff"/>
              {/* N label */}
              <text x="25" y="17.5" textAnchor="middle" fontSize="5" fontWeight="800" fill="#a855f7" fontFamily="DM Sans,sans-serif" opacity="0.9">N</text>
            </svg>
          </div>
          <div>
            <div style={{fontSize:26,fontWeight:900,color:T.text,letterSpacing:-1}}>VTS Compass</div>
            <div style={{fontSize:11,color:T.muted,marginTop:3}}>Powered by Claude</div>
          </div>
        </div>
        <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
          <button onClick={importProgress} style={{...S.btn("ghost"),fontSize:12,padding:"7px 14px",border:`1px solid ${T.teal}55`,color:T.teal}} title="Load previously saved progress from a .json file">
            ⬆️ Load Progress
          </button>
          <button onClick={exportProgress} style={{...S.btn("ghost"),fontSize:12,padding:"7px 14px",border:`1px solid ${T.green}55`,color:T.green}} title="Download all your progress as a .json file you can reload later">
            ⬇️ Save Progress
          </button>
          <button onClick={()=>setShowFeatureModal(true)} style={{...S.btn("ghost"),fontSize:12,padding:"7px 14px",border:`1px solid ${T.violet}55`,color:T.violetMid}}>✨ Request a Feature</button>
          <div style={{background:T.amberSoft,border:`1px solid ${T.amber}55`,borderRadius:8,padding:"7px 14px",fontSize:12,color:T.amber,fontWeight:700}}>⏰ DEC 31 DEADLINE</div>
        </div>
      </div>

      {/* Nav */}
      <div style={{display:"flex",padding:"0 40px",borderBottom:`1px solid ${T.border}`,background:T.panel,overflowX:"auto"}}>
        {TABS.map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"15px 20px",fontSize:14,fontWeight:tab===t.id?700:500,color:tab===t.id?T.violetMid:T.muted,background:"transparent",border:"none",borderBottom:tab===t.id?`2px solid ${T.violet}`:"2px solid transparent",cursor:"pointer",whiteSpace:"nowrap",transition:"all 0.15s"}}>
            {t.label}
          </button>
        ))}
      </div>

      {/* Content */}
      <div style={{padding:"30px 40px",maxWidth:1200,margin:"0 auto"}}>
        {tab==="readiness"&&<ReadinessTab checked={checked} caseLogCount={caseLogCount} caseLogRecords={caseLogRecords} asaCounts={asaCounts} sedationOnly={sedationOnly} reportScores={reportScores} reportUploadCount={reportUploadCount} ceHours={ceHours} coreSkillsPct={coreSkillsPct} suppSkillsPct={suppSkillsPct}/>}
        {tab==="checklist"&&<ChecklistTab checked={checked} setChecked={setCheckedWrapped}/>}
        {tab==="rejection"&&<RejectionTab/>}
        {tab==="caselogs"&&<CaseLogTab caseLogCount={caseLogCount} setCaseLogCount={setCaseLogCount} asaCounts={asaCounts} setAsaCounts={setAsaCounts} sedationOnly={sedationOnly} setSedationOnly={setSedationOnly} caseLogRecords={caseLogRecords} setCaseLogRecords={setCaseLogRecords}/>}
        {tab==="reports"&&<CaseReportTab reportScores={reportScores} setReportScores={setReportScores} setReportUploadCount={setReportUploadCount}/>}
        {tab==="builder"&&<CaseReportBuilderTab builderSections={builderSections} setBuilderSections={setBuilderSections}/>}
        {tab==="skills"&&<SkillsTab coreSkillsPct={coreSkillsPct} setCoreSkillsPct={setCoreSkillsPct} suppSkillsPct={suppSkillsPct} setSuppSkillsPct={setSuppSkillsPct}/>}
        {tab==="ce"&&<CETab ceHours={ceHours} setCeHours={setCeHours}/>}
      </div>

      {showFeatureModal && <FeatureRequestModal onClose={()=>setShowFeatureModal(false)}/>}
      <SaveIndicator saved={savedIndicator}/>
    </div>
  );
}
