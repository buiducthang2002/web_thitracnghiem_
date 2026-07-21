import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from "recharts";
import { BookOpen, Users, FileText, BarChart2, LogOut, Plus, Trash2, Clock, CheckCircle, XCircle, Award, Home, Play, TrendingUp, X, ChevronRight, Shield, ShieldCheck, Star, ArrowRight, ArrowLeft, Upload, Download, AlertCircle } from "lucide-react";
import { db } from "./firebase";
import { collection, doc, setDoc, deleteDoc, onSnapshot } from "firebase/firestore";

const COL = { employees:'employees', questions:'questions', exams:'exams', results:'results' };

const EMPLOYEES_INIT = [];
const QUESTIONS_INIT = [];
const EXAMS_INIT = [];
const RESULTS_INIT = [];

const AV_COLORS = ["bg-emerald-600","bg-teal-600","bg-violet-500","bg-amber-600","bg-rose-500"];
const getAv = n => n.split(' ').map(w=>w[0]).join('').slice(-2).toUpperCase();
const Avatar = ({name, sz="md"}) => {
  const s = {sm:"w-7 h-7 text-xs",md:"w-9 h-9 text-sm",lg:"w-12 h-12 text-base"}[sz];
  return <div className={`${s} ${AV_COLORS[name.charCodeAt(0)%5]} rounded-full flex items-center justify-center text-white font-bold flex-shrink-0`}>{getAv(name)}</div>;
};
const Badge = ({ok}) => <span className={`text-sm font-bold ${ok?'text-emerald-500':'text-red-400'}`}>{ok?'✓':'✗'}</span>;

// Stable question order: by explicit `order`; legacy questions without one keep id order and stay first
const orderedQuestions = (qs) => [...qs].sort((a,b)=>{
  if(a.order!=null && b.order!=null) return a.order-b.order;
  if(a.order==null && b.order==null) return (a.id||0)-(b.id||0);
  return a.order==null ? -1 : 1;
});


// ── SIDEBAR (desktop) + MOBILE NAV ──
const Sidebar = ({role, active, setActive, user, onLogout}) => {
  const nav = role==='admin'
    ? [{id:'dashboard',ic:<Home size={17}/>,lb:'Tổng quan'},{id:'questions',ic:<BookOpen size={17}/>,lb:'Câu hỏi'},{id:'exams',ic:<FileText size={17}/>,lb:'Đề thi'},{id:'results',ic:<Award size={17}/>,lb:'Kết quả'},{id:'employees',ic:<Users size={17}/>,lb:'Thí sinh'}]
    : [{id:'home',ic:<Home size={17}/>,lb:'Trang chủ'},{id:'results',ic:<Award size={17}/>,lb:'Kết quả'}];
  return (
    <>
      {/* Desktop sidebar */}
      <div className="hidden md:flex w-64 bg-green-950 h-screen flex-col fixed left-0 top-0 z-10">
        <div className="p-4 border-b border-green-800/50 flex flex-col items-center gap-2 text-center">
          <div className="w-8 h-8 bg-emerald-600 rounded-lg flex items-center justify-center flex-shrink-0"><Shield size={15} className="text-white"/></div>
          <div><div className="text-white font-bold text-[11px] leading-tight whitespace-nowrap">Cục hậu cần - kỹ thuật Quân khu 4</div><div className="text-slate-400 text-[10px] whitespace-nowrap">Hệ thống thi trắc nghiệm</div></div>
        </div>
        <nav className="flex-1 p-2.5 space-y-0.5">
          {nav.map(i=>(
            <button key={i.id} onClick={()=>setActive(i.id)} className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm transition-all ${active===i.id?'bg-emerald-600 text-white':'text-slate-400 hover:text-white hover:bg-green-900'}`}>
              {i.ic}<span>{i.lb}</span>
            </button>
          ))}
        </nav>
        <div className="p-2.5 border-t border-green-800/50">
          <div className="flex items-center gap-2 px-2 py-1.5">
            <Avatar name={user.name} sz="sm"/>
            <div className="flex-1 min-w-0">
              <div className="text-white text-xs font-medium truncate">{user.name}</div>
              <div className="text-slate-400 text-xs">{role==='admin'?'Quản trị viên':user.dept}</div>
            </div>
            <button onClick={onLogout} className="text-slate-400 hover:text-red-400"><LogOut size={14}/></button>
          </div>
        </div>
      </div>

      {/* Mobile top bar */}
      <div className="md:hidden fixed top-0 left-0 right-0 z-20 bg-green-950 flex items-center justify-between px-4 py-2.5 border-b border-green-800/50">
        <div className="flex items-center gap-2 flex-shrink-0">
          <div className="w-7 h-7 bg-emerald-600 rounded-lg flex items-center justify-center flex-shrink-0"><Shield size={13} className="text-white"/></div>
          <div>
            <div className="text-white font-bold text-[10px] leading-tight whitespace-nowrap">Cục hậu cần - kỹ thuật Quân khu 4</div>
            <div className="text-slate-400 text-[9px] whitespace-nowrap">Hệ thống thi trắc nghiệm</div>
          </div>
        </div>
        <div className="flex items-center gap-2 min-w-0">
          <span className="text-white text-xs truncate">{user.name.split(' ').pop()}</span>
          <button onClick={onLogout} className="text-slate-400 hover:text-red-300 p-1"><LogOut size={15}/></button>
        </div>
      </div>

      {/* Mobile bottom nav */}
      <div className="md:hidden fixed bottom-0 left-0 right-0 z-20 bg-green-950 border-t border-green-800/50 flex">
        {nav.map(i=>(
          <button key={i.id} onClick={()=>setActive(i.id)} className={`flex-1 flex flex-col items-center gap-0.5 py-2.5 transition-all ${active===i.id?'text-emerald-400':'text-slate-500'}`}>
            {i.ic}
            <span className="text-[10px]">{i.lb}</span>
          </button>
        ))}
      </div>
    </>
  );
};

// ── DASHBOARD ──
const Dashboard = ({results, exams, questions, employees}) => {
  const total = results.length;
  const passed = results.filter(r=>{const e=exams.find(x=>x.id===r.examId);return e&&r.score>=e.pass;}).length;
  const avg = total?Math.round(results.reduce((a,r)=>a+r.score,0)/total):0;
  const barData = exams.map(e=>{const rs=results.filter(r=>r.examId===e.id);return {name:e.title.replace('Kiểm tra ','').replace(' 2024',''),avg:rs.length?Math.round(rs.reduce((a,r)=>a+r.score,0)/rs.length):0};});
  const pie = [{name:'Đạt',value:passed,c:'#22c55e'},{name:'Chưa đạt',value:total-passed,c:'#f87171'}];
  const recent = [...results].sort((a,b)=>b.id-a.id).slice(0,6);
  const stats = [
    {ic:<FileText size={18}/>,val:total,lb:'Lượt thi',col:'bg-emerald-50 text-emerald-700'},
    {ic:<CheckCircle size={18}/>,val:`${total?Math.round(passed/total*100):0}%`,lb:'Tỉ lệ đạt',col:'bg-teal-100 text-teal-700'},
    {ic:<TrendingUp size={18}/>,val:`${avg}%`,lb:'Điểm trung bình',col:'bg-amber-100 text-amber-700'},
    {ic:<BookOpen size={18}/>,val:questions.length,lb:'Câu hỏi',col:'bg-violet-100 text-violet-700'},
  ];
  return (
    <div>
      <div className="mb-4"><h1 className="text-lg md:text-xl font-bold text-slate-800">Tổng quan</h1><p className="text-slate-500 text-xs md:text-sm mt-0.5">Thống kê hệ thống thi trắc nghiệm nội bộ</p></div>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-2 md:gap-3 mb-4">
        {stats.map((s,i)=>(
          <div key={i} className="bg-white rounded-xl p-3 md:p-4 shadow-sm border border-slate-100 flex items-center gap-3">
            <div className={`w-10 h-10 ${s.col} rounded-lg flex items-center justify-center flex-shrink-0`}>{s.ic}</div>
            <div>
              <div className="text-xl md:text-2xl font-bold text-slate-800">{s.val}</div>
              <div className="text-xs text-slate-500">{s.lb}</div>
            </div>
          </div>
        ))}
      </div>
      <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mb-4">
        <div className="md:lg:col-span-2 bg-white rounded-xl p-4 shadow-sm border border-slate-100">
          <h3 className="text-sm font-semibold text-slate-700 mb-3">Điểm TB theo đề thi</h3>
          <ResponsiveContainer width="100%" height={150}>
            <BarChart data={barData} margin={{left:-20,right:10}}>
              <XAxis dataKey="name" tick={{fontSize:10}} interval={0}/>
              <YAxis domain={[0,100]} tick={{fontSize:10}}/>
              <Tooltip cursor={false} formatter={v=>[`${v}%`,'Điểm TB']}/>
              <Bar dataKey="avg" fill="#15803d" radius={[4,4,0,0]} cursor={false}/>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="bg-white rounded-xl p-4 shadow-sm border border-slate-100">
          <h3 className="text-sm font-semibold text-slate-700 mb-2">Tỉ lệ đạt/trượt</h3>
          <div className="flex items-center gap-4">
            <div className="flex-shrink-0">
              <ResponsiveContainer width={110} height={110}>
                <PieChart><Pie data={pie} cx="50%" cy="50%" innerRadius={32} outerRadius={50} dataKey="value">{pie.map((d,i)=><Cell key={i} fill={d.c}/>)}</Pie><Tooltip/></PieChart>
              </ResponsiveContainer>
              <div className="flex justify-center gap-3 mt-1">
                {pie.map(d=><div key={d.name} className="flex items-center gap-1"><div className="w-2 h-2 rounded-full" style={{background:d.c}}/><span className="text-xs text-slate-500">{d.name}: {d.value}</span></div>)}
              </div>
            </div>
            <div className="flex-1 space-y-2">
              {[
                {lb:'Tổng lượt thi', val:total, cl:'text-slate-700'},
                {lb:'Số bài đạt', val:passed, cl:'text-emerald-600'},
                {lb:'Số bài trượt', val:total-passed, cl:'text-red-500'},
                {lb:'Tỉ lệ đạt', val:`${total?Math.round(passed/total*100):0}%`, cl:'text-emerald-600 font-bold'},
              ].map(s=>(
                <div key={s.lb} className="flex items-center justify-between bg-slate-50 rounded-lg px-3 py-1.5">
                  <span className="text-xs text-slate-500">{s.lb}</span>
                  <span className={`text-sm font-semibold ${s.cl}`}>{s.val}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
      <div className="bg-white rounded-xl shadow-sm border border-slate-100 overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-100"><h3 className="text-sm font-semibold text-slate-700">Kết quả gần đây</h3></div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[480px]">
            <thead><tr className="bg-slate-50">{['Thí sinh','Đơn vị','Đề thi','Điểm','Kết quả','Thời gian','Ngày'].map(h=><th key={h} className={`px-3 py-2 text-xs font-medium text-slate-400 uppercase ${h==='Kết quả'?'text-center':'text-left'}`}>{h}</th>)}</tr></thead>
            <tbody>
              {recent.map(r=>{
                const emp=employees.find(e=>e.id===r.empId); const exam=exams.find(e=>e.id===r.examId); const ok=exam&&r.score>=exam.pass;
                const mins=r.timeTaken!=null?Math.floor(r.timeTaken/60):null;
                const secs=r.timeTaken!=null?r.timeTaken%60:null;
                const timeStr=mins!=null?`${mins}p${String(secs).padStart(2,'0')}s`:'--';
                return <tr key={r.id} className="border-t border-slate-50 hover:bg-slate-50/50">
                  <td className="px-3 py-2 text-xs text-slate-700">{emp?.name}</td>
                  <td className="px-3 py-2"><span className="px-2 py-0.5 bg-slate-100 text-slate-600 rounded-full text-xs">{emp?.dept||'--'}</span></td>
                  <td className="px-3 py-2 text-xs text-slate-600 max-w-[140px] truncate">{exam?.title}</td>
                  <td className="px-3 py-2 text-xs font-bold text-slate-800">{r.score}%</td>
                  <td className="px-3 py-2 text-center"><Badge ok={ok}/></td>
                  <td className="px-3 py-2"><span className="flex items-center gap-1 text-xs text-slate-500"><Clock size={10}/>{timeStr}</span></td>
                  <td className="px-3 py-2 text-xs text-slate-400">{r.date}</td>
                </tr>;
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// ── QUESTIONS ──
const Questions = ({questions, setQuestions}) => {
  const [modal, setModal] = useState(false);
  const [nq, setNq] = useState({topic:'Nội quy',level:'Dễ',text:'',opts:['','','',''],ans:0});
  const [customTopic, setCustomTopic] = useState('');
  const [importing, setImporting] = useState(false);
  const [importResult, setImportResult] = useState(null);
  const [previewList, setPreviewList] = useState(null); // questions parsed, waiting confirm
  const wordRef = useRef();

  const topics = ['Nội quy','An toàn lao động','Kỹ năng bán hàng',
    ...[...new Set(questions.map(q=>q.topic))].filter(t=>!['Nội quy','An toàn lao động','Kỹ năng bán hàng'].includes(t))
  ];

  const effectiveTopic = nq.topic==='__custom__' ? customTopic.trim() : nq.topic;

  const add = () => {
    if(!nq.text.trim()||nq.opts.some(o=>!o.trim())) return;
    if(!effectiveTopic) return;
    setQuestions(p=>[...p,{...nq, topic:effectiveTopic, id:Date.now(), order: p.reduce((m,q)=>Math.max(m, q.order||0), 0) + 1}]);
    setModal(false);
    setNq({topic:'Nội quy',level:'Dễ',text:'',opts:['','','',''],ans:0});
    setCustomTopic('');
  };

  // Parse text extracted from Word doc
  const parseWordText = (text) => {
    const parsed = [];
    let skipped = 0;
    let norm = text.replace(/\r\n?/g, '\n');
    // Detach an option marker glued right after sentence punctuation ("...?A. xxx")
    norm = norm.replace(/([?”’"):.])(?=[A-Da-d]\s*[\.\)]\s)/g, '$1\n');
    // Split into blocks at each "Câu N" marker
    const blocks = norm.split(/\n(?=\s*Câu\s*\d+\s*[\.\:\)])/i).filter(b=>b.trim());
    // Word often lays options out in a table / 2-column grid, so mammoth puts two
    // options on one line ("A. xxx  B. yyy"). Split every line at each inline
    // option marker or meta marker so each option/meta lands on its own line.
    const splitRe = /\s+(?=(?:[A-Da-d]\s*[\.\)]\s)|(?:Đáp\s*án(?:\s*đúng)?|ĐA|Chủ\s*đề|Mức\s*độ)\s*[:\.\-])/i;
    for(const block of blocks) {
      const lines = [];
      for(const raw of block.split('\n')) {
        const t = raw.trim();
        if(!t) continue;
        t.split(splitRe).forEach(s=>{ const x=s.trim(); if(x) lines.push(x); });
      }
      if(lines.length < 2) { skipped++; continue; }

      let ansIdx = -1;
      let topic = 'Nội quy';
      let level = 'Dễ';
      // Question text = first line with the "Câu N:" prefix stripped
      const qText = lines[0].replace(/^\s*Câu\s*\d+\s*[\.\:\)]\s*/i,'').trim();
      // Collect candidate option lines (everything after the question that isn't meta)
      const cand = [];
      for(let i=1;i<lines.length;i++){
        const l = lines[i];
        const ansMatch = l.match(/^(?:Đáp\s*án(?:\s*đúng)?|ĐA)\s*[:\.\-]\s*([A-Da-d])/i);
        if(ansMatch){ ansIdx = 'abcd'.indexOf(ansMatch[1].toLowerCase()); continue; }
        const topicMatch = l.match(/^Chủ\s*đề\s*[:\.\-]\s*(.+)/i);
        if(topicMatch){ topic = topicMatch[1].trim(); continue; }
        const lvlMatch = l.match(/^Mức\s*độ\s*[:\.\-]\s*(.+)/i);
        if(lvlMatch){ level = lvlMatch[1].trim(); continue; }
        cand.push(l);
      }
      // Assign candidate lines to option slots A..D by position. An explicit
      // "A./B./C./D." label resyncs the slot, so options that lost their label
      // (or are glued onto another line) still land in the right place.
      const slots = [];
      let next = 0;
      for(const c of cand){
        const m = c.match(/^([A-Da-d])\s*[\.\)]\s*(.*)/);
        if(m){ const idx = 'abcd'.indexOf(m[1].toLowerCase()); slots[idx] = (m[2]||'').trim(); next = idx+1; }
        else { slots[next] = c; next++; }
      }
      // Take options contiguously from slot A so the answer index stays aligned
      const opts = [];
      for(let k=0;k<4;k++){ if(!slots[k] || !slots[k].trim()) break; opts.push(slots[k].trim()); }
      // Keep a question if it has text, 2-4 options, and a valid answer pointing at a real option
      if(qText && opts.length >= 2 && opts.length <= 4 && ansIdx >= 0 && ansIdx < opts.length)
        parsed.push({id:Date.now()+Math.random(), text:qText, opts, ans:ansIdx, topic, level});
      else
        skipped++;
    }
    return {parsed, skipped};
  };

  const handleWordImport = async (e) => {
    const file = e.target.files[0]; if(!file) return;
    setImporting(true); setImportResult(null);
    try {
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.extractRawText({arrayBuffer});
      const {parsed, skipped} = parseWordText(result.value);
      if(parsed.length===0){
        setImportResult({error:'Không tìm thấy câu hỏi hợp lệ nào. Mỗi câu cần có nội dung, từ 2 đến 4 đáp án (A, B, C, D) và dòng "Đáp án: X". Vui lòng kiểm tra lại định dạng file.'});
      } else {
        if(skipped>0) setImportResult({error:`Lưu ý: có ${skipped} câu bị bỏ qua do thiếu đáp án (cần 2-4 đáp án) hoặc thiếu dòng "Đáp án: X".`});
        setPreviewList(parsed);
      }    } catch(err) {
      setImportResult({error:'Không đọc được file Word. Vui lòng dùng định dạng .docx — lỗi: ' + err.message});
    }
    setImporting(false);
    e.target.value='';
  };

  const confirmImport = () => {
    setQuestions(p=>{
      // Append imported questions after existing ones with a continuous order
      const base = p.reduce((m,q)=>Math.max(m, q.order||0), 0);
      const ordered = previewList.map((q,i)=>({...q, order: base + i + 1}));
      return [...p, ...ordered];
    });
    setImportResult({added:previewList.length});
    setPreviewList(null);
    setTimeout(()=>setImportResult(null), 2000);
  };

  const downloadWordTemplate = () => {
    // Generate a Word-openable file (HTML-based .doc) with a ready-to-edit sample.
    const esc = s => s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    const body = `HƯỚNG DẪN: Mỗi câu hỏi viết theo đúng mẫu dưới đây (giữ nguyên các dòng). Nội dung câu hỏi hoặc đáp án có thể xuống dòng, hệ thống vẫn đọc đủ.

Câu 1: Nội dung câu hỏi viết ở đây?
A. Đáp án A
B. Đáp án B
C. Đáp án C
D. Đáp án D
Đáp án: B
Chủ đề: Nội quy
Mức độ: Dễ

Câu 2: Câu hỏi thứ hai?
A. Đáp án A
B. Đáp án B
C. Đáp án C
D. Đáp án D
Đáp án: A
Chủ đề: An toàn lao động
Mức độ: TB

LƯU Ý:
- Đáp án điền chữ cái: A, B, C hoặc D (tối thiểu 2 đáp án).
- Mức độ: Dễ / TB / Khó.
- Chủ đề: điền tên chủ đề bất kỳ.
- Lưu file dưới dạng .docx trước khi import.`;
    const paragraphs = body.split('\n').map(l => `<p style="margin:0">${esc(l) || '&nbsp;'}</p>`).join('');
    const html = `<!DOCTYPE html><html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>Mẫu import câu hỏi</title></head><body style="font-family:'Times New Roman',serif;font-size:13pt">${paragraphs}</body></html>`;
    const blob = new Blob(['﻿', html], {type:'application/msword'});
    const a = document.createElement('a'); a.href=URL.createObjectURL(blob);
    a.download='mau_import_cau_hoi.doc'; a.click();
    URL.revokeObjectURL(a.href);
  };

  return (
    <div>
      {/* Preview modal */}
      {previewList && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl w-full max-w-xl w-full max-h-[85vh] flex flex-col shadow-2xl">
            <div className="flex items-center justify-between p-5 border-b">
              <div>
                <h2 className="font-bold text-slate-800">Xem trước câu hỏi import</h2>
                <p className="text-xs text-slate-500 mt-0.5">Tìm thấy <span className="font-semibold text-emerald-600">{previewList.length} câu hỏi</span> — kiểm tra trước khi thêm vào ngân hàng</p>
              </div>
              <button onClick={()=>setPreviewList(null)} className="text-slate-400 hover:text-slate-600"><X size={18}/></button>
            </div>
            <div className="flex-1 overflow-y-auto p-4 space-y-3">
              {previewList.map((q,i)=>(
                <div key={i} className="bg-slate-50 rounded-xl p-4 border border-slate-100">
                  <div className="flex gap-2 mb-2">
                    <span className="text-xs font-mono text-slate-400">#{i+1}</span>
                  </div>
                  <p className="text-sm font-medium text-slate-700 mb-2">{q.text}</p>
                  <div className="grid grid-cols-2 gap-1">
                    {q.opts.map((o,j)=>(
                      <div key={j} className={`text-xs px-2.5 py-1.5 rounded-lg ${j===q.ans?'bg-emerald-50 text-emerald-700 font-medium border border-emerald-200':'bg-white text-slate-500 border border-slate-100'}`}>
                        {String.fromCharCode(65+j)}. {o}
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
            <div className="p-4 border-t flex gap-2">
              <button onClick={()=>setPreviewList(null)} className="flex-1 py-2.5 border border-slate-200 rounded-xl text-sm text-slate-600 hover:bg-slate-50">Hủy</button>
              <button onClick={confirmImport} className="flex-1 py-2.5 bg-emerald-500 text-white rounded-xl text-sm font-medium hover:bg-emerald-600">
                Thêm {previewList.length} câu hỏi vào ngân hàng
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-5 gap-3">
        <div><h1 className="text-lg md:text-xl font-bold text-slate-800">Ngân hàng câu hỏi</h1><p className="text-slate-500 text-sm">{questions.length} câu hỏi</p></div>
        <div className="flex gap-2">
          <button onClick={downloadWordTemplate} className="flex items-center gap-2 px-3 py-2 border border-slate-200 rounded-lg text-sm text-slate-600 hover:bg-slate-50">
            <Download size={14}/>File mẫu
          </button>
          <button onClick={()=>wordRef.current.click()} disabled={importing} className="flex items-center gap-2 px-3 py-2 border border-slate-200 rounded-lg text-sm text-slate-600 hover:bg-slate-50 disabled:opacity-60">
            <Upload size={14}/>{importing?'Đang đọc...':'Import Word'}
          </button>
          <input ref={wordRef} type="file" accept=".docx" className="hidden" onChange={handleWordImport}/>
          <button onClick={()=>setModal(true)} className="flex items-center gap-2 bg-emerald-600 text-white px-4 py-2 rounded-lg hover:bg-emerald-700 text-sm font-medium"><Plus size={15}/>Thêm câu hỏi</button>
        </div>
      </div>

      {importResult && (
        <div className={`mb-4 rounded-xl p-4 flex items-start gap-3 ${importResult.error?'bg-red-50 border border-red-200':'bg-emerald-50 border border-emerald-200'}`}>
          {importResult.error ? <AlertCircle size={16} className="text-red-500 flex-shrink-0 mt-0.5"/> : <CheckCircle size={16} className="text-emerald-600 flex-shrink-0 mt-0.5"/>}
          <p className={`text-sm flex-1 ${importResult.error?'text-red-700':'text-emerald-800 font-medium'}`}>
            {importResult.error || `Đã thêm thành công ${importResult.added} câu hỏi vào ngân hàng!`}
          </p>
          <button onClick={()=>setImportResult(null)} className="text-slate-400 hover:text-slate-600"><X size={14}/></button>
        </div>
      )}
      {modal && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl w-full max-w-lg max-h-[90vh] overflow-y-auto shadow-2xl">
            <div className="flex items-center justify-between p-5 border-b sticky top-0 bg-white z-10">
              <h2 className="font-bold text-slate-800">Thêm câu hỏi mới</h2>
              <button onClick={()=>{setModal(false);setCustomTopic('');setNq({topic:'Nội quy',level:'Dễ',text:'',opts:['','','',''],ans:0});}} className="text-slate-400 hover:text-slate-600"><X size={18}/></button>
            </div>
            <div className="p-5 space-y-4">
              <div className="grid grid-cols-2 gap-3">
                <div><label className="text-xs font-medium text-slate-600 mb-1 block">Chủ đề</label>
                  <select className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" value={nq.topic} onChange={e=>setNq({...nq,topic:e.target.value})}>
                    {topics.map(t=><option key={t} value={t}>{t}</option>)}
                    <option value="__custom__">+ Tạo chủ đề mới...</option>
                  </select>
                  {nq.topic==='__custom__' && (
                    <input
                      autoFocus
                      className="mt-2 w-full border border-emerald-300 bg-emerald-50 rounded-lg px-3 py-2 text-sm focus:outline-none focus:border-emerald-500"
                      placeholder="Nhập tên chủ đề mới"
                      value={customTopic}
                      onChange={e=>setCustomTopic(e.target.value)}
                    />
                  )}
                </div>
              </div>
              <div><label className="text-xs font-medium text-slate-600 mb-1 block">Nội dung câu hỏi</label>
                <textarea className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none" rows={3} value={nq.text} onChange={e=>setNq({...nq,text:e.target.value})} placeholder="Nhập nội dung câu hỏi..."/></div>
              <div><label className="text-xs font-medium text-slate-600 mb-2 block">Các đáp án (chọn đáp án đúng)</label>
                <div className="space-y-2">
                  {['A','B','C','D'].map((lt,i)=>(
                    <div key={i} className="flex items-center gap-2">
                      <button onClick={()=>setNq({...nq,ans:i})} className={`w-7 h-7 rounded-full flex-shrink-0 border-2 flex items-center justify-center text-xs font-bold transition-all ${nq.ans===i?'border-emerald-600 bg-emerald-600 text-white':'border-slate-300 text-slate-400 hover:border-emerald-500'}`}>{lt}</button>
                      <input className="flex-1 border border-slate-200 rounded-lg px-3 py-1.5 text-sm" value={nq.opts[i]} onChange={e=>{const o=[...nq.opts];o[i]=e.target.value;setNq({...nq,opts:o});}} placeholder={`Đáp án ${lt}`}/>
                    </div>
                  ))}
                </div>
              </div>
              <div className="flex gap-2 pt-1">
                <button onClick={()=>{setModal(false);setCustomTopic('');setNq({topic:'Nội quy',level:'Dễ',text:'',opts:['','','',''],ans:0});}} className="flex-1 py-2 border border-slate-200 rounded-lg text-sm text-slate-600 hover:bg-slate-50">Hủy</button>
                <button onClick={add} className="flex-1 py-2 bg-emerald-600 text-white rounded-lg text-sm hover:bg-emerald-700">Thêm câu hỏi</button>
              </div>
            </div>
          </div>
        </div>
      )}
      <div className="space-y-3">
        {orderedQuestions(questions).map((q,idx)=>(
          <div key={q.id} className="bg-white rounded-xl p-4 shadow-sm border border-slate-100">
            <div className="flex items-start justify-between gap-3">
              <div className="flex-1">
                <div className="flex items-center gap-2 mb-2">
                  <span className="text-xs font-mono text-slate-400">#{idx+1}</span>
                </div>
                <p className="text-sm font-medium text-slate-700 mb-2">{q.text}</p>
                <div className="grid grid-cols-2 gap-1.5">
                  {q.opts.map((o,i)=>(
                    <div key={i} className={`text-xs px-2.5 py-1.5 rounded-lg ${i===q.ans?'bg-emerald-50 text-emerald-700 font-medium border border-emerald-200':'bg-slate-50 text-slate-500'}`}>
                      {String.fromCharCode(65+i)}. {o}
                    </div>
                  ))}
                </div>
              </div>
              <button onClick={()=>setQuestions(p=>p.filter(x=>x.id!==q.id))} className="text-slate-300 hover:text-red-400 flex-shrink-0 mt-1"><Trash2 size={16}/></button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

// ── EXAMS ──
const Exams = ({exams, setExams, questions}) => {
  const [modal, setModal] = useState(false);
  const [ne, setNe] = useState({title:'',desc:'',qIds:[],time:20,pass:70});
  const toggleQ = id => setNe(p=>({...p,qIds:p.qIds.includes(id)?p.qIds.filter(x=>x!==id):[...p.qIds,id]}));
  const create = () => {
    if(!ne.title.trim()||ne.qIds.length===0) return;
    setExams(p=>[...p,{...ne,id:Date.now()}]); setModal(false);
    setNe({title:'',desc:'',qIds:[],time:20,pass:70});
  };
  return (
    <div>
      <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-5 gap-3">
        <div><h1 className="text-lg md:text-xl font-bold text-slate-800">Quản lý đề thi</h1><p className="text-slate-500 text-sm">{exams.length} đề thi</p></div>
        <button onClick={()=>setModal(true)} className="flex items-center gap-2 bg-emerald-600 text-white px-4 py-2 rounded-lg hover:bg-emerald-700 text-sm font-medium"><Plus size={15}/>Tạo đề thi</button>
      </div>
      {modal && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl w-full max-w-lg max-h-[90vh] overflow-y-auto shadow-2xl">
            <div className="flex items-center justify-between p-5 border-b sticky top-0 bg-white z-10">
              <h2 className="font-bold text-slate-800">Tạo đề thi mới</h2>
              <button onClick={()=>setModal(false)} className="text-slate-400 hover:text-slate-600"><X size={18}/></button>
            </div>
            <div className="p-5 space-y-4">
              <div><label className="text-xs font-medium text-slate-600 mb-1 block">Tên đề thi</label>
                <input className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" value={ne.title} onChange={e=>setNe({...ne,title:e.target.value})} placeholder="VD: Kiểm tra đầu vào 2024"/></div>
              <div><label className="text-xs font-medium text-slate-600 mb-1 block">Mô tả</label>
                <textarea className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none" rows={2} value={ne.desc} onChange={e=>setNe({...ne,desc:e.target.value})} placeholder="Mô tả ngắn về bài thi..."/></div>
              <div className="grid grid-cols-2 gap-3">
                <div><label className="text-xs font-medium text-slate-600 mb-1 block">Thời gian (phút)</label>
                  <input type="number" className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" value={ne.time} onChange={e=>setNe({...ne,time:+e.target.value})}/></div>
                <div><label className="text-xs font-medium text-slate-600 mb-1 block">Điểm đạt (%)</label>
                  <input type="number" className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" value={ne.pass} onChange={e=>setNe({...ne,pass:+e.target.value})}/></div>
              </div>
              <div><label className="text-xs font-medium text-slate-600 mb-2 block">Chọn câu hỏi ({ne.qIds.length} đã chọn)</label>
                <div className="space-y-1 max-h-48 overflow-y-auto border border-slate-100 rounded-xl p-2">
                  {orderedQuestions(questions).map(q=>(
                    <label key={q.id} className="flex items-start gap-2 p-2 rounded-lg hover:bg-slate-50 cursor-pointer">
                      <input type="checkbox" className="mt-0.5 flex-shrink-0" checked={ne.qIds.includes(q.id)} onChange={()=>toggleQ(q.id)}/>
                      <div><p className="text-xs text-slate-700">{q.text}</p></div>
                    </label>
                  ))}
                </div>
              </div>
              <div className="flex gap-2 pt-1">
                <button onClick={()=>setModal(false)} className="flex-1 py-2 border border-slate-200 rounded-lg text-sm text-slate-600">Hủy</button>
                <button onClick={create} className="flex-1 py-2 bg-emerald-600 text-white rounded-lg text-sm">Tạo đề thi</button>
              </div>
            </div>
          </div>
        </div>
      )}
      <div className="space-y-3">
        {exams.map(exam=>(
          <div key={exam.id} className="bg-white rounded-xl p-5 shadow-sm border border-slate-100">
            <div className="flex items-start justify-between gap-4">
              <div className="flex-1">
                <h3 className="font-semibold text-slate-800 mb-1">{exam.title}</h3>
                <p className="text-sm text-slate-500 mb-3">{exam.desc}</p>
                <div className="flex gap-4 text-xs text-slate-500">
                  <span className="flex items-center gap-1"><FileText size={12}/>{exam.qIds.length} câu hỏi</span>
                  <span className="flex items-center gap-1"><Clock size={12}/>{exam.time} phút</span>
                  <span className="flex items-center gap-1"><Award size={12}/>Điểm đạt: {exam.pass}%</span>
                </div>
              </div>
              <button onClick={()=>setExams(p=>p.filter(e=>e.id!==exam.id))} className="text-slate-300 hover:text-red-400"><Trash2 size={16}/></button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

// ── EMPLOYEES ──
const EmployeesView = ({employees, setEmployees, results, exams}) => {
  const [importing, setImporting] = useState(false);
  const [importResult, setImportResult] = useState(null);
  const [modal, setModal] = useState(null); // null | {mode:'add'} | {mode:'edit', emp} | {mode:'delete', emp}
  const [form, setForm] = useState({name:'', dept:''});
  const [customDept, setCustomDept] = useState('');
  const [formErr, setFormErr] = useState('');
  const fileRef = useRef();

  const depts = [...new Set(employees.map(e=>e.dept))].sort();

  const [renamingDept, setRenamingDept] = useState(null);
  const [renameVal, setRenameVal]       = useState('');
  const startRename = dept => { setRenamingDept(dept); setRenameVal(dept); };
  const commitRename = () => {
    const newName = renameVal.trim();
    if(newName && newName !== renamingDept)
      setEmployees(p => p.map(e => e.dept===renamingDept ? {...e, dept:newName} : e));
    setRenamingDept(null);
  };

  const openAdd = () => { 
    setForm({name:'', dept: depts.length > 0 ? depts[0] : '__custom__'}); 
    setCustomDept(''); 
    setFormErr(''); 
    setModal({mode:'add'}); 
  };
  const openEdit = (emp) => { setForm({name:emp.name, dept:emp.dept}); setCustomDept(''); setFormErr(''); setModal({mode:'edit', emp}); };
  const openDelete = (emp) => setModal({mode:'delete', emp});
  const closeModal = () => { setModal(null); setFormErr(''); };

  const effectiveDept = form.dept==='__custom__' ? customDept.trim() : form.dept;

  const handleSave = () => {
    const name = form.name.trim();
    const dept = form.dept === '__custom__' ? customDept.trim() : form.dept;
    if(!name) return setFormErr('Vui lòng nhập họ và tên');
    if(!dept) return setFormErr('Vui lòng chọn hoặc nhập đơn vị');
    if(modal.mode==='add') {
      const dup = employees.find(e=>e.name.toLowerCase()===name.toLowerCase());
      if(dup) return setFormErr('Thí sinh này đã tồn tại trong hệ thống');
      setEmployees(p=>[...p, {id:Date.now(), name, dept}]);
    } else {
      const dup = employees.find(e=>e.name.toLowerCase()===name.toLowerCase() && e.id!==modal.emp.id);
      if(dup) return setFormErr('Tên này đã được dùng cho thí sinh khác');
      setEmployees(p=>p.map(e=>e.id===modal.emp.id?{...e,name,dept}:e));
    }
    closeModal();
  };

  const handleDelete = () => {
    setEmployees(p=>p.filter(e=>e.id!==modal.emp.id));
    closeModal();
  };

 const handleImport = (e) => {
  const file = e.target.files[0];
  if(!file) return;
  setImporting(true);
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const wb = XLSX.read(ev.target.result, {type:'binary'});
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, {defval:'', header:1});
      const added = [];
      const skipped = [];

      // Tìm index cột Họ và tên, Đơn vị
      let colName = -1, colDept = -1;
      for(let i = 0; i < Math.min(5, rows.length); i++) {
        const row = rows[i].map(c => String(c).trim().toLowerCase());
        const ni = row.findIndex(c => c.includes('họ') || c.includes('tên') || c === 'name');
        const di = row.findIndex(c => c.includes('đơn vị') || c.includes('phòng') || c.includes('dept') || c.includes('khoa'));
        if(ni !== -1) { colName = ni; colDept = di !== -1 ? di : ni+1; break; }
      }
      if(colName === -1) { colName = 1; colDept = 2; }

      rows.forEach(row => {
        const nameRaw = String(row[colName]||'').trim();
        const deptRaw = String(colDept >= 0 ? (row[colDept]||'') : '').trim();
        if(!nameRaw) return;
        if(nameRaw.toUpperCase().includes('KHỐI')) return;
        if(nameRaw.toLowerCase().includes('họ và tên') || nameRaw.toLowerCase() === 'name') return;
        const dup = employees.find(e => e.name.toLowerCase() === nameRaw.toLowerCase());
        if(dup) { skipped.push(`${nameRaw} (trùng)`); return; }
        added.push({id: Date.now() + Math.random(), name: nameRaw, dept: deptRaw || 'Chưa phân công'});
      });

      if(added.length > 0) setEmployees(p => [...p, ...added]);
      setImportResult({added: added.length, skipped: skipped.length, names: added.map(e=>e.name)});
      setTimeout(() => setImportResult(null), 3000);
    } catch(err) {
      setImportResult({error: 'Không đọc được file. Vui lòng kiểm tra định dạng.'});
    }
    setImporting(false);
  };
  reader.readAsBinaryString(file);
  e.target.value = '';
};

const downloadTemplate = () => {
  const ws = XLSX.utils.aoa_to_sheet([
    ['TT', 'Họ và tên', 'Đơn vị'],
    ['KHỐI CƠ QUAN', '', ''],
    [1, 'Nguyễn Thị A', 'Phòng KHTH'],
    [2, 'Trần Văn B', 'Ban ĐD'],
    ['KHỐI NỘI', '', ''],
    [3, 'Lê Thị C', 'Khoa A1'],
    [4, 'Phạm Văn D', 'Khoa A2'],
  ]);
  ws['!cols'] = [{wch:6},{wch:25},{wch:20}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Thí sinh');
  XLSX.writeFile(wb, 'mau_import_nhan_vien.xlsx');
};
  const [activeDept, setActiveDept] = useState('all');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const PAGE_SIZE = 10;

  const deptCounts = depts.map(d=>({dept:d, count:employees.filter(e=>e.dept===d).length}));
  const filtered = employees.filter(e=>{
    const matchDept = search ? true : (activeDept==='all' || e.dept===activeDept);
    const matchSearch = e.name.toLowerCase().includes(search.toLowerCase());
    return matchDept && matchSearch;
  });
  const pagedEmps = filtered.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);
  const totalPages = Math.ceil(filtered.length/PAGE_SIZE);

  // Global stats
  const allRs = results;
  const totalPassed = allRs.filter(r=>{const e=exams.find(x=>x.id===r.examId);return e&&r.score>=e.pass;}).length;
  const globalAvg = allRs.length?Math.round(allRs.reduce((a,r)=>a+r.score,0)/allRs.length):0;
  const avgAttempts = employees.length?+(allRs.length/employees.length).toFixed(1):0;
  const passRate = allRs.length?Math.round(totalPassed/allRs.length*100):0;

  const scoreColor = (avg) => avg>=80?'bg-emerald-500':avg>=60?'bg-amber-400':'bg-red-400';

  // ── Render ──
  return (
    <div>
      {/* Modal Thêm / Sửa */}
      {(modal?.mode==='add' || modal?.mode==='edit') && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl w-full max-w-md w-full shadow-2xl">
            <div className="flex items-center justify-between p-5 border-b">
              <h2 className="font-bold text-slate-800">{modal.mode==='add'?'Thêm thí sinh mới':'Chỉnh sửa thí sinh'}</h2>
              <button onClick={closeModal} className="text-slate-400 hover:text-slate-600"><X size={18}/></button>
            </div>
            <div className="p-5 space-y-4">
              <div>
                <label className="text-xs font-medium text-slate-600 mb-1.5 block">Họ và tên <span className="text-red-400">*</span></label>
                <input
                  autoFocus
                  className="w-full border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm focus:outline-none focus:border-emerald-500"
                  placeholder="VD: Nguyễn Văn A"
                  value={form.name}
                  onChange={e=>{ setForm(f=>({...f,name:e.target.value})); setFormErr(''); }}
                />
              </div>
              <div>
                <label className="text-xs font-medium text-slate-600 mb-1.5 block">Đơn vị <span className="text-red-400">*</span></label>
                <select
                  className="w-full border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm focus:outline-none focus:border-emerald-500 mb-2"
                  value={form.dept}
                  onChange={e=>{ setForm(f=>({...f,dept:e.target.value})); setFormErr(''); }}
                >
                  {depts.map(d=><option key={d} value={d}>{d}</option>)}
                  <option value="__custom__">+ Tạo đơn vị mới...</option>
                </select>
                {form.dept==='__custom__' && (
                  <input
                    className="w-full border border-emerald-300 bg-emerald-50 rounded-xl px-3.5 py-2.5 text-sm focus:outline-none focus:border-emerald-500"
                    placeholder="Nhập tên đơn vị mới"
                    value={customDept}
                    onChange={e=>{ setCustomDept(e.target.value); setFormErr(''); }}
                  />
                )}
              </div>
              {formErr && (
                <div className="flex items-center gap-2 bg-red-50 border border-red-200 rounded-xl px-3 py-2">
                  <AlertCircle size={14} className="text-red-500 flex-shrink-0"/>
                  <span className="text-red-600 text-xs">{formErr}</span>
                </div>
              )}
              <div className="flex gap-2 pt-1">
                <button onClick={closeModal} className="flex-1 py-2.5 border border-slate-200 rounded-xl text-sm text-slate-600 hover:bg-slate-50">Hủy</button>
                <button onClick={handleSave} className="flex-1 py-2.5 bg-emerald-600 text-white rounded-xl text-sm font-medium hover:bg-emerald-700">
                  {modal.mode==='add'?'Thêm thí sinh':'Lưu thay đổi'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Modal Xóa */}
      {modal?.mode==='delete' && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl w-full max-w-sm sm:max-w-sm w-full shadow-2xl p-6">
            <div className="w-12 h-12 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-4">
              <Trash2 size={20} className="text-red-500"/>
            </div>
            <h3 className="font-bold text-slate-800 text-center mb-1">Xóa thí sinh?</h3>
            <p className="text-sm text-slate-500 text-center mb-1">
              Bạn có chắc muốn xóa <span className="font-semibold text-slate-700">{modal.emp.name}</span>?
            </p>
            <p className="text-xs text-slate-400 text-center mb-5">Lịch sử thi của thí sinh này sẽ vẫn được giữ lại.</p>
            <div className="flex gap-2">
              <button onClick={closeModal} className="flex-1 py-2.5 border border-slate-200 rounded-xl text-sm text-slate-600 hover:bg-slate-50">Hủy</button>
              <button onClick={handleDelete} className="flex-1 py-2.5 bg-red-500 text-white rounded-xl text-sm font-medium hover:bg-red-600">Xóa thí sinh</button>
            </div>
          </div>
        </div>
      )}

      <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-5 gap-3">
        <div className="flex items-center gap-3">
          <div className="w-11 h-11 bg-emerald-100 rounded-xl flex items-center justify-center"><Users size={22} className="text-emerald-600"/></div>
          <div><h1 className="text-lg md:text-xl font-bold text-slate-800">Thí sinh</h1><p className="text-slate-500 text-xs">{employees.length} thí sinh</p></div>
        </div>
        <div className="flex flex-wrap gap-2">
          <button onClick={downloadTemplate} className="flex items-center gap-1.5 px-3 py-2 border border-slate-200 rounded-lg text-sm text-slate-600 hover:bg-slate-50"><Download size={14}/>Tải file mẫu</button>
          <button onClick={()=>fileRef.current.click()} disabled={importing} className="flex items-center gap-1.5 px-3 py-2 border border-slate-200 rounded-lg text-sm text-slate-600 hover:bg-slate-50 disabled:opacity-60"><Upload size={14}/>{importing?'Đang import...':'Import Excel'}</button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleImport}/>
          <button onClick={openAdd} className="flex items-center gap-1.5 bg-emerald-600 text-white px-4 py-2 rounded-lg hover:bg-emerald-700 text-sm font-medium"><Plus size={15}/>Thêm thí sinh</button>
        </div>
      </div>

      {importResult && (
        <div className={`mb-4 rounded-xl p-3 flex items-center gap-3 ${importResult.error?'bg-red-50 border border-red-200':'bg-emerald-50 border border-emerald-200'}`}>
          {importResult.error ? <AlertCircle size={15} className="text-red-500"/> : <CheckCircle size={15} className="text-emerald-600"/>}
          <p className={`text-sm flex-1 ${importResult.error?'text-red-700':'text-emerald-800 font-medium'}`}>
            {importResult.error || `Thêm ${importResult.added} thí sinh thành công!`}
          </p>
          <button onClick={()=>setImportResult(null)} className="text-slate-400 hover:text-slate-600"><X size={14}/></button>
        </div>
      )}

      {/* Stats */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-5">
        {[
          {ic:<Users size={18}/>, val:employees.length, lb:'Tổng thí sinh', col:'bg-emerald-100 text-emerald-600'},
          {ic:<CheckCircle size={18}/>, val:<span>{totalPassed} <span className="text-xs font-normal text-slate-400">{passRate}%</span></span>, lb:'Đạt', col:'bg-emerald-100 text-emerald-600'},
          {ic:<TrendingUp size={18}/>, val:`${globalAvg}%`, lb:'Điểm TB chung', col:'bg-amber-100 text-amber-600'},
          {ic:<FileText size={18}/>, val:avgAttempts, lb:'Lượt thi TB', col:'bg-violet-100 text-violet-600'},
        ].map((s,i)=>(
          <div key={i} className="bg-white rounded-xl p-4 shadow-sm border border-slate-100 flex items-center gap-3">
            <div className={`w-10 h-10 ${s.col} rounded-xl flex items-center justify-center flex-shrink-0`}>{s.ic}</div>
            <div><div className="text-xl font-bold text-slate-800">{s.val}</div><div className="text-xs text-slate-500">{s.lb}</div></div>
          </div>
        ))}
      </div>

      {/* Dept filter tabs + search */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-4 gap-3 flex-wrap gap-2">
        <div className="flex gap-2 flex-wrap">
          {deptCounts.map(({dept,count})=>(
            <button key={dept} onClick={()=>{setActiveDept(dept);setPage(1);}}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-full text-sm font-medium border transition-all ${activeDept===dept?'bg-emerald-600 text-white border-emerald-600':'bg-white text-slate-600 border-slate-200 hover:border-emerald-300'}`}>
              {dept}<span className={`text-xs px-1.5 py-0.5 rounded-full ml-0.5 ${activeDept===dept?'bg-white/25 text-white':'bg-slate-100 text-slate-500'}`}>{count}</span>
            </button>
          ))}
        </div>
        <div className="flex items-center gap-2 border border-slate-200 rounded-lg px-3 py-1.5 bg-white">
          <FileText size={13} className="text-slate-400"/>
          <input className="text-sm outline-none w-44 text-slate-700 placeholder-slate-400" placeholder="Tìm kiếm thí sinh..." value={search} onChange={e=>{setSearch(e.target.value);setPage(1);}}/>
        </div>
      </div>

      {/* Table */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-x-auto">
        <table className="w-full min-w-[640px]">
          <thead>
            <tr className="border-b border-slate-100">
              {['Thí sinh','Lượt thi','Đạt','Điểm TB','Kết quả gần nhất','Thao tác'].map(h=>(
                <th key={h} className={`px-5 py-3 text-xs font-semibold text-slate-400 uppercase tracking-wide ${h==='Kết quả gần nhất'||h==='Thao tác'?'text-center':'text-left'}`}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {depts.filter(d=>activeDept==='all'||d===activeDept).map(dept=>{
              const deptEmps = pagedEmps.filter(e=>e.dept===dept);
              if(!deptEmps.length) return null;
              return [
                <tr key={`hd-${dept}`} className="bg-slate-50/80 border-t border-slate-100">
                  <td colSpan={6} className="px-5 py-2">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <div className="w-6 h-6 bg-emerald-100 rounded-lg flex items-center justify-center"><Users size={11} className="text-emerald-500"/></div>
                        {renamingDept===dept ? (
                          <div className="flex items-center gap-2">
                            <input autoFocus className="border border-emerald-400 bg-emerald-50 rounded px-2 py-0.5 text-sm font-semibold focus:outline-none w-36" value={renameVal} onChange={e=>setRenameVal(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')commitRename();if(e.key==='Escape')setRenamingDept(null);}}/>
                            <button onClick={commitRename} className="px-2 py-0.5 bg-emerald-600 text-white rounded text-xs">Lưu</button>
                            <button onClick={()=>setRenamingDept(null)} className="px-2 py-0.5 border border-slate-200 rounded text-xs text-slate-500">Hủy</button>
                          </div>
                        ) : (
                          <>
                            <span className="text-sm font-semibold text-slate-700">{dept}</span>
                            <button onClick={()=>startRename(dept)} className="text-xs text-slate-400 hover:text-emerald-600 ml-1 px-1.5 py-0.5 rounded hover:bg-emerald-50">✎ Đổi tên</button>
                          </>
                        )}
                      </div>
                      <span className="text-xs text-slate-400">{deptEmps.length} thí sinh</span>
                    </div>
                  </td>
                </tr>,
                ...deptEmps.map(emp=>{
                  const rs = results.filter(r=>r.empId===emp.id);
                  const passed = rs.filter(r=>{const e=exams.find(x=>x.id===r.examId);return e&&r.score>=e.pass;}).length;
                  const avg = rs.length?Math.round(rs.reduce((a,r)=>a+r.score,0)/rs.length):0;
                  const last = rs.slice(-1)[0];
                  const lastExam = last&&exams.find(e=>e.id===last.examId);
                  const lastOk = lastExam&&last.score>=lastExam.pass;
                  return (
                    <tr key={emp.id} className="border-t border-slate-50 hover:bg-slate-50/50 transition-colors">
                      <td className="px-5 py-3 text-sm font-medium text-slate-800">{emp.name}</td>
                      <td className="px-5 py-3 text-sm text-slate-600">{rs.length}</td>
                      <td className="px-5 py-3 text-sm font-semibold text-emerald-600">{passed}</td>
                      <td className="px-5 py-3">
                        <div className="flex items-center gap-2">
                          <span className="text-sm font-bold text-slate-800 w-10">{avg?`${avg}%`:'--'}</span>
                          {avg>0&&<div className="flex-1 h-1.5 bg-slate-100 rounded-full overflow-hidden min-w-16"><div className={`h-full rounded-full ${scoreColor(avg)}`} style={{width:`${avg}%`}}/></div>}
                        </div>
                      </td>
                      <td className="px-5 py-3 text-center">
                        {last ? (
                          <div>
                            <div className={`text-sm font-medium ${lastOk?'text-emerald-600':'text-red-500'}`}>{lastOk?'✓ Đạt':'✗ Chưa đạt'}</div>
                            <div className="text-xs text-slate-400">{last.date}</div>
                          </div>
                        ) : <span className="text-xs text-slate-400">Chưa thi</span>}
                      </td>
                      <td className="px-5 py-3">
                        <div className="flex items-center justify-center gap-1.5">
                          <button onClick={()=>openEdit(emp)} className="w-8 h-8 flex items-center justify-center rounded-lg border border-slate-200 text-slate-500 hover:border-emerald-300 hover:text-emerald-600 transition-all"><FileText size={13}/></button>
                          <button onClick={()=>openDelete(emp)} className="w-8 h-8 flex items-center justify-center rounded-lg border border-slate-200 text-slate-500 hover:border-red-300 hover:text-red-600 transition-all"><Trash2 size={13}/></button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              ];
            })}
          </tbody>
        </table>

        {filtered.length===0 && (
          <div className="text-center py-16 text-slate-400">
            <Users size={36} className="mx-auto mb-3 opacity-20"/>
            <p className="text-sm">Không tìm thấy thí sinh nào</p>
          </div>
        )}

        {totalPages>1 && (
          <div className="flex items-center justify-center gap-2 py-4 border-t border-slate-100">
            <span className="text-xs text-slate-500 mr-2">Hiển thị {filtered.length} thí sinh</span>
            <button onClick={()=>setPage(p=>Math.max(1,p-1))} disabled={page===1} className="w-7 h-7 flex items-center justify-center rounded-lg border border-slate-200 text-slate-500 hover:border-emerald-300 disabled:opacity-40">‹</button>
            {Array.from({length:totalPages},(_,i)=>i+1).map(p=>(
              <button key={p} onClick={()=>setPage(p)} className={`w-7 h-7 rounded-lg text-xs font-medium ${p===page?'bg-emerald-600 text-white':'border border-slate-200 text-slate-500 hover:border-emerald-300'}`}>{p}</button>
            ))}
            <button onClick={()=>setPage(p=>Math.min(totalPages,p+1))} disabled={page===totalPages} className="w-7 h-7 flex items-center justify-center rounded-lg border border-slate-200 text-slate-500 hover:border-emerald-300 disabled:opacity-40">›</button>
          </div>
        )}
      </div>
    </div>
  );
};

const Reports = ({results, exams, employees}) => {
  const depts = [...new Set(employees.map(e=>e.dept))];

  const deptData = depts.map(dept=>{
    const emps = employees.filter(e=>e.dept===dept);
    const rs = results.filter(r=>emps.some(e=>e.id===r.empId));
    const passed = rs.filter(r=>{const e=exams.find(x=>x.id===r.examId);return e&&r.score>=e.pass;}).length;
    const avg = rs.length?Math.round(rs.reduce((a,r)=>a+r.score,0)/rs.length):0;
    const rate = rs.length?Math.round(passed/rs.length*100):0;
    return {name:dept, attempts:rs.length, passed, failed:rs.length-passed, avg, rate};
  });

  const totalAttempts = results.length;
  const globalAvg = totalAttempts?Math.round(results.reduce((a,r)=>a+r.score,0)/totalAttempts):0;
  const bestDept = deptData.reduce((a,b)=>b.avg>a.avg?b:a, deptData[0]||{name:'--',avg:0});

  // Distribution: Đạt ≥ pass%, Trung bình 50-69%, Chưa đạt <50%
  const dat = results.filter(r=>{const e=exams.find(x=>x.id===r.examId);return e&&r.score>=e.pass;}).length;
  const trungBinh = results.filter(r=>{const e=exams.find(x=>x.id===r.examId);return r.score>=50&&(!e||r.score<e.pass);}).length;
  const chuaDat = results.filter(r=>r.score<50).length;
  const pieData = [
    {name:`Đạt (≥ 70%)`, value:dat, c:'#22c55e'},
    {name:`Trung bình (50–69%)`, value:trungBinh, c:'#f59e0b'},
    {name:`Chưa đạt (< 50%)`, value:chuaDat, c:'#ef4444'},
  ];

  const barColors = ['#15803d','#0d9488','#f59e0b','#3b82f6','#8b5cf6','#ec4899'];
  const barData = deptData.map((d,i)=>({...d, fill:barColors[i%barColors.length]}));

  const trends = ['+15%','-10%','+8%','+20%'];

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Tổng hợp phòng  ban
    const s1 = XLSX.utils.aoa_to_sheet([
      ['BÁO CÁO KẾT QUẢ THI - Cục hậu cần - kỹ thuật Quân khu 4'],
      ['Xuất ngày: ' + new Date().toLocaleDateString('vi-VN')],
      [],
      ['Đơn vị','Lượt thi','Số đạt','Số chưa đạt','Điểm TB (%)','Tỉ lệ đạt (%)'],
      ...deptData.map(d=>[d.name, d.attempts, d.passed, d.failed, d.avg, d.rate]),
    ]);
    s1['!cols']=[{wch:22},{wch:12},{wch:12},{wch:16},{wch:14},{wch:14}];
    s1['!merges']=[{s:{r:0,c:0},e:{r:0,c:5}}];
    XLSX.utils.book_append_sheet(wb, s1, 'Tổng hợp đơn vị');

    // Sheet 2: Chi tiết kết quả
    const s2rows = [['Họ và tên','Đơn vị','Đề thi','Điểm (%)','Số câu đúng','Tổng câu','Kết quả','Điểm đạt yêu cầu (%)','Thời gian làm bài','Ngày thi']];
    results.forEach(r=>{
      const emp = employees.find(e=>e.id===r.empId);
      const exam = exams.find(e=>e.id===r.examId);
      const ok = exam && r.score >= exam.pass;
      const mins = r.timeTaken!=null ? Math.floor(r.timeTaken/60) : null;
      const secs = r.timeTaken!=null ? r.timeTaken%60 : null;
      const timeStr = mins!=null ? `${mins}p${String(secs).padStart(2,'0')}s` : '--';
      s2rows.push([emp?.name||'', emp?.dept||'', exam?.title||'', r.score, r.correct, exam?.qIds?.length||'', ok?'Đạt':'Không đạt', exam?.pass||'', timeStr, r.date]);
    });
    const s2 = XLSX.utils.aoa_to_sheet(s2rows);
    s2['!cols']=[{wch:22},{wch:18},{wch:30},{wch:12},{wch:14},{wch:12},{wch:14},{wch:20},{wch:18},{wch:14}];
    XLSX.utils.book_append_sheet(wb, s2, 'Chi tiết kết quả');

    // Sheet 3: Danh sách thí sinh
    const s3rows = [['Họ và tên','Đơn vị','Tổng lượt thi','Số bài đạt','Số bài chưa đạt','Điểm TB (%)','Kết quả gần nhất','Ngày thi gần nhất','Trạng thái']];
    employees.forEach(emp=>{
      const rs = results.filter(r=>r.empId===emp.id);
      const passed = rs.filter(r=>{const e=exams.find(x=>x.id===r.examId);return e&&r.score>=e.pass;}).length;
      const avg = rs.length ? Math.round(rs.reduce((a,r)=>a+r.score,0)/rs.length) : 0;
      const last = rs.slice(-1)[0];
      const lastExam = last && exams.find(e=>e.id===last.examId);
      const lastOk = lastExam && last.score >= lastExam.pass;
      s3rows.push([emp.name, emp.dept, rs.length, passed, rs.length-passed, avg||'--', last?(lastOk?'Đạt':'Không đạt'):'Chưa thi', last?.date||'--', rs.length===0?'Chưa thi':'Đang học']);
    });
    const s3 = XLSX.utils.aoa_to_sheet(s3rows);
    s3['!cols']=[{wch:22},{wch:18},{wch:16},{wch:14},{wch:18},{wch:14},{wch:18},{wch:18},{wch:14}];
    XLSX.utils.book_append_sheet(wb, s3, 'Danh sách thí sinh');

    XLSX.writeFile(wb, `BaoCao_BenhVienQuanY4_${new Date().toLocaleDateString('vi-VN').replace(/\//g,'-')}.xlsx`);
  };

  const worstDept = deptData.reduce((a,b)=>b.avg<a.avg?b:a, deptData[0]||{name:'--'});

  return (
    <div>
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-5 gap-3">
        <div className="flex items-center gap-3">
          <div className="w-11 h-11 bg-emerald-100 rounded-xl flex items-center justify-center"><TrendingUp size={20} className="text-emerald-600"/></div>
          <div><h1 className="text-lg md:text-xl font-bold text-slate-800">Báo cáo & Phân tích</h1><p className="text-slate-500 text-xs">Phân tích kết quả thi theo đơn vị</p></div>
        </div>
        <button onClick={exportExcel} className="flex items-center gap-2 bg-emerald-600 text-white px-4 py-2 rounded-lg hover:bg-emerald-700 text-sm font-medium">
          <Download size={15}/>Xuất Excel <ChevronRight size={14} className="rotate-90"/>
        </button>
      </div>

      {/* Stat cards with sparklines */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-5">
        {[
          {ic:<BarChart2 size={18}/>, val:depts.length, lb:'Đơn vị', sub:'Đã có dữ liệu', icCol:'bg-teal-100 text-teal-600', spark:'#5eead4', path:'M0,20 C10,15 20,25 30,18 C40,10 50,22 60,16 C70,10 80,20 90,14'},
          {ic:<CheckCircle size={18}/>, val:totalAttempts, lb:'Tổng lượt thi', sub:'Trong kỳ', icCol:'bg-emerald-100 text-emerald-600', spark:'#6ee7b7', path:'M0,22 C15,18 25,24 40,16 C55,8 65,20 80,14 C85,12 88,16 90,13'},
          {ic:<TrendingUp size={18}/>, val:`${globalAvg}%`, lb:'Điểm TB chung', sub:'Toàn hệ thống', icCol:'bg-amber-100 text-amber-600', spark:'#fcd34d', path:'M0,24 C10,20 20,22 35,14 C50,6 60,18 75,12 C82,9 87,15 90,10'},
          {ic:<Award size={18}/>, val:`${bestDept.avg}%`, lb:'Đơn vị cao nhất', sub:bestDept.name, icCol:'bg-violet-100 text-violet-600', spark:'#c4b5fd', path:'M0,20 C12,16 22,22 38,12 C54,2 62,18 78,10 C84,7 88,13 90,8'},
        ].map((s,i)=>(
          <div key={i} className="bg-white rounded-2xl p-4 shadow-sm border border-slate-100 overflow-hidden relative">
            <div className="flex items-start gap-3 mb-3">
              <div className={`w-10 h-10 ${s.icCol} rounded-xl flex items-center justify-center flex-shrink-0`}>{s.ic}</div>
              <div>
                <div className="text-2xl font-bold text-slate-800">{s.val}</div>
                <div className="text-xs font-medium text-slate-600">{s.lb}</div>
                <div className="text-xs text-slate-400">{s.sub}</div>
              </div>
            </div>
            <svg viewBox="0 0 90 30" className="w-full h-8" preserveAspectRatio="none">
              <path d={s.path} fill="none" stroke={s.spark} strokeWidth="2.5" strokeLinecap="round"/>
            </svg>
          </div>
        ))}
      </div>

      {/* Charts row */}
      <div className="grid grid-cols-1 lg:grid-cols-5 gap-4 mb-5">
        {/* Bar chart */}
        <div className="md:lg:col-span-3 bg-white rounded-2xl p-5 shadow-sm border border-slate-100">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-4 gap-3">
            <div className="flex items-center gap-2">
              <h3 className="text-sm font-semibold text-slate-700">So sánh điểm trung bình của các đơn vị</h3>
              <div className="w-5 h-5 rounded-full bg-slate-100 flex items-center justify-center cursor-help" title="Điểm trung bình của tất cả lần thi trong kỳ">
                <span className="text-xs text-slate-400 font-bold">i</span>
              </div>
            </div>
            <div className="flex items-center gap-1 border border-slate-200 rounded-lg px-2.5 py-1 text-xs text-slate-600 bg-slate-50">
              Điểm TB <ChevronRight size={12} className="rotate-90 text-slate-400"/>
            </div>
          </div>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={barData} margin={{left:-15,right:10,top:20}}>
              <XAxis dataKey="name" tick={{fontSize:11}} axisLine={false} tickLine={false}/>
              <YAxis domain={[0,100]} tick={{fontSize:11}} axisLine={false} tickLine={false} tickFormatter={v=>`${v}%`}/>
              <Tooltip cursor={false} formatter={v=>[`${v}%`,'Điểm TB']}/>
              <Bar dataKey="avg" radius={[6,6,0,0]} label={{position:'top',fontSize:11,fontWeight:600}}>
                {barData.map((d,i)=><Cell key={i} fill={d.fill}/>)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
          {/* Dept legend */}
          <div className="flex flex-wrap gap-3 mt-2 justify-center">
            {barData.map((d,i)=>(
              <div key={i} className="flex items-center gap-1.5">
                <div className="w-2.5 h-2.5 rounded-sm" style={{background:d.fill}}/>
                <span className="text-xs text-slate-500">{d.name}</span>
              </div>
            ))}
          </div>
        </div>

        {/* Donut chart */}
        <div className="md:lg:col-span-2 bg-white rounded-2xl p-5 shadow-sm border border-slate-100">
          <h3 className="text-sm font-semibold text-slate-700 mb-3">Phân bố kết quả</h3>
          <div className="flex items-center gap-4">
            <div className="relative flex-shrink-0">
              <ResponsiveContainer width={180} height={180}>
                <PieChart>
                  <Pie data={pieData} cx="50%" cy="50%" innerRadius={58} outerRadius={82} dataKey="value" startAngle={90} endAngle={-270}>
                    {pieData.map((d,i)=><Cell key={i} fill={d.c}/>)}
                  </Pie>
                  <Tooltip formatter={(v,n)=>[v+' lượt',n]}/>
                </PieChart>
              </ResponsiveContainer>
              <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                <span className="text-xl font-bold text-slate-800">{totalAttempts}</span>
                <span className="text-xs text-slate-400">Tổng lượt thi</span>
              </div>
            </div>
            <div className="flex-1 space-y-2.5">
              {pieData.map(d=>(
                <div key={d.name}>
                  <div className="flex items-center gap-1.5 mb-0.5">
                    <div className="w-2.5 h-2.5 rounded-full flex-shrink-0" style={{background:d.c}}/>
                    <span className="text-xs text-slate-600">{d.name}</span>
                  </div>
                  <div className="flex items-center justify-between text-xs pl-4">
                    <span className="text-slate-400">{d.value} lượt</span>
                    <span className="font-semibold text-slate-700">{totalAttempts?Math.round(d.value/totalAttempts*100):0}%</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>

      {/* Detail table */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden mb-5">
        <div className="px-5 py-4 border-b border-slate-100">
          <h3 className="text-sm font-semibold text-slate-700">Chi tiết kết quả theo đơn vị</h3>
        </div>
        <table className="w-full">
          <thead>
            <tr className="border-b border-slate-100 bg-slate-50/50">
              {['Đơn vị','Lượt thi','Đạt','Chưa đạt','Điểm TB','Tiến độ đạt','Xu hướng'].map(h=>(
                <th key={h} className="px-5 py-3 text-left text-xs font-semibold text-slate-400 uppercase tracking-wide">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {deptData.map((d,i)=>{
              const isUp = i%2===0;
              const trendVal = trends[i]||'+0%';
              const trendUp = trendVal.startsWith('+');
              const barCol = d.rate>=80?'bg-emerald-500':d.rate>=60?'bg-amber-400':'bg-red-400';
              return (
                <tr key={d.name} className="border-t border-slate-50 hover:bg-slate-50/40 transition-colors">
                  <td className="px-5 py-4">
                    <div className="flex items-center gap-2">
                      <div className="w-7 h-7 rounded-lg flex items-center justify-center flex-shrink-0" style={{background:barColors[i%barColors.length]+'20'}}>
                        <BarChart2 size={13} style={{color:barColors[i%barColors.length]}}/>
                      </div>
                      <span className="text-sm font-medium text-slate-800">{d.name}</span>
                    </div>
                  </td>
                  <td className="px-5 py-4 text-sm text-slate-600">{d.attempts}</td>
                  <td className="px-5 py-4 text-sm font-semibold text-emerald-600">{d.passed}</td>
                  <td className="px-5 py-4 text-sm font-semibold text-red-500">{d.failed}</td>
                  <td className="px-5 py-4 text-sm font-bold text-slate-800">{d.avg}%</td>
                  <td className="px-5 py-4">
                    <div className="flex items-center gap-3">
                      <div className="flex-1 h-2 bg-slate-100 rounded-full overflow-hidden min-w-24">
                        <div className={`h-full rounded-full ${barCol}`} style={{width:`${d.rate}%`}}/>
                      </div>
                      <span className="text-xs font-semibold text-slate-600 w-10">{d.rate}%</span>
                    </div>
                  </td>
                  <td className="px-5 py-4">
                    <div className={`flex items-center gap-1 ${trendUp?'text-emerald-600':'text-red-500'}`}>
                      <span className="text-sm">{trendUp?'↗':'↘'}</span>
                      <div>
                        <div className="text-xs font-bold">{trendVal}</div>
                        <div className="text-xs text-slate-400">so với kỳ trước</div>
                      </div>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Summary box */}
      <div className="bg-blue-50 border border-blue-100 rounded-2xl p-5 flex items-center gap-4 overflow-hidden relative">
        <div className="w-9 h-9 bg-blue-100 rounded-xl flex items-center justify-center flex-shrink-0">
          <TrendingUp size={18} className="text-blue-600"/>
        </div>
        <div className="flex-1">
          <div className="font-semibold text-blue-700 mb-1 text-sm">Nhận xét tổng quan</div>
          <p className="text-sm text-slate-600">
            Đơn vị <strong>{bestDept.name}</strong> đang có kết quả tốt nhất với {bestDept.avg}% điểm trung bình.
            {worstDept&&worstDept.name!==bestDept.name&&<> <strong>{worstDept.name}</strong> cần cải thiện thêm để nâng cao chất lượng kết quả.</>}
          </p>
        </div>
        {/* Decorative illustration */}
        <div className="flex-shrink-0 opacity-20 select-none" aria-hidden>
          <svg width="100" height="70" viewBox="0 0 100 70">
            <rect x="5" y="40" width="12" height="25" rx="3" fill="#15803d"/>
            <rect x="22" y="28" width="12" height="37" rx="3" fill="#4ade80"/>
            <rect x="39" y="18" width="12" height="47" rx="3" fill="#15803d"/>
            <rect x="56" y="30" width="12" height="35" rx="3" fill="#4ade80"/>
            <circle cx="80" cy="28" r="18" fill="none" stroke="#15803d" strokeWidth="3"/>
            <path d="M72 28 L78 34 L88 22" stroke="#15803d" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" fill="none"/>
          </svg>
        </div>
      </div>
    </div>
  );
};


// ── EMPLOYEE HOME ──
const EmpHome = ({user, exams, results, onStart}) => {
  const my = results.filter(r=>r.empId===user.id);
  const passed = my.filter(r=>{const e=exams.find(x=>x.id===r.examId);return e&&r.score>=e.pass;}).length;
  const avg = my.length?Math.round(my.reduce((a,r)=>a+r.score,0)/my.length):0;
  return (
    <div>
      <div className="mb-4">
        <h1 className="text-lg md:text-xl font-bold text-slate-800">Xin chào, {user.name.split(' ').pop()}! 👋</h1>
        <p className="text-slate-500 text-xs md:text-sm">{user.dept} • {new Date().toLocaleDateString('vi-VN')}</p>
      </div>
      <div className="grid grid-cols-3 gap-2 md:gap-3 mb-5">
        {[{v:my.length,lb:'Lần đã thi',cl:'text-emerald-700'},{v:passed,lb:'Bài đạt',cl:'text-emerald-600'},{v:`${avg}%`,lb:'Điểm TB',cl:'text-emerald-600'}].map(s=>(
          <div key={s.lb} className="bg-white rounded-xl p-3 md:p-4 shadow-sm border border-slate-100 text-center">
            <div className={`text-xl md:text-2xl font-bold ${s.cl}`}>{s.v}</div>
            <div className="text-xs text-slate-500 mt-0.5">{s.lb}</div>
          </div>
        ))}
      </div>
      <h2 className="font-semibold text-slate-700 mb-3 text-sm">Danh sách bài thi</h2>
      <div className="space-y-3">
        {exams.map(exam=>{
          const last = my.filter(r=>r.examId===exam.id).slice(-1)[0];
          const ok = last&&last.score>=exam.pass;
          return (
            <div key={exam.id} className="bg-white rounded-xl p-4 shadow-sm border border-slate-100">
              <div className="flex items-start gap-3">
                <div className="w-10 h-10 bg-emerald-50 rounded-xl flex items-center justify-center flex-shrink-0"><FileText size={18} className="text-emerald-700"/></div>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 mb-1 flex-wrap">
                    <h3 className="font-semibold text-slate-800 text-sm">{exam.title}</h3>
                    {ok&&<span className="px-2 py-0.5 bg-emerald-100 text-emerald-700 rounded-full text-xs">✓ Đạt</span>}
                    {last&&!ok&&<span className="px-2 py-0.5 bg-red-100 text-red-700 rounded-full text-xs">✗ Chưa đạt</span>}
                  </div>
                  <p className="text-xs text-slate-500 mb-2 line-clamp-1">{exam.desc}</p>
                  <div className="flex flex-wrap gap-2 text-xs text-slate-400">
                    <span className="flex items-center gap-1"><FileText size={10}/>{exam.qIds.length} câu</span>
                    <span className="flex items-center gap-1"><Clock size={10}/>{exam.time} phút</span>
                    <span className="flex items-center gap-1"><Award size={10}/>Đạt: {exam.pass}%</span>
                    {last&&<span className="text-slate-500">Điểm: {last.score}%</span>}
                  </div>
                </div>
                <button onClick={()=>onStart(exam)} className="flex items-center gap-1 bg-emerald-600 text-white px-3 py-2 rounded-lg text-xs hover:bg-emerald-700 flex-shrink-0">
                  <Play size={11}/>{last?'Thi lại':'Bắt đầu'}
                </button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};

// ── MY RESULTS ──
const MyResults = ({user, exams, results}) => {
  const my = results.filter(r=>r.empId===user.id);
  return (
    <div>
      <div className="mb-5"><h1 className="text-xl font-bold text-slate-800">Kết quả của tôi</h1><p className="text-slate-500 text-sm">Lịch sử thi và điểm số</p></div>
      {my.length===0?(
        <div className="text-center py-20 text-slate-400"><Award size={40} className="mx-auto mb-3 opacity-20"/><p>Bạn chưa tham gia kỳ thi nào</p></div>
      ):(
        <div className="space-y-3">
          {my.map(r=>{const exam=exams.find(e=>e.id===r.examId);const ok=exam&&r.score>=exam.pass;return(
            <div key={r.id} className="bg-white rounded-xl p-5 shadow-sm border border-slate-100 flex items-center gap-4">
              <div className={`w-14 h-14 rounded-xl flex items-center justify-center flex-shrink-0 ${ok?'bg-emerald-100':'bg-red-100'}`}>
                <span className={`text-lg font-bold ${ok?'text-emerald-600':'text-red-600'}`}>{r.score}%</span>
              </div>
              <div className="flex-1">
                <h3 className="font-semibold text-slate-800 text-sm">{exam?.title}</h3>
                <div className="flex gap-3 mt-1 text-xs text-slate-500">
                  <span>{r.correct}/{exam?.qIds.length} câu đúng</span>
                  <span>Cần: {exam?.pass}%</span>
                  <span>{r.date}</span>
                </div>
              </div>
              <Badge ok={ok}/>
            </div>
          );})}
        </div>
      )}
    </div>
  );
};

// ── EXAM SCREEN ──
const ExamScreen = ({user, exam, questions, onFinish}) => {
  const qs = exam.qIds.map(id => questions.find(q => q.id === id)).filter(Boolean);
  const [cur, setCur] = useState(0);
  const [ans, setAns] = useState(Array(qs.length).fill(-1));
  const [tLeft, setTLeft] = useState(exam.time*60);
  const [confirm, setConfirm] = useState(false);
  const ansRef = useRef(ans);
  const doneRef = useRef(false);

  useEffect(()=>{ansRef.current=ans;},[ans]);

  const submit = (a) => {
    if(doneRef.current) return; doneRef.current=true;
    const correct = a.filter((x,i)=>x===qs[i].ans).length;
    const score = Math.round(correct/qs.length*100);
    const timeTaken = exam.time*60 - tLeft; // seconds used
    onFinish({id:Date.now(),empId:user.id,examId:exam.id,score,correct,timeTaken,date:new Date().toLocaleDateString('vi-VN'),answers:[...a]});
  };

  useEffect(()=>{
    const t = setInterval(()=>{
      setTLeft(v=>{
        if(v<=1){clearInterval(t);setTimeout(()=>submit(ansRef.current),0);return 0;}
        return v-1;
      });
    },1000);
    return()=>clearInterval(t);
  },[]);

  const mins = Math.floor(tLeft/60); const secs = tLeft%60;
  const low = tLeft<60; const done = ans.filter(x=>x!==-1).length;
  const q = qs[cur];

  return (
    <div className="min-h-screen bg-green-950 flex flex-col">
      <div className="bg-green-950 px-6 py-4 flex items-center justify-between border-b border-green-900">
        <div><div className="text-white font-semibold">{exam.title}</div><div className="text-slate-400 text-sm">{user.name}</div></div>
        <div className={`flex items-center gap-2 text-lg font-mono font-bold px-4 py-2 rounded-lg ${low?'bg-red-500 text-white animate-pulse':'bg-green-900 text-white'}`}>
          <Clock size={16}/>{String(mins).padStart(2,'0')}:{String(secs).padStart(2,'0')}
        </div>
      </div>
      <div className="px-6 py-2 bg-green-950/50">
        <div className="flex items-center gap-3">
          <div className="flex-1 h-1.5 bg-green-900 rounded-full overflow-hidden"><div className="h-full bg-emerald-600 rounded-full transition-all" style={{width:`${(cur+1)/qs.length*100}%`}}/></div>
          <span className="text-slate-400 text-xs">{cur+1}/{qs.length}</span>
        </div>
      </div>
      <div className="flex flex-1 gap-4 p-3 md:p-6">
        <div className="flex-1">
          <div className="bg-white rounded-2xl p-4 md:p-6 mb-4">
            <div className="flex items-center gap-2 mb-4">
              <span className="px-2.5 py-1 bg-emerald-50 text-emerald-700 rounded-lg text-xs font-medium">Câu {cur+1}</span>
            </div>
            <p className="text-slate-800 font-medium text-base leading-relaxed mb-6">{q.text}</p>
            <div className="space-y-3">
              {q.opts.map((o,i)=>(
                <button key={i} onClick={()=>{const a=[...ans];a[cur]=i;setAns(a);}} className={`w-full flex items-center gap-3 p-4 rounded-xl border-2 text-left transition-all ${ans[cur]===i?'border-emerald-600 bg-emerald-50':'border-slate-100 hover:border-slate-300 bg-slate-50/50'}`}>
                  <div className={`w-7 h-7 rounded-full flex items-center justify-center text-xs font-bold flex-shrink-0 transition-all ${ans[cur]===i?'bg-emerald-600 text-white':'bg-white border-2 border-slate-300 text-slate-500'}`}>{String.fromCharCode(65+i)}</div>
                  <span className={`text-sm ${ans[cur]===i?'text-emerald-700 font-medium':'text-slate-700'}`}>{o}</span>
                </button>
              ))}
            </div>
          </div>
          <div className="flex gap-3">
            <button onClick={()=>setCur(c=>Math.max(0,c-1))} disabled={cur===0} className="px-5 py-2.5 bg-slate-700 text-white rounded-xl text-sm disabled:opacity-40 hover:bg-slate-600">← Trước</button>
            {cur<qs.length-1
              ? <button onClick={()=>setCur(c=>c+1)} className="flex-1 py-2.5 bg-emerald-600 text-white rounded-xl text-sm font-medium hover:bg-emerald-700">Câu tiếp →</button>
              : <button onClick={()=>setConfirm(true)} className="flex-1 py-2.5 bg-emerald-500 text-white rounded-xl text-sm font-medium hover:bg-emerald-600">Nộp bài ✓</button>
            }
          </div>
        </div>
        <div className="w-48 flex-shrink-0">
          <div className="bg-white rounded-2xl p-4">
            <div className="text-sm font-semibold text-slate-700 mb-3">Bảng câu hỏi</div>
            <div className="grid grid-cols-5 sm:grid-cols-4 gap-1.5 mb-4">
              {qs.map((_,i)=>(
                <button key={i} onClick={()=>setCur(i)} className={`w-9 h-9 rounded-lg text-xs font-medium transition-all ${i===cur?'bg-[#0B4F32] text-white ring-2 ring-emerald-400':ans[i]!==-1?'bg-emerald-100 text-emerald-700 border border-emerald-200':'bg-slate-100 text-slate-500 hover:bg-slate-200'}`}>{i+1}</button>
              ))}
            </div>
            <div className="space-y-1.5 text-xs text-slate-400 mb-4">
              <div className="flex items-center gap-2"><div className="w-3 h-3 bg-[#0B4F32] ring-2 ring-emerald-400 rounded"/>Đang xem</div>
              <div className="flex items-center gap-2"><div className="w-3 h-3 bg-emerald-100 border border-emerald-300 rounded"/>Đã trả lời</div>
              <div className="flex items-center gap-2"><div className="w-3 h-3 bg-slate-100 rounded"/>Chưa trả lời</div>
            </div>
            <div className="pt-3 border-t border-slate-100 text-xs text-slate-500 mb-3">Đã trả lời: <span className="font-bold text-slate-700">{done}/{qs.length}</span></div>
            <button onClick={()=>setConfirm(true)} className="w-full py-2 bg-emerald-500 text-white rounded-xl text-xs font-medium hover:bg-emerald-600">Nộp bài</button>
          </div>
        </div>
      </div>
      {confirm&&(
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50">
          <div className="bg-white rounded-2xl p-6 w-80 shadow-2xl">
            <h3 className="font-bold text-slate-800 mb-2">Xác nhận nộp bài?</h3>
            <p className="text-sm text-slate-500 mb-4">Đã trả lời {done}/{qs.length} câu.{done<qs.length&&` Còn ${qs.length-done} câu chưa trả lời.`}</p>
            <div className="flex gap-2">
              <button onClick={()=>setConfirm(false)} className="flex-1 py-2 border border-slate-200 rounded-xl text-sm text-slate-600 hover:bg-slate-50">Tiếp tục làm</button>
              <button onClick={()=>submit(ans)} className="flex-1 py-2 bg-emerald-500 text-white rounded-xl text-sm font-medium hover:bg-emerald-600">Nộp bài</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// ── RESULT SCREEN ──
const ResultScreen = ({result, exam, questions, onBack}) => {
  const qs = exam.qIds.map(id => questions.find(q => q.id === id)).filter(Boolean);
  const ok = result.score>=exam.pass;
  return (
    <div className="min-h-screen bg-slate-50">
      <div className="max-w-2xl mx-auto p-4 md:p-6">
        <div className={`rounded-2xl p-8 text-center mb-6 text-white ${ok?'bg-emerald-500':'bg-red-500'}`}>
          <div className="text-5xl mb-3">{ok?'🎉':'😔'}</div>
          <div className="text-4xl font-bold mb-1">{result.score}%</div>
          <div className="text-base font-medium opacity-90 mb-2">{ok?'Chúc mừng! Bạn đã đạt yêu cầu':'Bạn chưa đạt yêu cầu lần này'}</div>
          <div className="opacity-75 text-sm">{result.correct}/{qs.length} câu đúng • Cần đạt {exam.pass}%</div>
        </div>
        <h2 className="font-semibold text-slate-700 mb-3 text-sm">Chi tiết đáp án</h2>
        <div className="space-y-3 mb-6">
          {qs.map((q,i)=>{
            const ua=result.answers[i]; const correct=ua===q.ans;
            return (
              <div key={q.id} className="bg-white rounded-xl p-4 shadow-sm border border-slate-100">
                <div className="flex items-start gap-2 mb-3">
                  <div className={`w-6 h-6 rounded-full flex items-center justify-center flex-shrink-0 mt-0.5 ${correct?'bg-emerald-100 text-emerald-600':'bg-red-100 text-red-600'}`}>
                    {correct?<CheckCircle size={14}/>:<XCircle size={14}/>}
                  </div>
                  <p className="text-sm font-medium text-slate-700">{i+1}. {q.text}</p>
                </div>
                <div className="pl-8 space-y-1">
                  {q.opts.map((o,j)=>(
                    <div key={j} className={`text-xs px-3 py-1.5 rounded-lg ${j===q.ans?'bg-emerald-50 text-emerald-700 font-medium':j===ua&&!correct?'bg-red-50 text-red-600':'text-slate-500'}`}>
                      {String.fromCharCode(65+j)}. {o}{j===q.ans?' ✓':''}{j===ua&&!correct?' ✗ (bạn chọn)':''}
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
        <button onClick={onBack} className="w-full py-3 bg-emerald-600 text-white rounded-xl font-medium hover:bg-emerald-700">← Quay lại trang chủ</button>
      </div>
    </div>
  );
};

// ── LOGIN ──
// Ngôi sao 5 cánh dạng polygon
const starPts = (cx, cy, r, ri) => Array.from({length:10}, (_,i)=>{
  const R = i%2 ? ri : r, a = -Math.PI/2 + i*Math.PI/5;
  return `${(cx+R*Math.cos(a)).toFixed(2)},${(cy+R*Math.sin(a)).toFixed(2)}`;
}).join(' ');

// Quốc huy: vòng nguyệt quế + khiên + sao (SVG thuần, không cần file ảnh)
const Emblem = ({size=92}) => {
  const branch = (
    <g>
      <path d="M45 91 C28 87 16 73 13 53" fill="none" stroke="#166534" strokeWidth="2.2" strokeLinecap="round"/>
      {Array.from({length:9}, (_,i)=>{
        const th = (100 + i*12) * Math.PI/180, s = 1 - i*0.045;
        const x = 50 + 39*Math.cos(th), y = 52 + 39*Math.sin(th);
        return <ellipse key={i} cx={x} cy={y} rx={7.2*s} ry={3.2*s} fill={i%2?'#15803d':'#166534'}
                        transform={`rotate(${th*180/Math.PI+90} ${x} ${y})`}/>;
      })}
    </g>
  );
  return (
    <svg width={size} height={size} viewBox="0 0 100 100" className="drop-shadow-md">
      <defs>
        <linearGradient id="em-shield" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor="#ef4444"/><stop offset="100%" stopColor="#991b1b"/>
        </linearGradient>
      </defs>
      {branch}
      <g transform="translate(100,0) scale(-1,1)">{branch}</g>
      <path d="M50 7 L78 17 V45 C78 61 66 74 50 82 C34 74 22 61 22 45 V17 Z"
            fill="url(#em-shield)" stroke="#facc15" strokeWidth="2.5" strokeLinejoin="round"/>
      <polygon points={starPts(50,43,19,8)} fill="#fde047" stroke="#f59e0b" strokeWidth="0.8"/>
    </svg>
  );
};

// Nền: sóng xanh + watermark trụ sở và ngôi sao
const LoginBackdrop = () => (
  <div className="absolute inset-0 overflow-hidden bg-[#f5faf6]" aria-hidden>
    <div className="absolute -top-48 -right-40 w-[40rem] h-[40rem] rounded-full bg-emerald-200/30 blur-3xl"/>
    <div className="absolute -bottom-32 -left-32 w-[34rem] h-[34rem] rounded-full bg-green-300/20 blur-3xl"/>

    <svg viewBox="0 0 300 240" className="absolute left-0 bottom-16 w-[26rem] max-w-[45vw] text-emerald-900/[0.07]" fill="currentColor">
      <rect x="10" y="206" width="280" height="14"/>
      <rect x="55" y="192" width="190" height="14"/>
      <rect x="48" y="92" width="204" height="100"/>
      <polygon points="38,92 150,48 262,92"/>
      {Array.from({length:6},(_,i)=><rect key={i} x={64+i*32} y="104" width="16" height="88" className="text-emerald-900/[0.12]"/>)}
      <rect x="147" y="6" width="3" height="44"/>
      <polygon points="150,10 190,20 150,30"/>
    </svg>

    <svg viewBox="0 0 100 100" className="absolute right-4 top-[22%] w-64 max-w-[30vw] text-emerald-800/[0.06]">
      <polygon points={starPts(50,50,44,19)} fill="currentColor"/>
      <circle cx="50" cy="50" r="47" fill="none" stroke="currentColor" strokeWidth="1.5"/>
    </svg>

    <svg className="absolute inset-x-0 bottom-0 w-full h-[38vh] min-h-[220px]" viewBox="0 0 1440 320" preserveAspectRatio="none">
      <path fill="#bbf7d0" fillOpacity=".5"  d="M0,190 C240,280 480,140 720,180 C960,220 1200,290 1440,196 L1440,320 L0,320 Z"/>
      <path fill="#22c55e" fillOpacity=".22" d="M0,246 C260,318 520,198 780,236 C1040,274 1240,318 1440,248 L1440,320 L0,320 Z"/>
      <path fill="#15803d" fillOpacity=".92" d="M0,292 C300,332 600,258 900,282 C1140,302 1290,318 1440,290 L1440,320 L0,320 Z"/>
    </svg>
  </div>
);

const StarDivider = () => (
  <div className="flex items-center justify-center gap-3 my-5">
    <span className="h-px w-16 sm:w-24 bg-gradient-to-r from-transparent to-emerald-400/70"/>
    <Star size={12} className="text-emerald-600 fill-emerald-600"/>
    <span className="h-px w-16 sm:w-24 bg-gradient-to-l from-transparent to-emerald-400/70"/>
  </div>
);

const BackBtn = ({onClick}) => (
  <button onClick={onClick} className="text-emerald-800/70 hover:text-emerald-900 text-xs mb-4 inline-flex items-center gap-1.5 transition-colors">
    <ArrowLeft size={13}/> Quay lại
  </button>
);

const Login = ({onLogin, employees}) => {
  const [step, setStep] = useState('role'); // 'role' | 'admin' | 'dept' | 'employee'
  const [adminForm, setAdminForm] = useState({user:'', pass:'', err:''});
  const [selectedDept, setSelectedDept] = useState('');

  const depts = [...new Set(employees.map(e=>e.dept))].sort();
  const deptEmployees = employees.filter(e=>e.dept===selectedDept);

  const handleAdminLogin = () => {
    if(adminForm.user==='admin' && adminForm.pass==='123') {
      onLogin({role:'admin', id:0, name:'Quản trị viên', dept:'Admin'});
    } else {
      setAdminForm(f=>({...f, err:'Tên đăng nhập hoặc mật khẩu không đúng'}));
    }
  };

  const back = (to) => { setStep(to); setAdminForm({user:'',pass:'',err:''}); setSelectedDept(''); };

  return (
    <div className="relative min-h-screen flex items-center justify-center p-4 sm:p-8">
      <LoginBackdrop/>

      <div className="relative w-full max-w-3xl rounded-[2rem] bg-white/75 backdrop-blur-xl border border-white/70 shadow-[0_30px_80px_-25px_rgba(6,78,59,0.35)] px-5 py-8 sm:px-14 sm:py-11">
        {/* HEADER */}
        <div className="text-center">
          <div className="flex justify-center"><Emblem/></div>
          <h1 className="mt-3 text-xl sm:text-3xl font-bold text-[#0B4F32] tracking-tight">Cục hậu cần - kỹ thuật Quân khu 4</h1>
          <p className="mt-1.5 text-slate-600 text-sm">Hệ thống thi trắc nghiệm nội bộ</p>
          <StarDivider/>
        </div>

        <div className="mx-auto w-full max-w-lg">
          {/* STEP 1: Chọn vai trò */}
          {step==='role' && (
            <div className="space-y-3">
              <p className="text-slate-500 text-xs text-center mb-5">Chọn vai trò của bạn để đăng nhập</p>
              {[
                {key:'admin', em:'👔', tt:'Quản trị viên', sub:'Quản lý câu hỏi, đề thi & báo cáo', to:'admin',
                 tile:'bg-emerald-100', ring:'hover:border-emerald-300', dot:'bg-emerald-50 text-emerald-700 group-hover:bg-emerald-600'},
                {key:'emp',   em:'👤', tt:'Thí sinh',     sub:'Tham gia thi và xem kết quả',      to:'dept',
                 tile:'bg-violet-100', ring:'hover:border-violet-300',  dot:'bg-violet-50 text-violet-700 group-hover:bg-violet-600'},
              ].map(item=>(
                <button key={item.key} onClick={()=>setStep(item.to)}
                        className={`group w-full bg-white border border-slate-200/80 ${item.ring} shadow-sm hover:shadow-lg hover:-translate-y-0.5 rounded-2xl p-4 sm:p-5 text-left transition-all`}>
                  <div className="flex items-center gap-4">
                    <div className={`w-12 h-12 sm:w-14 sm:h-14 ${item.tile} rounded-2xl flex items-center justify-center text-2xl flex-shrink-0`}>{item.em}</div>
                    <div className="h-10 w-px bg-slate-200 hidden sm:block"/>
                    <div className="flex-1 min-w-0">
                      <div className="font-bold text-[#0B4F32] text-base">{item.tt}</div>
                      <div className="text-slate-500 text-xs mt-0.5">{item.sub}</div>
                    </div>
                    <div className={`w-9 h-9 rounded-full ${item.dot} flex items-center justify-center group-hover:text-white transition-colors flex-shrink-0`}>
                      <ArrowRight size={16}/>
                    </div>
                  </div>
                </button>
              ))}
            </div>
          )}

          {/* STEP 2A: Đăng nhập Admin */}
          {step==='admin' && (
            <div>
              <BackBtn onClick={()=>back('role')}/>
              <div className="bg-white border border-slate-200/80 shadow-sm rounded-2xl p-5 sm:p-6 space-y-4">
                <div className="flex items-center gap-3 pb-1">
                  <div className="w-11 h-11 bg-emerald-100 rounded-2xl flex items-center justify-center text-xl">👔</div>
                  <div><div className="text-[#0B4F32] font-bold text-sm">Quản trị viên</div><div className="text-slate-500 text-xs">Nhập thông tin đăng nhập</div></div>
                </div>
                <div>
                  <label className="text-slate-600 text-xs mb-1.5 block font-medium">Tên đăng nhập</label>
                  <input
                    className="w-full bg-slate-50/70 border border-slate-200 rounded-xl px-3.5 py-2.5 text-slate-800 text-sm placeholder-slate-400 focus:outline-none focus:bg-white focus:border-emerald-500 focus:ring-2 focus:ring-emerald-500/20 transition"
                    placeholder="Nhập tên đăng nhập"
                    value={adminForm.user}
                    onChange={e=>setAdminForm(f=>({...f,user:e.target.value,err:''}))}
                    onKeyDown={e=>e.key==='Enter'&&handleAdminLogin()}
                  />
                </div>
                <div>
                  <label className="text-slate-600 text-xs mb-1.5 block font-medium">Mật khẩu</label>
                  <input
                    type="password"
                    className="w-full bg-slate-50/70 border border-slate-200 rounded-xl px-3.5 py-2.5 text-slate-800 text-sm placeholder-slate-400 focus:outline-none focus:bg-white focus:border-emerald-500 focus:ring-2 focus:ring-emerald-500/20 transition"
                    placeholder="Nhập mật khẩu"
                    value={adminForm.pass}
                    onChange={e=>setAdminForm(f=>({...f,pass:e.target.value,err:''}))}
                    onKeyDown={e=>e.key==='Enter'&&handleAdminLogin()}
                  />
                </div>
                {adminForm.err && (
                  <div className="flex items-center gap-2 bg-emerald-50 border border-emerald-200 rounded-xl px-3 py-2">
                    <AlertCircle size={14} className="text-emerald-500 flex-shrink-0"/>
                    <span className="text-emerald-600 text-xs">{adminForm.err}</span>
                  </div>
                )}
                <button onClick={handleAdminLogin} className="w-full bg-gradient-to-r from-[#0B4F32] to-emerald-600 hover:from-[#0a4429] hover:to-emerald-700 text-white py-2.5 rounded-xl text-sm font-semibold shadow-lg shadow-emerald-900/15 transition-all">
                  Đăng nhập
                </button>
              </div>
            </div>
          )}

          {/* STEP 2B: Chọn đơn vị */}
          {step==='dept' && (
            <div>
              <BackBtn onClick={()=>back('role')}/>
              <p className="text-slate-500 text-xs text-center mb-4">Chọn đơn vị</p>
              <div className="space-y-2 max-h-[45vh] overflow-y-auto pr-1">
                {depts.map(dept=>{
                  const count = employees.filter(e=>e.dept===dept).length;
                  return (
                    <button key={dept} onClick={()=>{setSelectedDept(dept);setStep('employee');}}
                            className="group w-full bg-white border border-slate-200/80 hover:border-emerald-300 shadow-sm hover:shadow-md rounded-2xl p-4 text-left transition-all flex items-center gap-3">
                      <div className="w-10 h-10 bg-emerald-100 rounded-xl flex items-center justify-center text-base flex-shrink-0">🏢</div>
                      <div className="flex-1 min-w-0"><div className="font-semibold text-sm text-[#0B4F32] truncate">{dept}</div><div className="text-slate-500 text-xs">{count} thí sinh</div></div>
                      <ChevronRight size={16} className="text-slate-400 group-hover:text-emerald-600 transition-colors flex-shrink-0"/>
                    </button>
                  );
                })}
              </div>
            </div>
          )}

          {/* STEP 3: Chọn thí sinh */}
          {step==='employee' && (
            <div>
              <BackBtn onClick={()=>setStep('dept')}/>
              <div className="flex items-center gap-2 mb-4 px-1">
                <span className="text-slate-500 text-xs">Đơn vị:</span>
                <span className="bg-emerald-50 text-emerald-700 border border-emerald-200 text-xs px-2.5 py-1 rounded-full font-medium">{selectedDept}</span>
              </div>
              <p className="text-slate-500 text-xs text-center mb-3">Chọn tên của bạn</p>
              <div className="space-y-2 max-h-[45vh] overflow-y-auto pr-1">
                {deptEmployees.map(emp=>(
                  <button key={emp.id} onClick={()=>onLogin({role:'employee',...emp})}
                          className="group w-full bg-white border border-slate-200/80 hover:border-emerald-300 shadow-sm hover:shadow-md rounded-2xl p-3.5 text-left transition-all flex items-center gap-3">
                    <Avatar name={emp.name} sz="md"/>
                    <div className="min-w-0"><div className="font-semibold text-sm text-[#0B4F32] truncate">{emp.name}</div><div className="text-slate-500 text-xs truncate">{emp.dept}</div></div>
                    <ChevronRight size={16} className="text-slate-400 group-hover:text-emerald-600 transition-colors ml-auto flex-shrink-0"/>
                  </button>
                ))}
                {deptEmployees.length===0 && <p className="text-center text-slate-400 text-sm py-4">Không có thí sinh trong đơn vị này</p>}
              </div>
            </div>
          )}
        </div>

        {/* FOOTER */}
        <div className="mt-9 flex items-center justify-center gap-2 text-emerald-900/60 text-xs">
          <ShieldCheck size={14}/> Bảo mật - Chính xác - Hiệu quả
        </div>
      </div>
    </div>
  );
};

// ── MAIN APP ──
// ── EXAM RESULTS (admin): pick an exam → list every attempt for it ──
const ExamResults = ({results, exams, employees, onClearAll}) => {
  const [selId, setSelId] = useState(null);
  const [confirmClear, setConfirmClear] = useState(false);
  const fmtTime = (t) => t!=null ? `${Math.floor(t/60)}p${String(t%60).padStart(2,'0')}s` : '--';

  // Confirmation dialog for wiping every result
  const clearModal = confirmClear && (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-2xl w-full max-w-sm shadow-2xl p-5">
        <div className="flex items-center gap-2.5 mb-2">
          <div className="w-9 h-9 bg-red-100 rounded-lg flex items-center justify-center flex-shrink-0"><AlertCircle size={18} className="text-red-600"/></div>
          <h2 className="font-bold text-slate-800">Xóa tất cả kết quả thi?</h2>
        </div>
        <p className="text-sm text-slate-600 mb-4">Toàn bộ <span className="font-semibold text-red-600">{results.length} lượt thi</span> của tất cả đề thi sẽ bị xóa vĩnh viễn và <span className="font-semibold">không thể khôi phục</span>. Câu hỏi, đề thi và thí sinh không bị ảnh hưởng.</p>
        <div className="flex gap-2">
          <button onClick={()=>setConfirmClear(false)} className="flex-1 py-2 border border-slate-200 rounded-lg text-sm text-slate-600 hover:bg-slate-50">Hủy</button>
          <button onClick={()=>{onClearAll&&onClearAll(); setConfirmClear(false); setSelId(null);}} className="flex-1 py-2 bg-red-600 text-white rounded-lg text-sm font-medium hover:bg-red-700">Xóa tất cả</button>
        </div>
      </div>
    </div>
  );

  // Step 1 — choose an exam
  if(selId==null){
    return (
      <div>
        {clearModal}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-4 gap-3">
          <div><h1 className="text-lg md:text-xl font-bold text-slate-800">Kết quả thi</h1><p className="text-slate-500 text-xs md:text-sm mt-0.5">Chọn một đề thi để xem danh sách người thi và kết quả</p></div>
          {results.length>0 && (
            <button onClick={()=>setConfirmClear(true)} className="flex items-center gap-2 border border-red-200 text-red-600 px-3 py-2 rounded-lg hover:bg-red-50 text-sm font-medium self-start"><Trash2 size={15}/>Xóa tất cả kết quả</button>
          )}
        </div>
        {exams.length===0 ? (
          <div className="bg-white rounded-xl p-8 text-center text-slate-400 border border-slate-100">Chưa có đề thi nào.</div>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
            {exams.map(exam=>{
              const count = results.filter(r=>r.examId===exam.id).length;
              return (
                <button key={exam.id} onClick={()=>setSelId(exam.id)} className="text-left bg-white rounded-xl p-5 shadow-sm border border-slate-100 hover:border-emerald-300 hover:shadow transition-all">
                  <h3 className="font-semibold text-slate-800 mb-1">{exam.title}</h3>
                  <p className="text-sm text-slate-500 mb-3 truncate">{exam.desc}</p>
                  <div className="flex items-center justify-between text-xs text-slate-500">
                    <span className="flex items-center gap-1"><Users size={12}/>{count} lượt thi</span>
                    <span className="flex items-center gap-1 text-emerald-600 font-medium">Xem <ChevronRight size={13}/></span>
                  </div>
                </button>
              );
            })}
          </div>
        )}
      </div>
    );
  }

  // Step 2 — results of the selected exam (all attempts, newest first)
  const exam = exams.find(e=>e.id===selId);
  const totalQ = exam?.qIds?.length || 0;
  const rows = results.filter(r=>r.examId===selId).sort((a,b)=>b.id-a.id);

  const exportExcel = () => {
    const head = ['Họ và tên','Đơn vị','Điểm (%)','Số câu đúng','Tổng câu','Kết quả','Thời gian làm bài','Ngày thi'];
    const body = rows.map(r=>{
      const emp = employees.find(e=>e.id===r.empId);
      const ok = exam && r.score >= exam.pass;
      return [emp?.name||'(Đã xóa)', emp?.dept||'', r.score, r.correct, totalQ, ok?'Đạt':'Không đạt', fmtTime(r.timeTaken), r.date];
    });
    const ws = XLSX.utils.aoa_to_sheet([
      ['KẾT QUẢ THI: '+(exam?.title||'')],
      ['Xuất ngày: '+new Date().toLocaleDateString('vi-VN')],
      [],
      head, ...body,
    ]);
    ws['!cols']=[{wch:22},{wch:18},{wch:10},{wch:12},{wch:10},{wch:14},{wch:16},{wch:14}];
    ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:7}}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Kết quả');
    const safeTitle = (exam?.title||'de_thi').replace(/[^\p{L}\p{N}]+/gu,'_').slice(0,40);
    XLSX.writeFile(wb, `KetQua_${safeTitle}_${new Date().toLocaleDateString('vi-VN').replace(/\//g,'-')}.xlsx`);
  };

  return (
    <div>
      <button onClick={()=>setSelId(null)} className="flex items-center gap-1 text-sm text-slate-500 hover:text-emerald-600 mb-3"><ChevronRight size={15} className="rotate-180"/>Chọn đề thi khác</button>
      <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-4 gap-3">
        <div>
          <h1 className="text-lg md:text-xl font-bold text-slate-800">{exam?.title||'Đề thi đã xóa'}</h1>
          <p className="text-slate-500 text-xs md:text-sm mt-0.5">{rows.length} lượt thi • {totalQ} câu • Điểm đạt: {exam?.pass}%</p>
        </div>
        {rows.length>0 && (
          <button onClick={exportExcel} className="flex items-center gap-2 bg-emerald-600 text-white px-4 py-2 rounded-lg hover:bg-emerald-700 text-sm font-medium self-start"><Download size={15}/>Xuất Excel</button>
        )}
      </div>

      {rows.length===0 ? (
        <div className="bg-white rounded-xl p-8 text-center text-slate-400 border border-slate-100">Chưa có ai thi đề này.</div>
      ) : (
        <div className="bg-white rounded-xl shadow-sm border border-slate-100 overflow-x-auto">
          <table className="w-full min-w-[640px]">
            <thead><tr className="bg-slate-50">{['Thí sinh','Đơn vị','Điểm','Số câu đúng','Kết quả','Thời gian','Ngày thi'].map(h=><th key={h} className={`px-3 py-2 text-xs font-medium text-slate-400 uppercase ${h==='Kết quả'?'text-center':'text-left'}`}>{h}</th>)}</tr></thead>
            <tbody>
              {rows.map(r=>{
                const emp=employees.find(e=>e.id===r.empId);
                const ok=exam&&r.score>=exam.pass;
                return (
                  <tr key={r.id} className="border-t border-slate-50 hover:bg-slate-50/50">
                    <td className="px-3 py-2 text-xs text-slate-700 font-medium">{emp?.name||<span className="text-slate-400 italic">(Đã xóa)</span>}</td>
                    <td className="px-3 py-2 text-xs text-slate-600">{emp?.dept||'--'}</td>
                    <td className="px-3 py-2 text-xs font-bold text-slate-800">{r.score}%</td>
                    <td className="px-3 py-2 text-xs text-slate-600">{r.correct}/{totalQ}</td>
                    <td className="px-3 py-2 text-center"><Badge ok={ok}/></td>
                    <td className="px-3 py-2 text-xs text-slate-500">{fmtTime(r.timeTaken)}</td>
                    <td className="px-3 py-2 text-xs text-slate-400">{r.date}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

export default function App() {
  const [user, setUser] = useState(null);
  const [view, setView] = useState('dashboard');
  const [employees, setEmployees] = useState([]);
  const [questions, setQuestions] = useState([]);
  const [exams, setExams] = useState([]);
  const [results, setResults] = useState([]);
  const [activeExam, setActiveExam] = useState(null);
  const [lastResult, setLastResult] = useState(null);
  const [loading, setLoading] = useState(true);

  // ── Realtime sync từ Firebase ──
  useEffect(() => {
    const unsubs = [];
    let loadCount = 0;
    const TOTAL = 4;
    const sub = (colName, setter) => {
      let first = true;
      const unsub = onSnapshot(collection(db, colName), snap => {
        setter(snap.docs.map(d => ({ ...d.data(), id: d.data().id || d.id })));
        if (first) { first = false; loadCount++; if (loadCount >= TOTAL) setLoading(false); }
      });
      unsubs.push(unsub);
    };
    sub(COL.employees, setEmployees);
    sub(COL.questions, setQuestions);
    sub(COL.exams, setExams);
    sub(COL.results, setResults);
    return () => unsubs.forEach(u => u());
  }, []);

  // ── Firebase write helpers ──
  const fbSet = (colName, item) => setDoc(doc(db, colName, String(item.id)), item);
  const fbDel = (colName, id)  => deleteDoc(doc(db, colName, String(id)));

  const makeSyncSetter = (colName, localSetter) => (updater) => {
    localSetter(prev => {
      const next = typeof updater === 'function' ? updater(prev) : updater;
      next.forEach(item => {
        const old = prev.find(e => e.id === item.id);
        if (!old || JSON.stringify(old) !== JSON.stringify(item)) fbSet(colName, item);
      });
      prev.forEach(item => { if (!next.find(e => e.id === item.id)) fbDel(colName, item.id); });
      return next;
    });
  };

  const setEmployeesSync = makeSyncSetter(COL.employees, setEmployees);
  const setQuestionsSync = makeSyncSetter(COL.questions, setQuestions);
  const setExamsSync     = makeSyncSetter(COL.exams, setExams);
  const setResultsSync   = makeSyncSetter(COL.results, setResults);

  const login     = u => { setUser(u); setView(u.role==="admin"?"dashboard":"home"); };
  const logout    = () => { setUser(null); setActiveExam(null); setLastResult(null); };
  const startExam = exam => { setActiveExam(exam); setLastResult(null); };
  const finishExam = async r => {
    await fbSet(COL.results, r);
    setLastResult(r);
    setActiveExam(null);
  };

  if (loading) return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center">
      <div className="text-center">
        <div className="flex justify-center mb-3 animate-pulse"><Emblem size={64}/></div>
        <p className="text-slate-500 text-sm">Đang tải dữ liệu...</p>
      </div>
    </div>
  );

  if (!user) return <Login onLogin={login} employees={employees}/>;
  if (activeExam) return <ExamScreen user={user} exam={activeExam} questions={questions} onFinish={finishExam}/>;
  if (lastResult) {
    const exam = exams.find(e=>e.id===lastResult.examId);
    return <ResultScreen result={lastResult} exam={exam} questions={questions} onBack={()=>{setLastResult(null);setView("home");}}/>;
  }

  const adminViews = {
    dashboard: <Reports results={results} exams={exams} employees={employees}/>,
    questions: <Questions questions={questions} setQuestions={setQuestionsSync}/>,
    exams:     <Exams exams={exams} setExams={setExamsSync} questions={questions}/>,
    results:   <ExamResults results={results} exams={exams} employees={employees} onClearAll={()=>setResultsSync([])}/>,
    employees: <EmployeesView employees={employees} setEmployees={setEmployeesSync} results={results} exams={exams}/>,
  };
  const empViews = {
    home:    <EmpHome user={user} exams={exams} results={results} onStart={startExam}/>,
    results: <MyResults user={user} exams={exams} results={results}/>,
  };

  return (
    <div className="flex bg-slate-50 min-h-screen">
      <Sidebar role={user.role} active={view} setActive={setView} user={user} onLogout={logout}/>
      <div className="flex-1 md:ml-64 overflow-auto pt-14 md:pt-0 pb-16 md:pb-0">
        <div className="p-3 sm:p-4 md:p-7">{user.role==="admin"?adminViews[view]:empViews[view]}</div>
      </div>
    </div>
  );
}
