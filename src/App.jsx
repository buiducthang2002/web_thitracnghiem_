import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import legacyDocToText from "legacy-doc-reader";
import JSZip from "jszip";
import confetti from "canvas-confetti";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from "recharts";
import { BookOpen, Users, FileText, BarChart2, LogOut, Plus, Trash2, Clock, CheckCircle, Award, Home, Play, TrendingUp, TrendingDown, X, ChevronRight, Shield, ShieldCheck, Star, ArrowRight, ArrowLeft, Upload, Download, AlertCircle, Info, FileSearch, PieChart as PieChartIcon, Lock, LockOpen, Eye, EyeOff, User } from "lucide-react";
import { db, auth, missingConfig, projectId, toEmail, genPassword, createLoginAccount } from "./firebase";
import { collection, doc, setDoc, getDoc, deleteDoc, onSnapshot, query, where } from "firebase/firestore";
import { onAuthStateChanged, signInWithEmailAndPassword, signOut, setPersistence, browserLocalPersistence, browserSessionPersistence, EmailAuthProvider, reauthenticateWithCredential, updatePassword } from "firebase/auth";

const COL = { employees:'employees', questions:'questions', exams:'exams', results:'results', users:'users' };

// Thông báo lỗi đăng nhập thân thiện từ mã lỗi Firebase Auth
const authErrorMessage = err => {
  switch (err && err.code) {
    case 'auth/invalid-credential':
    case 'auth/user-not-found':
    case 'auth/wrong-password':
    case 'auth/invalid-email':    return 'Tài khoản hoặc mật khẩu không đúng.';
    case 'auth/too-many-requests': return 'Đăng nhập sai quá nhiều lần. Vui lòng thử lại sau ít phút.';
    case 'auth/network-request-failed': return 'Không kết nối được máy chủ. Kiểm tra mạng rồi thử lại.';
    case 'auth/operation-not-allowed': return 'Chưa bật đăng nhập Email/Mật khẩu trong Firebase Authentication.';
    default: return `Đăng nhập thất bại (${(err && (err.code || err.message)) || 'lỗi không xác định'}).`;
  }
};

// Nhãn phương án trả lời. Firestore rules cho phép 2–10 phương án mỗi câu.
const OPT_LETTERS = 'ABCDEFGHIJ';
const MAX_LEAVES = 3; // rời màn hình thi đến lần này thì tự động nộp bài
const MAX_OPTS = OPT_LETTERS.length;
// Ký tự đánh dấu chỗ in đậm/gạch chân khi đọc file Word (không xuất hiện trong văn bản thật).
const BOLD = '\u0001';

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

// Trạng thái rỗng dùng chung
const EmptyState = ({icon, title, sub}) => (
  <div className="flex flex-col items-center justify-center text-center py-6">
    <div className="w-12 h-12 rounded-full bg-slate-50 border border-slate-100 flex items-center justify-center text-slate-300 mb-2.5">{icon}</div>
    <div className="text-sm font-medium text-slate-500">{title}</div>
    {sub && <div className="text-xs text-slate-400 mt-0.5">{sub}</div>}
  </div>
);

// Khung trục giả cho biểu đồ khi chưa có dữ liệu
const ChartPlaceholder = ({cols, title, sub, icon}) => (
  <div className="relative h-[240px]">
    <div className="absolute inset-0 flex flex-col justify-between pb-7">
      {[100,75,50,25,0].map(v=>(
        <div key={v} className="flex items-center gap-2">
          <span className="w-9 text-right text-[11px] text-slate-400 flex-shrink-0">{v}%</span>
          <div className="flex-1 border-t border-dashed border-slate-100"/>
        </div>
      ))}
    </div>
    <div className="absolute bottom-0 left-11 right-0 flex justify-around">
      {cols.map(c=><span key={c} className="text-[11px] text-slate-400">{c}</span>)}
    </div>
    <div className="absolute inset-0 flex items-center justify-center"><EmptyState icon={icon} title={title} sub={sub}/></div>
  </div>
);

// Stable question order: by explicit `order`; legacy questions without one keep id order and stay first
const orderedQuestions = (qs) => [...qs].sort((a,b)=>{
  if(a.order!=null && b.order!=null) return a.order-b.order;
  if(a.order==null && b.order==null) return (a.id||0)-(b.id||0);
  return a.order==null ? -1 : 1;
});


// ── SIDEBAR (desktop) + MOBILE NAV ──
// ── Modal hiển thị tài khoản/mật khẩu vừa cấp (mật khẩu chỉ hiện ở đây, không lưu ở đâu) ──
const CredentialsModal = ({data, onClose}) => {
  const {creds, fails} = data;
  const download = () => {
    const ws = XLSX.utils.aoa_to_sheet([
      ['Họ và tên', 'Đơn vị', 'Tên đăng nhập', 'Mật khẩu'],
      ...creds.map(c => [c.name, c.dept, c.account, c.password]),
    ]);
    ws['!cols'] = [{wch:28},{wch:22},{wch:18},{wch:14}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Tai khoan');
    XLSX.writeFile(wb, 'tai-khoan-dang-nhap.xlsx');
  };
  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-[60] p-4">
      <div className="bg-white rounded-2xl w-full max-w-2xl shadow-2xl max-h-[90vh] flex flex-col">
        <div className="flex items-center justify-between p-5 border-b">
          <h2 className="font-bold text-slate-800">Tài khoản đăng nhập vừa cấp</h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18}/></button>
        </div>
        <div className="p-5 space-y-3 overflow-auto">
          {creds.length > 0 && (
            <>
              <div className="flex items-start gap-2 bg-amber-50 border border-amber-200 rounded-xl px-3 py-2">
                <AlertCircle size={14} className="text-amber-600 flex-shrink-0 mt-0.5"/>
                <span className="text-amber-800 text-xs">Mật khẩu chỉ hiển thị một lần này và không được lưu ở đâu. Hãy tải file Excel để gửi cho thí sinh trước khi đóng.</span>
              </div>
              <div className="border border-slate-100 rounded-xl overflow-auto max-h-72">
                <table className="w-full min-w-[480px]">
                  <thead><tr className="bg-slate-50 border-b border-slate-100">
                    {['Họ và tên','Đơn vị','Tên đăng nhập','Mật khẩu'].map(h=><th key={h} className="px-4 py-2 text-xs font-semibold text-slate-400 text-left">{h}</th>)}
                  </tr></thead>
                  <tbody>{creds.map((c,i)=>(
                    <tr key={i} className="border-t border-slate-50">
                      <td className="px-4 py-2 text-sm text-slate-800">{c.name}</td>
                      <td className="px-4 py-2 text-xs text-slate-500">{c.dept}</td>
                      <td className="px-4 py-2 text-sm font-mono text-slate-600">{c.account}</td>
                      <td className="px-4 py-2 text-sm font-mono font-semibold text-slate-800">{c.password}</td>
                    </tr>
                  ))}</tbody>
                </table>
              </div>
            </>
          )}
          {fails.length > 0 && (
            <div className="bg-red-50 border border-red-200 rounded-xl px-3 py-2">
              <p className="text-red-700 text-xs font-semibold mb-1">{fails.length} tài khoản không cấp được:</p>
              <ul className="text-red-600 text-xs space-y-0.5 max-h-32 overflow-auto">
                {fails.map((f,i)=><li key={i}>{f.emp.name}{f.emp.account ? ` (${f.emp.account})` : ''} — {f.reason}</li>)}
              </ul>
            </div>
          )}
          <div className="flex gap-2 pt-1">
            {creds.length > 0 && <button onClick={download} className="flex-1 py-2.5 bg-emerald-600 text-white rounded-xl text-sm font-medium hover:bg-emerald-700 flex items-center justify-center gap-2"><Download size={15}/>Tải file Excel</button>}
            <button onClick={onClose} className="flex-1 py-2.5 border border-slate-200 rounded-xl text-sm text-slate-600 hover:bg-slate-50">Đóng</button>
          </div>
        </div>
      </div>
    </div>
  );
};

// ── Modal đổi mật khẩu của chính người đang đăng nhập ──
const ChangePasswordModal = ({onClose}) => {
  const [oldPw, setOldPw] = useState('');
  const [newPw, setNewPw] = useState('');
  const [newPw2, setNewPw2] = useState('');
  const [err, setErr] = useState('');
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState(false);

  const submit = async e => {
    e.preventDefault();
    if (!oldPw) return setErr('Vui lòng nhập mật khẩu hiện tại.');
    if (newPw.length < 6) return setErr('Mật khẩu mới phải có ít nhất 6 ký tự.');
    if (newPw !== newPw2) return setErr('Hai lần nhập mật khẩu mới không khớp.');
    if (newPw === oldPw) return setErr('Mật khẩu mới phải khác mật khẩu hiện tại.');
    setBusy(true);
    try {
      const u = auth.currentUser;
      await reauthenticateWithCredential(u, EmailAuthProvider.credential(u.email, oldPw));
      await updatePassword(u, newPw);
      setDone(true);
    } catch (ex) {
      setErr(ex.code === 'auth/invalid-credential' || ex.code === 'auth/wrong-password'
        ? 'Mật khẩu hiện tại không đúng.'
        : `Không đổi được mật khẩu (${ex.code || ex.message}).`);
    }
    setBusy(false);
  };

  const inputCls = 'w-full border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm focus:outline-none focus:border-emerald-500';
  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-[60] p-4">
      <div className="bg-white rounded-2xl w-full max-w-sm shadow-2xl">
        <div className="flex items-center justify-between p-5 border-b">
          <h2 className="font-bold text-slate-800">Đổi mật khẩu</h2>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18}/></button>
        </div>
        {done ? (
          <div className="p-5 space-y-4">
            <p className="text-sm text-emerald-700">Đã đổi mật khẩu thành công.</p>
            <button onClick={onClose} className="w-full py-2.5 bg-emerald-600 text-white rounded-xl text-sm font-medium hover:bg-emerald-700">Đóng</button>
          </div>
        ) : (
          <form onSubmit={submit} className="p-5 space-y-3" autoComplete="off">
            <input type="password" className={inputCls} placeholder="Mật khẩu hiện tại" value={oldPw} onChange={e=>{setOldPw(e.target.value);setErr('');}} autoFocus/>
            <input type="password" className={inputCls} placeholder="Mật khẩu mới (tối thiểu 6 ký tự)" value={newPw} onChange={e=>{setNewPw(e.target.value);setErr('');}}/>
            <input type="password" className={inputCls} placeholder="Nhập lại mật khẩu mới" value={newPw2} onChange={e=>{setNewPw2(e.target.value);setErr('');}}/>
            {err && <p role="alert" className="text-red-600 text-xs">{err}</p>}
            <div className="flex gap-2 pt-1">
              <button type="button" onClick={onClose} className="flex-1 py-2.5 border border-slate-200 rounded-xl text-sm text-slate-600 hover:bg-slate-50">Hủy</button>
              <button type="submit" disabled={busy} className="flex-1 py-2.5 bg-emerald-600 text-white rounded-xl text-sm font-medium hover:bg-emerald-700 disabled:opacity-60">{busy?'Đang đổi...':'Đổi mật khẩu'}</button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
};

const Sidebar = ({role, active, setActive, user, onLogout, onChangePw, rail, setRail}) => {
  const nav = role==='admin'
    ? [{id:'dashboard',ic:<Home size={18}/>,lb:'Tổng quan'},{id:'questions',ic:<BookOpen size={18}/>,lb:'Câu hỏi'},{id:'exams',ic:<FileText size={18}/>,lb:'Đề thi'},{id:'results',ic:<Award size={18}/>,lb:'Kết quả'},{id:'employees',ic:<Users size={18}/>,lb:'Thí sinh'},{id:'accounts',ic:<ShieldCheck size={18}/>,lb:'Tài khoản'}]
    : [{id:'home',ic:<Home size={18}/>,lb:'Trang chủ'},{id:'results',ic:<Award size={18}/>,lb:'Kết quả'}];
  return (
    <>
      {/* Desktop sidebar */}
      <div className={`hidden md:flex ${rail?'w-[78px]':'w-[168px]'} bg-white h-screen flex-col fixed left-0 top-0 z-10 border-r border-slate-200/80 transition-[width] duration-200`}>
        <div className="py-5 flex flex-col items-center gap-1.5 flex-shrink-0">
          <Emblem size={rail?52:76}/>
          {!rail && <div className="text-[11px] font-extrabold tracking-wide text-[#0B4F32] text-center leading-tight px-2">BỆNH VIỆN QUÂN Y 4</div>}
        </div>

        <nav className="flex-1 px-2.5 space-y-1.5 relative">
          {nav.map(i=>(
            <button key={i.id} onClick={()=>setActive(i.id)} title={i.lb}
              className={`w-full flex items-center ${rail?'justify-center':'gap-2.5 px-3'} py-2.5 rounded-xl text-[13px] font-medium transition-all ${active===i.id?'bg-[#0B4F32] text-white shadow-md shadow-emerald-900/20':'text-slate-500 hover:bg-emerald-50 hover:text-[#0B4F32]'}`}>
              {i.ic}{!rail && <span className="truncate">{i.lb}</span>}
            </button>
          ))}
        </nav>

        <div className="p-2.5 border-t border-slate-100 flex-shrink-0 relative">
          <div className={`flex items-center ${rail?'justify-center':'gap-2'} py-1`}>
            <Avatar name={user.name} sz="sm"/>
            {!rail && (
              <div className="flex-1 min-w-0">
                <div className="text-slate-700 text-[11px] font-semibold truncate">{user.name}</div>
                <div className="text-slate-400 text-[10px] truncate">{role==='admin'?'Quản trị viên':user.dept}</div>
              </div>
            )}
          </div>
          <div className={`flex items-center ${rail?'flex-col gap-1.5':'justify-between'} mt-1.5`}>
            <button onClick={onLogout} title="Đăng xuất" className="p-1.5 rounded-lg text-slate-400 hover:text-red-500 hover:bg-red-50 transition-colors"><LogOut size={15}/></button>
            <button onClick={onChangePw} title="Đổi mật khẩu" className="p-1.5 rounded-lg text-slate-400 hover:text-[#0B4F32] hover:bg-emerald-50 transition-colors"><Lock size={15}/></button>
            <button onClick={()=>setRail(r=>!r)} title={rail?'Mở rộng':'Thu gọn'} className="p-1.5 rounded-lg text-slate-400 hover:text-[#0B4F32] hover:bg-emerald-50 transition-colors">
              <ChevronRight size={16} className={rail?'':'rotate-180'}/>
            </button>
          </div>
        </div>
      </div>

      {/* Mobile top bar */}
      <div className="md:hidden fixed top-0 left-0 right-0 z-20 bg-white flex items-center justify-between px-3 py-2 border-b border-slate-200">
        <div className="flex items-center gap-2 flex-shrink-0">
          <Emblem size={34}/>
          <div>
            <div className="text-[#0B4F32] font-bold text-[10px] leading-tight whitespace-nowrap">BỆNH VIỆN QUÂN Y 4</div>
            <div className="text-slate-400 text-[9px] whitespace-nowrap">Hệ thống thi trắc nghiệm</div>
          </div>
        </div>
        <div className="flex items-center gap-1.5 min-w-0">
          <span className="text-slate-600 text-xs truncate">{user.name.split(' ').pop()}</span>
          <button onClick={onChangePw} title="Đổi mật khẩu" className="text-slate-400 hover:text-[#0B4F32] p-1"><Lock size={15}/></button>
          <button onClick={onLogout} className="text-slate-400 hover:text-red-500 p-1"><LogOut size={15}/></button>
        </div>
      </div>

      {/* Mobile bottom nav */}
      <div className="md:hidden fixed bottom-0 left-0 right-0 z-20 bg-white border-t border-slate-200 flex">
        {nav.map(i=>(
          <button key={i.id} onClick={()=>setActive(i.id)} className={`flex-1 flex flex-col items-center gap-0.5 py-2 transition-all ${active===i.id?'text-[#0B4F32]':'text-slate-400'}`}>
            <div className={`px-3 py-1 rounded-lg ${active===i.id?'bg-emerald-100':''}`}>{i.ic}</div>
            <span className="text-[10px] font-medium">{i.lb}</span>
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
              <PieChart width={110} height={110}>
                <Pie data={pie} cx="50%" cy="50%" innerRadius={32} outerRadius={50} dataKey="value">{pie.map((d,i)=><Cell key={i} fill={d.c}/>)}</Pie><Tooltip/>
              </PieChart>
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
  const [editingId, setEditingId] = useState(null);
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

  const resetQuestionForm = () => {
    setModal(false);
    setEditingId(null);
    setCustomTopic('');
    setNq({topic:'Nội quy',level:'Dễ',text:'',opts:['','','',''],ans:0});
  };

  const openAdd = () => {
    setEditingId(null);
    setCustomTopic('');
    setNq({topic:'Nội quy',level:'Dễ',text:'',opts:['','','',''],ans:0});
    setModal(true);
  };

  const openEdit = q => {
    setEditingId(q.id);
    setCustomTopic('');
    setNq({
      topic: q.topic || 'Nội quy',
      level: q.level || 'Dễ',
      text: q.text || '',
      opts: [...(q.opts || ['', ''])],
      ans: Number.isInteger(q.ans) ? q.ans : 0,
    });
    setModal(true);
  };

  const saveQuestion = () => {
    if(!nq.text.trim()||nq.opts.some(o=>!o.trim())) return;
    if(!effectiveTopic) return;
    const updated = {...nq, text:nq.text.trim(), opts:nq.opts.map(o=>o.trim()), topic:effectiveTopic};
    if (editingId !== null) {
      setQuestions(p=>p.map(q=>q.id===editingId
        ? {...q, ...updated, needsReview:false, _issues:undefined}
        : q));
    } else {
      setQuestions(p=>[...p,{...updated, id:Date.now(), order: p.reduce((m,q)=>Math.max(m, q.order||0), 0) + 1}]);
    }
    resetQuestionForm();
  };

  // ── Đọc câu hỏi từ file Word ────────────────────────────────────────────
  // Nguyên tắc: tìm thấy bao nhiêu mốc "Câu N" trong file thì phải nhập ra đúng
  // bấy nhiêu câu, không bỏ sót câu nào. Câu nào thiếu thông tin (không có dòng
  // "Đáp án", thiếu phương án...) vẫn được nhập và gắn cờ needsReview để sửa tay,
  // thay vì bị loại âm thầm như trước.

  // mammoth.extractRawText làm mất định dạng, mà rất nhiều file lại đánh dấu đáp án
  // đúng bằng cách bôi đậm / gạch chân. Đọc thêm bản HTML rồi giữ lại dấu đó.
  const markStyledRuns = html => html.replace(/<span\b([^>]*)>([\s\S]*?)<\/span>/gi, (all, attrs, content) => {
    const marked = /font-weight\s*:\s*(?:bold|[6-9]00)/i.test(attrs) ||
      /text-decoration[^;]*(?:underline)/i.test(attrs) ||
      /color\s*:\s*(?:#(?:f00|ff0000)\b|rgb\(\s*255\s*,\s*0\s*,\s*0\s*\))/i.test(attrs);
    return marked ? `${BOLD}${content}${BOLD}` : content;
  });

  // Phần mở rộng file rất hay bị đặt sai (đổi tên .doc thành .docx, file do phần mềm
  // khác xuất ra...) nên cách đọc phải căn cứ vào chữ ký ở đầu file.
  const decodeBytes = (bytes) => {
    const head = String.fromCharCode.apply(null, bytes.slice(0, 2048));
    const declared = /charset\s*=\s*["']?\s*([\w-]+)/i.exec(head);
    for (const enc of [declared && declared[1], 'utf-8']) {
      if (!enc) continue;
      try { return new TextDecoder(enc).decode(bytes); } catch { /* bảng mã lạ thì thử cách sau */ }
    }
    return new TextDecoder().decode(bytes);
  };

  const sniffFileKind = (buffer) => {
    const b = new Uint8Array(buffer);
    if (b[0] === 0x50 && b[1] === 0x4b) return 'zip';                                   // .docx (thực chất là file nén)
    if (b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0) return 'ole';  // .doc Word 97–2003
    const head = decodeBytes(b.slice(0, 4096)).replace(/^﻿/, '').trimStart();
    if (/^\{\\rtf/i.test(head)) return 'rtf';
    if (/^(?:<!doctype\s+html|<html|<body|mime-version:)/i.test(head)) return 'html';
    return 'text';
  };

  const unescapeXml = s => s
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(+d))
    .replace(/&amp;/g, '&');

  // Bóc chữ thẳng từ XML bên trong file nén. Dùng khi mammoth từ chối đọc file
  // (báo "Could not find the body element") vì file do phần mềm khác sinh ra:
  // thẻ w:body bị đổi tên, thiếu khai báo namespace... Ở đây chỉ dựa vào tên thẻ
  // nên đọc được cả những file không đúng chuẩn.
  const wordXmlToMarkedText = (xml) => {
    // Mỗi đoạn chữ (w:r) in đậm / gạch chân được đánh dấu lại, vì rất nhiều file
    // đánh dấu đáp án đúng bằng cách này.
    const marked = xml.replace(/<(?:\w+:)?r(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?r>/g, (all, run) => {
      const pr = /<(?:\w+:)?rPr(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?rPr>/.exec(run);
      const props = pr ? pr[1] : '';
      const on = (tag) => {
        const m = new RegExp(`<(?:\\w+:)?${tag}(?:\\s[^>]*)?/?>`).exec(props);
        return !!m && !/val="(?:0|false|off|none)"/i.test(m[0]);
      };
      const body = run
        .replace(/<(?:\w+:)?rPr(?:\s[^>]*)?>[\s\S]*?<\/(?:\w+:)?rPr>/g, '')
        .replace(/<(?:\w+:)?(?:instrText|delText)(?:\s[^>]*)?>[\s\S]*?<\/(?:\w+:)?(?:instrText|delText)>/g, '')
        .replace(/<(?:\w+:)?tab(?:\s[^>]*)?\/?>/g, ' ')
        .replace(/<(?:\w+:)?br(?:\s[^>]*)?\/?>/g, '\n')
        .replace(/<[^>]+>/g, '');
      return body && (on('b') || on('u')) ? `${BOLD}${body}${BOLD}` : body;
    });
    return unescapeXml(marked
      .replace(/<\/(?:\w+:)?(?:p|tc|tr|h[1-6])>/gi, '\n')
      .replace(/<(?:\w+:)?br(?:\s[^>]*)?\/?>/gi, '\n')
      .replace(/<[^>]+>/g, ''))
      .replace(/[ \t]+\n/g, '\n');
  };

  const zipDocumentToMarkedText = async (buffer) => {
    const zip = await JSZip.loadAsync(buffer);
    // word/document.xml (.docx) trước, content.xml (.odt của LibreOffice) sau.
    const parts = Object.keys(zip.files)
      .filter(n => /(^|\/)(document\d*\.xml|content\.xml)$/i.test(n) && !/_rels/i.test(n))
      .sort((a, b) => (/document\.xml$/i.test(b) ? 1 : 0) - (/document\.xml$/i.test(a) ? 1 : 0));
    let out = '';
    for (const name of parts) out += wordXmlToMarkedText(await zip.file(name).async('string')) + '\n';
    return out;
  };

  // RTF: file .doc/.rtf do máy Mac hoặc phần mềm soạn thảo khác xuất ra.
  const rtfToText = (rtf) => rtf
    .replace(/\{\\\*?\\(?:fonttbl|colortbl|stylesheet|info|generator|pict|themedata|colorschememapping)[\s\S]*?\}\s*/gi, '')
    .replace(/\\'([0-9a-f]{2})/gi, (m, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/\\u(-?\d+)\s?\??/g, (m, n) => String.fromCharCode(((+n) + 65536) % 65536))
    .replace(/\\pard\b/g, '')
    .replace(/\\(?:par|line)\b/g, '\n')
    .replace(/\\tab\b/g, '\t')
    .replace(/\\[a-z]+-?\d*\s?/gi, '')
    .replace(/[{}]/g, '')
    .replace(/\r\n?/g, '\n');

  const htmlToMarkedText = html => markStyledRuns(html)
    .replace(/<\/(p|div|li|tr|h[1-6])>/gi, '\n')
    .replace(/<\/t[dh]>/gi, '\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/?(strong|u)\b[^>]*>/gi, BOLD)
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"').replace(/&#39;/gi, "'")
    .replace(/&amp;/gi, '&');

  // Parser cho format a) b) c) d) với X mark (định dạng mới)
  const parseWordTextNewFormat = (text) => {
    const norm = text
      .replace(/\r\n?/g, '\n')
      .replace(/ /g, ' ');

    const lines = norm.split('\n').map(l => l.trim()).filter(l => l);
    const parsed = [];
    let i = 0;

    while (i < lines.length) {
      const line = lines[i];
      // Tìm dấu hiệu câu hỏi: "01", "02", "03"... hoặc "1.", "2."
      const qMatch = line.match(/^(\d{1,2})\s*[\.\)]?\s*(.*)/);
      if (!qMatch) { i++; continue; }

      const qNum = qMatch[1];
      const firstLineText = qMatch[2];
      const qText = (firstLineText || lines[i + 1] || '').trim();

      if (!qText) { i++; continue; }

      const opts = [];
      let ansIdx = -1;
      const optMarks = {};
      i++;

      // Thu thập các dòng tiếp theo đến khi gặp câu hỏi mới hoặc end
      while (i < lines.length) {
        const currLine = lines[i];
        if (/^\d{1,2}\s*[\.\)]?\s+/.test(currLine) && /^[a-d]\)/.test(lines[i - 1] || '')) break;

        const optMatch = currLine.match(/^([a-d])\)\s*(.*)/);
        if (optMatch) {
          const optIdx = optMatch[1].charCodeAt(0) - 97; // a=0, b=1, c=2, d=3
          optMarks[optIdx] = optMatch[2].trim();
          opts[optIdx] = optMatch[2].trim();
          i++;

          // Kiểm tra dòng tiếp theo xem có X mark không
          if (i < lines.length && lines[i].trim() === 'X') {
            ansIdx = optIdx;
            i++;
          }
        } else if (/^X$|^x$/.test(currLine)) {
          // X ở dòng riêng → đáp án cho option trước đó
          if (opts.length > 0) {
            ansIdx = opts.length - 1;
          }
          i++;
        } else {
          i++;
        }
      }

      const opts_clean = opts.filter(o => o).slice(0, 4);
      if (opts_clean.length >= 2) {
        parsed.push({
          id: Date.now() + Math.random(),
          text: qText,
          opts: opts_clean,
          ans: ansIdx >= 0 && ansIdx < opts_clean.length ? ansIdx : 0,
          topic: 'Nội quy',
          level: 'Dễ',
          ...(ansIdx < 0 ? { needsReview: true, _issues: ['chưa rõ đáp án đúng'] } : {})
        });
      }
    }

    return { parsed, found: parsed.length };
  };

  const parseWordText = (text) => {
    const norm = text
      .replace(/\r\n?/g, '\n')
      .replace(/\u00a0/g, ' ')
      // "...?A. xxx" → tách nhãn phương án bị dính ngay sau dấu câu
      .replace(/([?”’"):.])(?=\u0001*[A-Ja-j]\s*[\.\)]\s)/g, '$1\n')
      // "...D. xxx  Câu 2: ..." / "Câu hỏi 2: ..." → tách mốc câu bị dính giữa dòng
      .replace(/([^\n])[ \t]+(?=\u0001*c[âa]u(?:\s+hỏi)?\s*\d+\s*[\.\:\)\-–]?\s)/gi, '$1\n');

    // Vị trí bắt đầu của từng câu. Đây là căn cứ duy nhất để đếm số câu trong file.
    const findMarks = re => {
      const out = []; let m;
      while ((m = re.exec(norm)) !== null) {
        out.push(m.index + m[1].length + m[2].length);
        if (re.lastIndex === m.index) re.lastIndex++;
      }
      return out;
    };
    // Lớp ký tự "rác" đứng đầu dòng: khoảng trắng, bullet, và dấu in đậm (BOLD) —
    // dòng "Câu N" rất hay được bôi đậm nên dấu này nằm ngay trước chữ "Câu".
    let marks = findMarks(/(^|\n)([ \t>•\-*\u0001]*)(c[âa]u(?:\s+hỏi)?\s*\d+\s*[\.\:\)\-–]?)/gi);
    // File không dùng chữ "Câu" thì thử kiểu đánh số thuần "1." / "1)"
    if (!marks.length) marks = findMarks(/(^|\n)([ \t>•\-*\u0001]*)(\d{1,3}\s*[\.\)]\s)/g);
    if (!marks.length) return { parsed: [], found: 0 };

    // Word hay trình bày phương án dạng bảng/2 cột nên mammoth dồn 2 phương án vào
    // một dòng ("A. xxx  B. yyy") — tách để mỗi phương án / dòng meta nằm riêng.
    const splitRe = /\s+(?=\u0001*(?:(?:[A-Ja-j]\s*[\.\)]\s)|(?:Đáp\s*án(?:\s*đúng)?|ĐA|Chủ\s*đề|Mức\s*độ)\s*[:\.\-]))/i;
    const strip = s => (s || '').split(BOLD).join('').trim();
    const parsed = [];

    marks.forEach((start, bi) => {
      const block = norm.slice(start, bi + 1 < marks.length ? marks[bi + 1] : undefined);
      const lines = [];
      for (const raw of block.split('\n')) {
        const t = raw.trim();
        if (!t) continue;
        t.split(splitRe).forEach(s => { const x = s.trim(); if (x) lines.push(x); });
      }

      const issues = [];
      const qText = strip(lines[0]).replace(/^(?:c[âa]u(?:\s+hỏi)?\s*\d+|\d{1,3})\s*[\.\:\)\-–]?\s*/i, '').trim();

      let ansIdx = -1, topic = 'Nội quy', level = 'Dễ';
      const cand = [];
      for (let i = 1; i < lines.length; i++) {
        const bare = strip(lines[i]);
        const a = bare.match(/^(?:Đáp\s*án(?:\s*đúng)?(?:\s*là)?|ĐA|Answer|Key)\s*[:\.\-–]?\s*([A-Ja-j])\b/i);
        if (a) { ansIdx = OPT_LETTERS.indexOf(a[1].toUpperCase()); continue; }
        const tp = bare.match(/^Chủ\s*đề\s*[:\.\-]\s*(.+)/i);
        if (tp) { topic = tp[1].trim(); continue; }
        const lv = bare.match(/^Mức\s*độ\s*[:\.\-]\s*(.+)/i);
        if (lv) { level = lv[1].trim(); continue; }
        cand.push(lines[i]);
      }

      // Xếp các dòng phương án vào đúng ô A..J. Nhãn tường minh ("C.") sẽ đồng bộ
      // lại vị trí, nên phương án bị mất nhãn vẫn rơi đúng chỗ.
      const slots = [];
      let next = 0;
      for (const c of cand) {
        const m = c.match(/^\u0001*([A-Ja-j])\s*[\.\)]\s*(.*)/);
        if (m) { const idx = OPT_LETTERS.indexOf(m[1].toUpperCase()); slots[idx] = m[2] || ''; next = idx + 1; }
        else { slots[next] = c; next++; }
      }
      // Lấy liên tiếp từ ô A để chỉ số đáp án không bị lệch
      const rawOpts = [];
      for (let k = 0; k < MAX_OPTS; k++) {
        if (slots[k] === undefined || !strip(slots[k])) break;
        rawOpts.push(slots[k]);
      }

      // Không có dòng "Đáp án" thì suy từ phương án được bôi đậm — chỉ tin khi
      // đúng một phương án được bôi (cả khối cùng đậm thì tín hiệu vô nghĩa).
      if (ansIdx < 0) {
        const bold = rawOpts.map((o, i) => (o.includes(BOLD) ? i : -1)).filter(i => i >= 0);
        if (bold.length === 1) ansIdx = bold[0];
      }

      const opts = rawOpts.map(strip);
      if (cand.length > MAX_OPTS) issues.push(`chỉ giữ ${MAX_OPTS} phương án đầu`);
      if (opts.length < 2) { issues.push('thiếu phương án trả lời'); while (opts.length < 2) opts.push(''); }
      if (ansIdx < 0) { issues.push('chưa rõ đáp án đúng'); ansIdx = 0; }
      else if (ansIdx >= opts.length) { issues.push(`đáp án ${OPT_LETTERS[ansIdx]} không có phương án tương ứng`); ansIdx = 0; }
      if (!qText) issues.push('thiếu nội dung câu hỏi');

      parsed.push({
        id: Date.now() + Math.random(),
        text: qText || `(Câu ${bi + 1} — chưa đọc được nội dung, cần nhập tay)`,
        opts, ans: ansIdx, topic, level,
        ...(issues.length ? { needsReview: true, _issues: issues } : {}),
      });
    });

    return { parsed, found: marks.length };
  };

  // Lỗi kỹ thuật của thư viện đọc file không nói lên điều gì với người dùng, nên
  // dịch sang việc cần làm tiếp theo.
  const wordImportError = (kind, gotText, details) => {
    const detail = details.length ? ` (chi tiết kỹ thuật: ${details[0]})` : '';
    if (gotText) {
      return 'Đọc được nội dung file nhưng không tìm thấy câu hỏi nào. Mỗi câu phải bắt đầu bằng "Câu 1:", "Câu hỏi 1:"... Bấm "File mẫu" để xem đúng định dạng.';
    }
    if (kind === 'zip') {
      return 'File .docx này không đúng chuẩn Word nên không lấy được nội dung — thường gặp ở file do phần mềm khác tạo ra (chuyển từ PDF, ứng dụng đọc file trên điện thoại...). Hãy mở file bằng Word rồi chọn File → Save As → Word Document (*.docx) và import lại.' + detail;
    }
    if (kind === 'ole') {
      return 'File .doc (Word 97–2003) bị hỏng, đặt mật khẩu hoặc thuộc phiên bản quá cũ. Hãy mở bằng Word rồi lưu lại dưới dạng .docx và import lại.' + detail;
    }
    return 'File không có nội dung chữ nào đọc được. Hãy mở bằng Word rồi lưu lại dưới dạng .docx và import lại.' + detail;
  };

  const handleWordImport = async (e) => {
    const file = e.target.files[0]; if(!file) return;
    setImporting(true); setImportResult(null);
    try {
      const buffer = await file.arrayBuffer();
      const bytes = new Uint8Array(buffer);
      const kind = sniffFileKind(buffer);   // tin chữ ký trong file, không tin phần mở rộng
      const details = [];
      let gotText = false;

      // Thử nhiều cách đọc rồi lấy cách nhận ra nhiều câu hơn — mục tiêu là không sót câu.
      let best = { parsed: [], found: 0 };
      const read = async (fn) => {
        try {
          const text = await fn();
          if (!text || !text.trim()) return;
          gotText = true;
          // Thử parser format cũ (Câu 1, A. B. C. D.)
          const res1 = parseWordText(text);
          if (res1.found > best.found) best = res1;
          // Thử parser format mới (01, a) b) c) d) với X mark)
          const res2 = parseWordTextNewFormat(text);
          if (res2.found > best.found) best = res2;
        } catch (err) { details.push(err.message); }
      };

      if (kind === 'zip') {
        // .docx: bản HTML giữ được in đậm/gạch chân (dấu hiệu đáp án đúng), bản thô để đối chiếu.
        await read(async () => htmlToMarkedText((await mammoth.convertToHtml({ arrayBuffer: buffer }, { styleMap: ['u => u'] })).value));
        await read(async () => (await mammoth.extractRawText({ arrayBuffer: buffer })).value);
        // mammoth chỉ đọc được file đúng chuẩn Word; file lệch chuẩn thì bóc thẳng XML.
        if (best.found === 0) await read(() => zipDocumentToMarkedText(buffer));
      } else if (kind === 'ole') {
        // Word 97–2003 (.doc) là định dạng OLE nhị phân; mammoth chỉ hỗ trợ .docx.
        // legacy-doc-reader chạy hoàn toàn trong trình duyệt và giữ tài liệu ở máy người dùng.
        await read(() => { const s = legacyDocToText.html(buffer); return s && s.body ? htmlToMarkedText(s.body) : ''; });
        await read(() => legacyDocToText(buffer) || '');
      } else if (kind === 'html') {
        // File Word lưu dạng trang web, hoặc chính file mẫu tải về từ nút "File mẫu".
        await read(() => htmlToMarkedText(decodeBytes(bytes)));
      } else if (kind === 'rtf') {
        await read(() => rtfToText(decodeBytes(bytes)));
      } else {
        await read(() => decodeBytes(bytes));
      }

      if (best.found === 0) {
        setImportResult({ error: wordImportError(kind, gotText, details) });
      } else {
        // Giữ dấu vết file nguồn để lúc tạo đề có thể chọn cả file bằng một lần bấm.
        const importBatchId = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
        setPreviewList(best.parsed.map(q => ({
          ...q,
          importBatchId,
          importFileName: file.name.replace(/\.(?:docx?|docm|rtf|txt|html?|odt)$/i, ''),
        })));
      }
    } catch(err) {
      setImportResult({error:'Không đọc được file Word (.doc hoặc .docx) — lỗi: ' + err.message});
    }
    setImporting(false);
    e.target.value='';
  };

  const confirmImport = () => {
    const n = previewList.length;
    const review = previewList.filter(q => q.needsReview).length;
    setQuestions(p=>{
      // Chuẩn hóa câu cũ trước, rồi nối file mới ngay sau câu cuối:
      // file đầu 1..40, file tiếp theo 41..80, v.v.
      const existing = orderedQuestions(p).map((q, i) => ({ ...q, order: i + 1 }));
      const base = existing.length;
      const ordered = previewList.map((q,i)=>{
        const { _issues, ...rest } = q;   // _issues chỉ dùng để hiển thị ở màn xem trước
        return { ...rest, order: base + i + 1 };
      });
      return [...existing, ...ordered];
    });
    setImportResult({added:n, review});
    setPreviewList(null);
    setTimeout(()=>setImportResult(null), 6000);
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
          <div className="bg-white rounded-2xl w-full max-w-xl max-h-[85vh] flex flex-col shadow-2xl">
            <div className="flex items-center justify-between p-5 border-b">
              <div>
                <h2 className="font-bold text-slate-800">Xem trước câu hỏi import</h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Đọc được <span className="font-semibold text-emerald-600">{previewList.length} câu</span> — sẽ nhập đủ cả {previewList.length} câu
                  {previewList.some(q=>q.needsReview) && <>, trong đó <span className="font-semibold text-amber-600">{previewList.filter(q=>q.needsReview).length} câu cần sửa tay</span></>}
                </p>
              </div>
              <button onClick={()=>setPreviewList(null)} className="text-slate-400 hover:text-slate-600"><X size={18}/></button>
            </div>
            <div className="flex-1 overflow-y-auto p-4 space-y-3">
              {previewList.map((q,i)=>(
                <div key={i} className={`rounded-xl p-4 border ${q.needsReview?'bg-amber-50 border-amber-200':'bg-slate-50 border-slate-100'}`}>
                  <div className="flex items-center gap-2 mb-2 flex-wrap">
                    <span className="text-xs font-mono text-slate-400">#{i+1}</span>
                    {q.needsReview && (
                      <span className="flex items-center gap-1 text-[11px] font-semibold text-amber-700 bg-amber-100 px-2 py-0.5 rounded-full">
                        <AlertCircle size={11}/>Cần sửa tay: {q._issues.join(', ')}
                      </span>
                    )}
                  </div>
                  <p className="text-sm font-medium text-slate-700 mb-2">{q.text}</p>
                  <div className="grid grid-cols-2 gap-1">
                    {q.opts.map((o,j)=>(
                      <div key={j} className={`text-xs px-2.5 py-1.5 rounded-lg ${j===q.ans?'bg-emerald-50 text-emerald-700 font-medium border border-emerald-200':'bg-white text-slate-500 border border-slate-100'}`}>
                        {OPT_LETTERS[j]}. {o || <span className="italic text-slate-300">(trống)</span>}
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
            <Upload size={14}/>{importing?'Đang đọc...':'Import Word (.doc/.docx)'}
          </button>
          <input ref={wordRef} type="file" accept=".doc,.docx,.docm,.rtf,.txt,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" className="hidden" onChange={handleWordImport}/>
          <button onClick={openAdd} className="flex items-center gap-2 bg-emerald-600 text-white px-4 py-2 rounded-lg hover:bg-emerald-700 text-sm font-medium"><Plus size={15}/>Thêm câu hỏi</button>
        </div>
      </div>

      {importResult && (
        <div className={`mb-4 rounded-xl p-4 flex items-start gap-3 ${importResult.error?'bg-red-50 border border-red-200':'bg-emerald-50 border border-emerald-200'}`}>
          {importResult.error ? <AlertCircle size={16} className="text-red-500 flex-shrink-0 mt-0.5"/> : <CheckCircle size={16} className="text-emerald-600 flex-shrink-0 mt-0.5"/>}
          <p className={`text-sm flex-1 ${importResult.error?'text-red-700':'text-emerald-800 font-medium'}`}>
            {importResult.error || (
              <>
                Đã thêm đủ {importResult.added} câu hỏi vào ngân hàng!
                {importResult.review > 0 && <span className="block font-normal text-amber-700 mt-1">
                  Trong đó {importResult.review} câu còn thiếu thông tin (đánh dấu <b>Cần sửa</b> ở danh sách bên dưới) — bấm sửa lại đáp án đúng trước khi đưa vào đề thi.
                </span>}
              </>
            )}
          </p>
          <button onClick={()=>setImportResult(null)} className="text-slate-400 hover:text-slate-600"><X size={14}/></button>
        </div>
      )}
      {modal && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl w-full max-w-lg max-h-[90vh] overflow-y-auto shadow-2xl">
            <div className="flex items-center justify-between p-5 border-b sticky top-0 bg-white z-10">
              <h2 className="font-bold text-slate-800">{editingId!==null?'Chỉnh sửa câu hỏi':'Thêm câu hỏi mới'}</h2>
              <button onClick={resetQuestionForm} className="text-slate-400 hover:text-slate-600"><X size={18}/></button>
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
                <div><label className="text-xs font-medium text-slate-600 mb-1 block">Mức độ</label>
                  <select className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" value={nq.level} onChange={e=>setNq({...nq,level:e.target.value})}>
                    <option value="Dễ">Dễ</option><option value="TB">Trung bình</option><option value="Khó">Khó</option>
                  </select>
                </div>
              </div>
              <div><label className="text-xs font-medium text-slate-600 mb-1 block">Nội dung câu hỏi</label>
                <textarea className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none" rows={3} value={nq.text} onChange={e=>setNq({...nq,text:e.target.value})} placeholder="Nhập nội dung câu hỏi..."/></div>
              <div><label className="text-xs font-medium text-slate-600 mb-2 block">Các đáp án (chọn đáp án đúng)</label>
                <div className="space-y-2">
                  {nq.opts.map((opt,i)=>(
                    <div key={i} className="flex items-center gap-2">
                      <button type="button" onClick={()=>setNq({...nq,ans:i})} className={`w-7 h-7 rounded-full flex-shrink-0 border-2 flex items-center justify-center text-xs font-bold transition-all ${nq.ans===i?'border-emerald-600 bg-emerald-600 text-white':'border-slate-300 text-slate-400 hover:border-emerald-500'}`}>{OPT_LETTERS[i]}</button>
                      <input className="flex-1 border border-slate-200 rounded-lg px-3 py-1.5 text-sm" value={opt} onChange={e=>{const o=[...nq.opts];o[i]=e.target.value;setNq({...nq,opts:o});}} placeholder={`Đáp án ${OPT_LETTERS[i]}`}/>
                      {nq.opts.length>2 && <button type="button" title="Xóa phương án" onClick={()=>setNq(prev=>{
                        const opts=prev.opts.filter((_,j)=>j!==i);
                        const ans=prev.ans===i?0:(prev.ans>i?prev.ans-1:prev.ans);
                        return {...prev,opts,ans};
                      })} className="text-slate-300 hover:text-red-500"><X size={15}/></button>}
                    </div>
                  ))}
                  {nq.opts.length<MAX_OPTS && <button type="button" onClick={()=>setNq({...nq,opts:[...nq.opts,'']})} className="text-xs text-emerald-700 hover:text-emerald-800 font-medium">+ Thêm phương án</button>}
                </div>
              </div>
              <div className="flex gap-2 pt-1">
                <button onClick={resetQuestionForm} className="flex-1 py-2 border border-slate-200 rounded-lg text-sm text-slate-600 hover:bg-slate-50">Hủy</button>
                <button onClick={saveQuestion} className="flex-1 py-2 bg-emerald-600 text-white rounded-lg text-sm hover:bg-emerald-700">{editingId!==null?'Lưu thay đổi':'Thêm câu hỏi'}</button>
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
                <div className="flex items-center gap-2 mb-2 flex-wrap">
                  <span className="text-xs font-mono text-slate-400">#{q.order ?? idx+1}</span>
                  {q.needsReview && (
                    <span className="flex items-center gap-1 text-[11px] font-semibold text-amber-700 bg-amber-100 px-2 py-0.5 rounded-full">
                      <AlertCircle size={11}/>Cần sửa
                    </span>
                  )}
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
              <div className="flex items-center gap-1 flex-shrink-0 mt-1">
                <button title="Sửa câu hỏi" onClick={()=>openEdit(q)} className="w-8 h-8 flex items-center justify-center rounded-lg border border-slate-200 text-slate-500 hover:border-emerald-300 hover:text-emerald-600"><FileText size={14}/></button>
                <button title="Xóa câu hỏi" onClick={()=>setQuestions(p=>p.filter(x=>x.id!==q.id))} className="w-8 h-8 flex items-center justify-center rounded-lg border border-slate-200 text-slate-400 hover:border-red-300 hover:text-red-500"><Trash2 size={14}/></button>
              </div>
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
  const [ne, setNe] = useState({title:'',desc:'',qIds:[],time:20,pass:70,sessions:[]});
  const [editingExamId, setEditingExamId] = useState(null);
  const toggleQ = id => setNe(p=>({...p,qIds:p.qIds.includes(id)?p.qIds.filter(x=>x!==id):[...p.qIds,id]}));
  const sortedQuestions = orderedQuestions(questions);
  const importedGroups = Object.values(sortedQuestions.reduce((groups, q) => {
    if (!q.importBatchId) return groups;
    if (!groups[q.importBatchId]) groups[q.importBatchId] = {
      id: q.importBatchId,
      name: q.importFileName || 'File Word đã import',
      qIds: [],
    };
    groups[q.importBatchId].qIds.push(q.id);
    return groups;
  }, {}));
  // Dữ liệu đã import trước khi có importBatchId vẫn được chọn nhanh theo từng khối 40 câu.
  const legacyQuestions = sortedQuestions.filter(q => !q.importBatchId);
  const legacyGroups = [];
  for (let i = 0; i < legacyQuestions.length; i += 40) {
    const chunk = legacyQuestions.slice(i, i + 40);
    const first = chunk[0]?.order ?? i + 1;
    const last = chunk[chunk.length - 1]?.order ?? i + chunk.length;
    legacyGroups.push({
      id: `legacy-${i}`,
      name: `Câu ${first}–${last} (dữ liệu cũ)`,
      qIds: chunk.map(q => q.id),
    });
  }
  const importGroups = [...importedGroups, ...legacyGroups];
  const toggleImportGroup = group => setNe(p => {
    const allSelected = group.qIds.every(id => p.qIds.includes(id));
    return {
      ...p,
      qIds: allSelected
        ? p.qIds.filter(id => !group.qIds.includes(id))
        : [...new Set([...p.qIds, ...group.qIds])],
    };
  });
  const create = () => {
    if(!ne.title.trim()||ne.qIds.length===0) return;
    if(editingExamId) {
      setExams(p=>p.map(e=>e.id===editingExamId?{...ne,id:editingExamId}:e));
      setEditingExamId(null);
    } else {
      setExams(p=>[...p,{...ne,id:Date.now()}]);
    }
    setModal(false);
    setNe({title:'',desc:'',qIds:[],time:20,pass:70,sessions:[]});
  };

  const generateCode = () => Math.random().toString(36).substring(2,6).toUpperCase();

  const addSession = () => {
    const newSession = {id:Date.now(),name:`Kíp ${ne.sessions.length+1}`,code:generateCode()};
    setNe(p=>({...p,sessions:[...p.sessions,newSession]}));
  };

  const removeSession = (id) => {
    setNe(p=>({...p,sessions:p.sessions.filter(s=>s.id!==id)}));
  };

  const updateSession = (id, field, value) => {
    setNe(p=>({...p,sessions:p.sessions.map(s=>s.id===id?{...s,[field]:value}:s)}));
  };

  const openEdit = (exam) => {
    setNe({...exam});
    setEditingExamId(exam.id);
    setModal(true);
  };

  const toggleLock = (examId) => {
    setExams(p=>p.map(e=>e.id===examId?{...e,locked:!e.locked}:e));
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
              <h2 className="font-bold text-slate-800">{editingExamId?'Sửa đề thi':'Tạo đề thi mới'}</h2>
              <button onClick={()=>{setModal(false);setEditingExamId(null);setNe({title:'',desc:'',qIds:[],time:20,pass:70,sessions:[]});}} className="text-slate-400 hover:text-slate-600"><X size={18}/></button>
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
              <div><label className="text-xs font-medium text-slate-600 mb-2 block">Quản lý mã kíp thi</label>
                <div className="space-y-2 mb-3">
                  {ne.sessions.map(session=>(
                    <div key={session.id} className="flex items-center gap-2 p-3 bg-slate-50 rounded-lg border border-slate-100">
                      <div className="flex-1">
                        <input type="text" className="w-full border border-slate-200 rounded-lg px-2 py-1 text-xs mb-1" value={session.name} onChange={e=>updateSession(session.id,'name',e.target.value)} placeholder="Tên kíp"/>
                        <div className="flex items-center gap-2">
                          <span className="text-xs text-slate-500">Mã:</span>
                          <code className="text-xs font-mono bg-white border border-slate-200 rounded px-2 py-1 flex-1">{session.code}</code>
                          <button type="button" onClick={()=>updateSession(session.id,'code',generateCode())} className="text-xs px-2 py-1 bg-blue-50 text-blue-600 rounded hover:bg-blue-100">Tạo mới</button>
                        </div>
                      </div>
                      <button type="button" onClick={()=>removeSession(session.id)} className="text-red-400 hover:text-red-600 flex-shrink-0"><Trash2 size={14}/></button>
                    </div>
                  ))}
                  <button type="button" onClick={addSession} className="w-full py-2 border border-emerald-300 bg-emerald-50 text-emerald-700 rounded-lg text-xs font-medium hover:bg-emerald-100">
                    <Plus size={12} className="inline mr-1"/>Thêm kíp thi
                  </button>
                </div>
              </div>
              <div><label className="text-xs font-medium text-slate-600 mb-2 block">Chọn câu hỏi ({ne.qIds.length} đã chọn)</label>
                {importGroups.length > 0 && (
                  <div className="space-y-2 mb-3">
                    <p className="text-[11px] text-slate-400">Chọn nhanh theo file Word đã import</p>
                    {importGroups.map(group => {
                      const selected = group.qIds.every(id => ne.qIds.includes(id));
                      return (
                        <button key={group.id} type="button" onClick={()=>toggleImportGroup(group)}
                          className={`w-full flex items-center justify-between gap-3 px-3 py-2 rounded-lg border text-left ${selected?'bg-emerald-50 border-emerald-300 text-emerald-700':'bg-white border-slate-200 text-slate-600 hover:border-emerald-300'}`}>
                          <span className="text-xs font-medium truncate">{group.name}</span>
                          <span className="text-[11px] flex-shrink-0">{selected?'Đã chọn':'Chọn cả'} {group.qIds.length} câu</span>
                        </button>
                      );
                    })}
                  </div>
                )}
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
          <div key={exam.id} className={`rounded-xl p-5 shadow-sm border ${exam.locked?'bg-slate-50 border-slate-200':'bg-white border-slate-100'}`}>
            <div className="flex items-start justify-between gap-4">
              <div className="flex-1">
                <div className="flex items-center gap-2 mb-1">
                  <h3 className="font-semibold text-slate-800">{exam.title}</h3>
                  {exam.locked && (
                    <span className="flex items-center gap-1 px-2 py-0.5 bg-amber-50 border border-amber-200 rounded-full text-xs font-medium text-amber-700">
                      <Lock size={11}/>Đã khoá
                    </span>
                  )}
                </div>
                <p className="text-sm text-slate-500 mb-3">{exam.desc}</p>
                <div className="flex gap-4 text-xs text-slate-500 flex-wrap">
                  <span className="flex items-center gap-1"><FileText size={12}/>{exam.qIds.length} câu hỏi</span>
                  <span className="flex items-center gap-1"><Clock size={12}/>{exam.time} phút</span>
                  <span className="flex items-center gap-1"><Award size={12}/>Điểm đạt: {exam.pass}%</span>
                  {exam.sessions && exam.sessions.length > 0 && (
                    <span className="flex items-center gap-1"><Shield size={12}/>{exam.sessions.length} kíp thi</span>
                  )}
                </div>
              </div>
              <div className="flex gap-2 flex-shrink-0">
                <button onClick={()=>toggleLock(exam.id)} className={`${exam.locked?'text-amber-600 hover:text-amber-700':'text-slate-400 hover:text-amber-600'}`} title={exam.locked?'Mở khoá':'Khoá đề thi'}>{exam.locked?<Lock size={16}/>:<LockOpen size={16}/>}</button>
                <button onClick={()=>openEdit(exam)} disabled={exam.locked} className={`${exam.locked?'text-slate-300 cursor-not-allowed':'text-slate-400 hover:text-blue-600'}`} title={exam.locked?'Không thể sửa đề đã khoá':'Chỉnh sửa'}><FileText size={16}/></button>
                <button onClick={()=>setExams(p=>p.filter(e=>e.id!==exam.id))} disabled={exam.locked} className={`${exam.locked?'text-slate-300 cursor-not-allowed':'text-slate-300 hover:text-red-400'}`} title={exam.locked?'Không thể xóa đề đã khoá':'Xóa'}><Trash2 size={16}/></button>
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

// ── EMPLOYEES ──
const EmployeesView = ({employees, setEmployees, results, exams, onProvision}) => {
  const [creds, setCreds] = useState(null);   // kết quả cấp tài khoản để hiện cho admin
  const runProvision = async (emps, replace=false) => {
    const r = await onProvision(emps, replace);
    if (r.creds.length || r.fails.length) setCreds(r);
  };
  const [importing, setImporting] = useState(false);
  const [importResult, setImportResult] = useState(null);
  const [modal, setModal] = useState(null); // null | {mode:'add'} | {mode:'edit', emp} | {mode:'delete', emp}
  const [form, setForm] = useState({name:'', dept:'', account:''});
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
    setForm({name:'', dept: depts.length > 0 ? depts[0] : '__custom__', account:''});
    setCustomDept('');
    setFormErr('');
    setModal({mode:'add'});
  };
  const openEdit = (emp) => { setForm({name:emp.name, dept:emp.dept, account:emp.account||''}); setCustomDept(''); setFormErr(''); setModal({mode:'edit', emp, password:'1'}); };
  const openDelete = (emp) => setModal({mode:'delete', emp});
  const closeModal = () => { setModal(null); setFormErr(''); };

  const effectiveDept = form.dept==='__custom__' ? customDept.trim() : form.dept;

  const handleSave = () => {
    const name = form.name.trim();
    const dept = form.dept === '__custom__' ? customDept.trim() : form.dept;
    const account = form.account.trim();
    if(!name) return setFormErr('Vui lòng nhập họ và tên');
    if(!dept) return setFormErr('Vui lòng chọn hoặc nhập đơn vị');
    if(!account) return setFormErr('Vui lòng nhập tài khoản');
    if(modal.mode==='add') {
      const dup = employees.find(e=>e.name.toLowerCase()===name.toLowerCase());
      if(dup) return setFormErr('Thí sinh này đã tồn tại trong hệ thống');
      const dupAcc = employees.find(e=>e.account===account);
      if(dupAcc) return setFormErr('Tài khoản này đã được sử dụng');
      const newEmp = {id:Date.now(), name, dept, account};
      setEmployees(p=>[...p, newEmp]);
      runProvision([newEmp]);
    } else {
      const dup = employees.find(e=>e.name.toLowerCase()===name.toLowerCase() && e.id!==modal.emp.id);
      if(dup) return setFormErr('Tên này đã được dùng cho thí sinh khác');
      const dupAcc = employees.find(e=>e.account===account && e.id!==modal.emp.id);
      if(dupAcc) return setFormErr('Tài khoản này đã được sử dụng');
      setEmployees(p=>p.map(e=>e.id===modal.emp.id?{...e,name,dept,account}:e));
      // Đổi tên đăng nhập → cấp tài khoản mới (mật khẩu mới), thu hồi tài khoản cũ
      if(account !== (modal.emp.account||'')) runProvision([{...modal.emp, name, dept, account}], true);
    }
    closeModal();
  };

  const handleDelete = () => {
    setEmployees(p=>p.filter(e=>e.id!==modal.emp.id));
    closeModal();
  };

  const handleDeleteAll = () => {
    setEmployees([]);
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

      // Tìm index cột Họ và tên, Đơn vị, Số điện thoại
      let colName = -1, colDept = -1, colPhone = -1;
      for(let i = 0; i < Math.min(5, rows.length); i++) {
        const row = rows[i].map(c => String(c).trim().toLowerCase());
        const ni = row.findIndex(c => c.includes('họ') || c.includes('tên') || c === 'name');
        const di = row.findIndex(c => c.includes('đơn vị') || c.includes('phòng') || c.includes('dept') || c.includes('khoa'));
        const pi = row.findIndex(c => c.includes('điện thoại') || c.includes('phone') || c.includes('tài khoản') || c.includes('số'));
        if(ni !== -1) { colName = ni; colDept = di !== -1 ? di : ni+1; colPhone = pi !== -1 ? pi : ni+2; break; }
      }
      if(colName === -1) { colName = 1; colDept = 2; colPhone = 0; }

      const skipReasons = {empty: 0, header: 0, khoi: 0, dupName: 0, dupAcc: 0};

      rows.forEach(row => {
        const nameRaw = String(row[colName]||'').trim();
        const deptRaw = String(colDept >= 0 ? (row[colDept]||'') : '').trim();
        const phoneRaw = String(colPhone >= 0 ? (row[colPhone]||'') : '').trim();
        if(!nameRaw) { skipReasons.empty++; return; }
        if(nameRaw.toUpperCase().includes('KHỐI')) { skipReasons.khoi++; return; }
        if(nameRaw.toLowerCase().includes('họ và tên') || nameRaw.toLowerCase() === 'name') { skipReasons.header++; return; }
        const dup = employees.find(e => e.name.toLowerCase() === nameRaw.toLowerCase());
        if(dup) { skipReasons.dupName++; skipped.push(`${nameRaw} (tên trùng)`); return; }
        added.push({id: Date.now() + Math.random(), name: nameRaw, dept: deptRaw || 'Chưa phân công', account: phoneRaw});
      });

      if(added.length > 0) { setEmployees(p => [...p, ...added]); runProvision(added); }
      const totalSkipped = Object.values(skipReasons).reduce((a,b) => a+b, 0);
      const skipMsg = totalSkipped > 0 ? ` — ${skipReasons.header} header, ${skipReasons.empty} tên trống, ${skipReasons.khoi} khối, ${skipReasons.dupName} trùng tên` : '';
      setImportResult({added: added.length, skipped: skipped.length, totalSkipped, skipMsg, review: skipped.slice(0, 10)});
      setTimeout(() => setImportResult(null), 5000);
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
    ['TT', 'Họ và tên', 'Số điện thoại', 'Đơn vị'],
    ['KHỐI CƠ QUAN', '', '', ''],
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

  const naturalSort = (a, b) => {
    const regex = /(\d+)/g;
    const aParts = a.split(regex);
    const bParts = b.split(regex);
    for(let i=0; i<Math.min(aParts.length, bParts.length); i++){
      const aNum = /^\d+$/.test(aParts[i]);
      const bNum = /^\d+$/.test(bParts[i]);
      if(aNum && bNum) { const cmp = +aParts[i] - +bParts[i]; if(cmp!==0) return cmp; }
      else if(aNum || bNum) return aNum ? 1 : -1;
      else { const cmp = aParts[i].localeCompare(bParts[i]); if(cmp!==0) return cmp; }
    }
    return aParts.length - bParts.length;
  };

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
      {creds && <CredentialsModal data={creds} onClose={()=>setCreds(null)}/>}
      {/* Modal Thêm / Sửa */}
      {(modal?.mode==='add' || modal?.mode==='edit') && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl w-full max-w-md shadow-2xl">
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
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs font-medium text-slate-600 mb-1.5 block">Tài khoản <span className="text-red-400">*</span></label>
                  <input
                    className="w-full border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm focus:outline-none focus:border-emerald-500"
                    placeholder="VD: nguyenvanA"
                    value={form.account}
                    onChange={e=>{ setForm(f=>({...f,account:e.target.value})); setFormErr(''); }}
                  />
                </div>
                <div>
                  <label className="text-xs font-medium text-slate-600 mb-1.5 block">Mật khẩu</label>
                  <input
                    className="w-full border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm bg-slate-50 text-slate-600"
                    value={form.account || '(chưa có tài khoản)'}
                    disabled
                  />
                  <p className="text-xs text-slate-400 mt-1">Tự động = tài khoản</p>
                </div>
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
          <div className="bg-white rounded-2xl w-full max-w-sm sm:max-w-sm shadow-2xl p-6">
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

      {/* Modal Xóa tất cả */}
      {modal?.mode==='deleteAll' && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl w-full max-w-sm shadow-2xl p-6">
            <div className="w-12 h-12 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-4">
              <AlertCircle size={20} className="text-red-500"/>
            </div>
            <h3 className="font-bold text-slate-800 text-center mb-2">Xóa tất cả thí sinh?</h3>
            <p className="text-sm text-slate-500 text-center mb-1">
              Bạn có chắc muốn xóa <span className="font-bold text-red-600">{employees.length}</span> thí sinh?
            </p>
            <p className="text-xs text-slate-400 text-center mb-5">Hành động này KHÔNG THỂ HOÀN TÁC. Lịch sử thi sẽ vẫn được giữ lại.</p>
            <div className="flex gap-2">
              <button onClick={closeModal} className="flex-1 py-2.5 border border-slate-200 rounded-xl text-sm text-slate-600 hover:bg-slate-50">Hủy</button>
              <button onClick={handleDeleteAll} className="flex-1 py-2.5 bg-red-600 text-white rounded-xl text-sm font-medium hover:bg-red-700">Xóa tất cả</button>
            </div>
          </div>
        </div>
      )}

      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between mb-5 gap-3">
        <div className="flex flex-wrap gap-2">
          <button onClick={downloadTemplate} className="flex items-center gap-1.5 px-3 py-2 border border-slate-200 rounded-lg text-sm text-slate-600 hover:bg-slate-50"><Download size={14}/>Tải file mẫu</button>
          <button onClick={()=>fileRef.current.click()} disabled={importing} className="flex items-center gap-1.5 px-3 py-2 border border-slate-200 rounded-lg text-sm text-slate-600 hover:bg-slate-50 disabled:opacity-60"><Upload size={14}/>{importing?'Đang import...':'Import Excel'}</button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleImport}/>
          <button onClick={openAdd} className="flex items-center gap-1.5 bg-emerald-600 text-white px-4 py-2 rounded-lg hover:bg-emerald-700 text-sm font-medium"><Plus size={15}/>Thêm thí sinh</button>
        </div>
        {employees.length > 0 && (
          <button onClick={()=>setModal({mode:'deleteAll'})} className="flex items-center gap-1.5 px-3 py-2 border border-red-200 rounded-lg text-sm text-red-600 hover:bg-red-50"><Trash2 size={14}/>Xóa tất cả</button>
        )}
      </div>

      {importResult && (
        <div className={`mb-4 rounded-xl p-4 flex flex-col gap-2 ${importResult.error?'bg-red-50 border border-red-200':'bg-emerald-50 border border-emerald-200'}`}>
          <div className="flex items-start gap-3">
            {importResult.error ? <AlertCircle size={15} className="text-red-500 flex-shrink-0 mt-0.5"/> : <CheckCircle size={15} className="text-emerald-600 flex-shrink-0 mt-0.5"/>}
            <div className="flex-1">
              <p className={`text-sm font-medium ${importResult.error?'text-red-700':'text-emerald-800'}`}>
                {importResult.error || `✓ Thêm ${importResult.added} thí sinh thành công!`}
              </p>
              {importResult.totalSkipped > 0 && !importResult.error && (
                <p className="text-xs text-emerald-700 mt-1">
                  Bỏ qua {importResult.totalSkipped} dòng{importResult.skipMsg}
                </p>
              )}
              {importResult.review?.length > 0 && (
                <div className="text-xs text-emerald-700 mt-2 space-y-1">
                  {importResult.review.map((r,i) => <div key={i} className="line-clamp-1">• {r}</div>)}
                </div>
              )}
            </div>
            <button onClick={()=>setImportResult(null)} className="text-slate-400 hover:text-slate-600 flex-shrink-0"><X size={14}/></button>
          </div>
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

      {/* Dept filter dropdown + search */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-4 gap-3">
        <select value={activeDept} onChange={e=>{setActiveDept(e.target.value);setPage(1);}} className="px-4 py-2 border border-slate-200 rounded-lg bg-white text-sm font-medium text-slate-700 hover:border-emerald-300 focus:outline-none focus:border-emerald-500">
          <option value="all">Tất cả phòng/khoa ({deptCounts.reduce((a,d)=>a+d.count,0)})</option>
          {[...deptCounts].sort((a,b)=>naturalSort(a.dept, b.dept)).map(({dept,count})=>(
            <option key={dept} value={dept}>{dept} ({count})</option>
          ))}
        </select>
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
              {['Thí sinh','Tài khoản','Mật khẩu','Lượt thi','Đạt','Điểm TB','Kết quả gần nhất','Thao tác'].map(h=>(
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
                  <td colSpan={8} className="px-5 py-2">
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
                      <td className="px-5 py-3 text-sm text-slate-600 font-mono">{emp.account||'--'}</td>
                      <td className="px-5 py-3 text-sm font-mono text-slate-600">1</td>
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
    // Same 3 buckets as the donut: Đạt ≥ điểm đạt, Trung bình 50–69%, Chưa đạt < 50%
    const mid = rs.filter(r=>{const e=exams.find(x=>x.id===r.examId);return r.score>=50&&(!e||r.score<e.pass);}).length;
    const low = rs.filter(r=>r.score<50).length;
    const avg = rs.length?Math.round(rs.reduce((a,r)=>a+r.score,0)/rs.length):0;
    const rate = rs.length?Math.round(passed/rs.length*100):0;
    return {name:dept, attempts:rs.length, passed, mid, low, failed:rs.length-passed, avg, rate};
  });

  const totalAttempts = results.length;
  const globalAvg = totalAttempts?Math.round(results.reduce((a,r)=>a+r.score,0)/totalAttempts):0;
  const bestDept = deptData.reduce((a,b)=>b.avg>a.avg?b:a, deptData[0]||{name:'--',avg:0});

  // "So với kỳ trước" trên các thẻ thống kê: tháng hiện tại so với tháng liền trước.
  // r.date được lưu dạng d/m/yyyy (toLocaleDateString('vi-VN')).
  const monthIdx = d => { const [,m,y] = String(d||'').split('/').map(Number); return (m&&y) ? y*12+(m-1) : null; };
  const curMonth  = new Date().getFullYear()*12 + new Date().getMonth();
  const rsThis = results.filter(r=>monthIdx(r.date)===curMonth);
  const rsPrev = results.filter(r=>monthIdx(r.date)===curMonth-1);
  const pctChange = (cur,prev) => prev ? Math.round((cur-prev)/prev*100) : (cur?100:0);
  const avgOf = rs => rs.length?Math.round(rs.reduce((a,r)=>a+r.score,0)/rs.length):0;
  const deptCountOf = rs => new Set(rs.map(r=>employees.find(e=>e.id===r.empId)?.dept).filter(Boolean)).size;
  const bestAvgOf = rs => {
    const byDept = {};
    rs.forEach(r=>{ const d=employees.find(e=>e.id===r.empId)?.dept; if(d) (byDept[d]=byDept[d]||[]).push(r.score); });
    const avgs = Object.values(byDept).map(ss=>Math.round(ss.reduce((a,s)=>a+s,0)/ss.length));
    return avgs.length?Math.max(...avgs):0;
  };

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
      ['BÁO CÁO KẾT QUẢ THI - Bệnh viện quân y 4'],
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

    XLSX.writeFile(wb, `Báo cáo_${new Date().toLocaleDateString('vi-VN').replace(/\//g,'-')}.xlsx`);
  };

  return (
    <div>
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-5 gap-3">
        <div className="flex items-center gap-3">
          <div><h1 className="text-xl md:text-2xl font-bold text-slate-800">Báo cáo & Phân tích</h1><p className="text-slate-500 text-xs md:text-sm">Phân tích kết quả thi theo đơn vị</p></div>
        </div>
        <button onClick={exportExcel} className="flex items-center gap-2 bg-gradient-to-r from-[#0B4F32] to-emerald-600 hover:from-[#0a4429] hover:to-emerald-700 text-white px-5 py-3 rounded-xl text-sm font-semibold shadow-lg shadow-emerald-900/20 transition-all">
          <Download size={16}/>Xuất Excel <ChevronRight size={14} className="rotate-90 opacity-70"/>
        </button>
      </div>

      {/* Stat cards */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-5">
        {[
          {val:depts.length, lb:'Đơn vị', sub:'Đã có dữ liệu',
           valCol:'text-teal-600', footBg:'bg-teal-50',
           trend:pctChange(deptCountOf(rsThis), deptCountOf(rsPrev))},
          {val:totalAttempts, lb:'Tổng lượt thi', sub:'Trong kỳ',
           valCol:'text-emerald-600', footBg:'bg-emerald-50',
           trend:pctChange(rsThis.length, rsPrev.length)},
          {val:`${globalAvg}%`, lb:'Điểm TB chung', sub:'Toàn hệ thống',
           valCol:'text-amber-500', footBg:'bg-amber-50',
           trend:pctChange(avgOf(rsThis), avgOf(rsPrev))},
          {val:`${bestDept.avg}%`, lb:'Đơn vị cao nhất', sub:bestDept.name,
           valCol:'text-violet-600', footBg:'bg-violet-50',
           trend:pctChange(bestAvgOf(rsThis), bestAvgOf(rsPrev))},
        ].map((s,i)=>{
          const down = s.trend < 0;
          return (
            <div key={i} className="bg-white rounded-2xl p-4 shadow-sm border border-slate-100">
              <div className="flex items-start gap-3">
                <div className="min-w-0">
                  <div className="flex items-baseline gap-2">
                    <span className={`text-2xl font-bold ${s.valCol} leading-tight flex-shrink-0`}>{s.val}</span>
                    <span className="text-xs font-semibold text-slate-600 truncate">{s.lb}</span>
                  </div>
                  <div className="text-xs text-slate-400 truncate">{s.sub}</div>
                </div>
              </div>
              <div className={`mt-3 ${s.footBg} rounded-xl px-3 py-2 flex items-center justify-between gap-2`}>
                <span className={`flex items-center gap-1 text-xs font-bold flex-shrink-0 ${down?'text-red-500':s.valCol}`}>
                  {down ? <TrendingDown size={13}/> : <TrendingUp size={13}/>}{Math.abs(s.trend)}%
                </span>
                <span className="text-xs text-slate-400 truncate">so với kỳ trước</span>
              </div>
            </div>
          );
        })}
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
          {barData.length===0 ? (
            <ChartPlaceholder
              cols={[]}
              icon={<PieChartIcon size={22}/>}
              title="Chưa có dữ liệu để hiển thị"
              sub="Dữ liệu sẽ được cập nhật khi có kết quả thi"/>
          ) : (
            <>
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
            </>
          )}
        </div>

        {/* Donut chart */}
        <div className="md:lg:col-span-2 bg-white rounded-2xl p-5 shadow-sm border border-slate-100">
          <h3 className="text-sm font-semibold text-slate-700 mb-3">Phân bố kết quả</h3>
          <div className="flex items-center gap-4">
            <div className="relative flex-shrink-0 w-[180px] h-[180px] flex items-center justify-center">
              {totalAttempts===0 ? (
                <div className="w-[164px] h-[164px] rounded-full border-[22px] border-slate-100"/>
              ) : (
                <PieChart width={180} height={180}>
                  <Pie data={pieData} cx="50%" cy="50%" innerRadius={58} outerRadius={82} dataKey="value" startAngle={90} endAngle={-270}>
                    {pieData.map((d,i)=><Cell key={i} fill={d.c}/>)}
                  </Pie>
                  <Tooltip formatter={(v,n)=>[v+' lượt',n]}/>
                </PieChart>
              )}
              <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                <span className="text-2xl font-bold text-slate-800">{totalAttempts}</span>
                <span className="text-xs text-slate-400">Tổng lượt thi</span>
              </div>
            </div>
            <div className="flex-1 space-y-3.5">
              {pieData.map(d=>(
                <div key={d.name} className="flex items-center gap-2">
                  <div className="w-2.5 h-2.5 rounded-full flex-shrink-0 mt-0.5" style={{background:d.c}}/>
                  <div className="flex-1 min-w-0">
                    <div className="text-xs font-medium text-slate-700 truncate">{d.name}</div>
                    <div className="text-xs text-slate-400">{d.value} lượt</div>
                  </div>
                  <span className="text-sm font-bold text-slate-700 flex-shrink-0">{totalAttempts?Math.round(d.value/totalAttempts*100):0}%</span>
                </div>
              ))}
            </div>
          </div>
          <div className="mt-5 flex items-center justify-center gap-2 bg-emerald-50 border border-emerald-100 rounded-xl py-2.5 px-3 text-xs text-emerald-700">
            <Info size={14} className="flex-shrink-0"/> Dữ liệu được cập nhật theo thời gian thực
          </div>
        </div>
      </div>

      {/* Detail table */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden mb-5">
        <div className="px-5 py-4 border-b border-slate-100">
          <h3 className="text-sm font-semibold text-slate-700">Chi tiết kết quả theo đơn vị</h3>
        </div>
        <div className="overflow-x-auto">
        <table className="w-full min-w-[1000px]">
          <thead>
            <tr className="border-b border-slate-100 bg-slate-50/50">
              {['STT','Đơn vị','Lượt thi','Đạt (≥ 70%)','Trung bình (50–69%)','Chưa đạt (< 50%)','Điểm TB','Tiến độ đạt','Xu hướng'].map(h=>(
                <th key={h} className="px-5 py-3 text-left text-xs font-bold text-slate-600 uppercase tracking-wide whitespace-nowrap">{h}</th>
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
                  <td className="px-5 py-4 text-sm text-slate-500">{i+1}</td>
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
                  <td className="px-5 py-4 text-sm font-semibold text-amber-500">{d.mid}</td>
                  <td className="px-5 py-4 text-sm font-semibold text-red-500">{d.low}</td>
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
            {deptData.length===0 && (
              <tr><td colSpan={9} className="px-5 py-6">
                <EmptyState icon={<FileSearch size={22}/>} title="Chưa có dữ liệu" sub="Thông tin chi tiết sẽ hiển thị tại đây"/>
              </td></tr>
            )}
          </tbody>
        </table>
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
            <div key={exam.id} className={`rounded-xl p-4 shadow-sm border ${exam.locked?'bg-slate-50 border-slate-200':'bg-white border-slate-100'}`}>
              <div className="flex items-start gap-3">
                <div className={`w-10 h-10 rounded-xl flex items-center justify-center flex-shrink-0 bg-slate-100 text-slate-400`}><FileText size={18}/></div>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 mb-1 flex-wrap">
                    <h3 className={`font-semibold text-sm ${exam.locked?'text-slate-500':'text-slate-800'}`}>{exam.title}</h3>
                    {exam.locked && (
                      <span className="flex items-center gap-1 px-2 py-0.5 bg-amber-50 border border-amber-200 rounded-full text-xs font-medium text-amber-700">
                        <Lock size={10}/>Đã khoá
                      </span>
                    )}
                    {ok&&!exam.locked&&<span className="px-2 py-0.5 bg-emerald-100 text-emerald-700 rounded-full text-xs">✓ Đạt</span>}
                  </div>
                  <p className={`text-xs mb-2 line-clamp-1 ${exam.locked?'text-slate-400':'text-slate-500'}`}>{exam.desc}</p>
                  <div className={`flex flex-wrap gap-2 text-xs ${exam.locked?'text-slate-400':'text-slate-400'}`}>
                    <span className="flex items-center gap-1"><FileText size={10}/>{exam.qIds.length} câu</span>
                    <span className="flex items-center gap-1"><Clock size={10}/>{exam.time} phút</span>
                    <span className="flex items-center gap-1"><Award size={10}/>Đạt: {exam.pass}%</span>
                    {last&&<span className={exam.locked?'text-slate-400':'text-slate-500'}>Điểm: {last.score}%</span>}
                  </div>
                </div>
                <button disabled={exam.locked||!!last} onClick={()=>onStart(exam)} className={`flex items-center gap-1 px-3 py-2 rounded-lg text-xs flex-shrink-0 ${exam.locked||last?'bg-slate-200 text-slate-400 cursor-not-allowed':'bg-emerald-600 text-white hover:bg-emerald-700'}`} title={exam.locked?'Đề thi đã bị khoá':last?'Bạn đã thi đề này, mỗi thí sinh chỉ được thi 1 lần':'Bắt đầu thi'}>
                  {last?<CheckCircle size={11}/>:<Play size={11}/>}{last?'Đã thi':'Bắt đầu'}
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
const MyResults = ({user, exams, questions, results}) => {
  const [openId, setOpenId] = useState(null);
  const fmtTime = (t) => t!=null ? `${Math.floor(t/60)}p${String(t%60).padStart(2,'0')}s` : '--';
  const my =results.filter(r=>r.empId===user.id);

  const attempt = my.find(r=>r.id===openId);
  if (attempt) {
    const exam = exams.find(e=>e.id===attempt.examId);
    const hasSnapshot = Array.isArray(attempt.questionSnapshot);
    const reviewQuestions = hasSnapshot ? attempt.questionSnapshot : (exam?.qIds || []).map(id=>(questions||[]).find(q=>q.id===id));
    const answers = Array.isArray(attempt.answers) ? attempt.answers : [];
    // Bài cũ chỉ lưu vị trí đáp án; thiếu câu hỏi thì không thể đối chiếu chính xác.
    const canMap = hasSnapshot || (reviewQuestions.every(Boolean) && reviewQuestions.length===answers.length);
    return (
      <div>
        <button onClick={()=>setOpenId(null)} className="flex items-center gap-1 text-sm text-slate-500 hover:text-emerald-600 mb-3"><ArrowLeft size={15}/>Quay lại kết quả</button>
        <h1 className="text-lg md:text-xl font-bold text-slate-800">{exam?.title || 'Đề thi đã xóa'}</h1>
        <p className="text-slate-500 text-sm mt-1 mb-4">{attempt.date} • {fmtTime(attempt.timeTaken)} • {attempt.score}% • {attempt.correct} câu đúng</p>

        {!canMap || answers.length===0 ? (
          <p className="bg-white border border-slate-200 rounded-xl p-5 text-slate-600">Không đủ dữ liệu để hiển thị bài làm của bạn.</p>
        ) : (
          <div className="space-y-4">
            {reviewQuestions.map((q,index)=>{
              const selected = answers[index];
              const answered = Number.isInteger(selected) && selected>=0 && selected<q.opts.length;
              const correct = answered && selected===q.ans;
              return (
                <section key={index} className="bg-white border border-slate-200 rounded-xl p-4 sm:p-5">
                  <div className="flex items-center justify-between gap-3 mb-3">
                    <h2 className="font-semibold text-slate-800">Câu {index+1}</h2>
                    <span className={`text-xs font-medium ${correct?'text-emerald-700':answered?'text-red-600':'text-slate-500'}`}>{correct?'Đúng':answered?'Sai':'Chưa trả lời'}</span>
                  </div>
                  <p className="text-slate-800 whitespace-pre-wrap mb-3">{q.text}</p>
                  <div className="space-y-2">
                    {q.opts.map((option,i)=>(
                      <div key={i} className={`border rounded-lg p-3 text-sm whitespace-pre-wrap ${i===selected?(correct?'bg-emerald-50 border-emerald-300 text-emerald-900':'bg-red-50 border-red-300 text-red-800'):'border-slate-200 text-slate-600'}`}>
                        <span className="font-semibold">{OPT_LETTERS[i]}. </span>{option}
                        {i===selected && <span className="block text-xs font-semibold mt-1">Bạn đã chọn</span>}
                      </div>
                    ))}
                  </div>
                </section>
              );
            })}
          </div>
        )}
      </div>
    );
  }

  return (
    <div>
      <div className="mb-5"><h1 className="text-xl font-bold text-slate-800">Kết quả của tôi</h1><p className="text-slate-500 text-sm">Lịch sử thi và điểm số</p></div>
      {my.length===0?(
        <div className="text-center py-20 text-slate-400"><Award size={40} className="mx-auto mb-3 opacity-20"/><p>Bạn chưa tham gia kỳ thi nào</p></div>
      ):(
        <div className="space-y-3">
          {my.map(r=>{const exam=exams.find(e=>e.id===r.examId);const ok=exam&&r.score>=exam.pass;return(
            <div key={r.id} onClick={()=>setOpenId(r.id)} className="bg-white rounded-xl p-5 shadow-sm border border-slate-100 flex items-center gap-4 cursor-pointer">
              <div className={`w-14 h-14 rounded-xl flex items-center justify-center flex-shrink-0 bg-slate-100`}>
                <span className="text-lg font-bold text-slate-700">{r.score}%</span>
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

// ── SESSION CODE MODAL ──
const SessionCodeModal = ({exam, onConfirm, onCancel}) => {
  const [code, setCode] = useState('');
  const [error, setError] = useState('');
  const isValid = exam.sessions && exam.sessions.some(s => s.code === code.trim().toUpperCase());
  const handleSubmit = () => {
    if (!code.trim()) return setError('Vui lòng nhập mã kíp');
    if (!isValid) return setError('Mã kíp không đúng. Vui lòng kiểm tra lại');
    onConfirm();
  };
  return (
    <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-2xl w-full max-w-sm shadow-2xl">
        <div className="px-6 py-4 border-b">
          <h2 className="font-bold text-slate-800 text-lg">Nhập mã kíp thi</h2>
          <p className="text-xs text-slate-500 mt-1">Để bắt đầu thi, bạn cần nhập mã kíp được cấp</p>
        </div>
        <div className="p-6 space-y-4">
          <div>
            <label className="text-xs font-medium text-slate-600 mb-2 block">Mã kíp thi</label>
            <input
              autoFocus
              type="text"
              className={`w-full border-2 rounded-xl px-4 py-3 text-lg font-mono uppercase tracking-widest text-center transition-all ${
                error ? 'border-red-300 bg-red-50' : 'border-slate-200 bg-slate-50 focus:border-emerald-500 focus:bg-white'
              }`}
              placeholder="VD: A1B2C3"
              value={code}
              onChange={(e) => { setCode(e.target.value); setError(''); }}
              onKeyDown={(e) => e.key === 'Enter' && handleSubmit()}
            />
            {error && <p className="text-xs text-red-600 mt-2 flex items-center gap-1"><AlertCircle size={12}/>{error}</p>}
          </div>
          <div className="flex gap-3 pt-2">
            <button onClick={onCancel} className="flex-1 py-2.5 border border-slate-200 rounded-lg text-sm text-slate-600 hover:bg-slate-50">Hủy</button>
            <button onClick={handleSubmit} disabled={!code.trim()} className="flex-1 py-2.5 bg-emerald-600 text-white rounded-lg text-sm font-medium hover:bg-emerald-700 disabled:opacity-50">Vào thi</button>
          </div>
        </div>
      </div>
    </div>
  );
};

// ── EXAM SCREEN ──
const ExamScreen = ({user, exam, questions, onFinish}) => {
  const [qs] = useState(() => exam.qIds.map(id => questions.find(q => q.id === id)).filter(Boolean));
  const [cur, setCur] = useState(0);
  const [ans, setAns] = useState(Array(qs.length).fill(-1));
  const [tLeft, setTLeft] = useState(exam.time*60);
  const [confirm, setConfirm] = useState(false);
  const ansRef = useRef(ans);
  const doneRef = useRef(false);

  useEffect(()=>{ansRef.current=ans;},[ans]);

  // Đổi câu (Trước / Tiếp / bảng câu hỏi) thì cuộn lên đầu để đọc từ đầu câu hỏi
  useEffect(()=>{ window.scrollTo({top:0, behavior:'smooth'}); },[cur]);

  const submit = (a) => {
    if(doneRef.current) return; doneRef.current=true;
    const correct = a.filter((x,i)=>x===qs[i].ans).length;
    const score = Math.round(correct/qs.length*100);
    const timeTaken = exam.time*60 - tLeft; // seconds used
    onFinish({id:Date.now(),empId:user.id,examId:exam.id,score,correct,timeTaken,date:new Date().toLocaleDateString('vi-VN'),leaves:leaveCount.current,answers:[...a],questionSnapshot:qs.map(q=>({id:q.id,text:q.text,opts:[...q.opts],ans:q.ans}))});
  };

  // Rời khỏi màn hình thi: lần 1 cảnh báo, lần 2 nộp bài luôn
  const submitRef = useRef(submit);
  submitRef.current = submit;
  const leaveCount = useRef(0);
  const awayRef = useRef(false);
  const [warned, setWarned] = useState(false);
  const [forcedEnd, setForcedEnd] = useState(false);

  useEffect(()=>{
    let lastLeaveAt = 0;
    const onLeave = () => {
      if(awayRef.current || doneRef.current) return; // blur + visibilitychange có thể bắn cùng lúc
      awayRef.current = true;
      lastLeaveAt = Date.now();
      leaveCount.current += 1;
      if(leaveCount.current >= MAX_LEAVES){ setForcedEnd(true); submitRef.current(ansRef.current); }
      else setWarned(true);
    };
    const onBack = () => { awayRef.current = false; };
    const onVis = () => { document.hidden ? onLeave() : onBack(); };
    document.addEventListener('visibilitychange', onVis);
    window.addEventListener('blur', onLeave);
    window.addEventListener('focus', onBack);
    // Trên điện thoại (Zalo, Safari iOS...) mở trình chuyển ứng dụng / thanh thông báo
    // có thể không bắn blur/visibilitychange: dùng thêm pagehide/pageshow
    window.addEventListener('pagehide', onLeave);
    window.addEventListener('pageshow', onBack);
    // Dự phòng: trang bị treo khi ra khỏi ứng dụng nên đồng hồ nhảy cóc một khoảng dài
    let lastTick = Date.now();
    const gap = setInterval(()=>{
      const now = Date.now();
      const suspended = now - lastTick > 3000;
      lastTick = now;
      if(suspended && lastLeaveAt < now - 3000){ awayRef.current = false; onLeave(); }
    },1000);
    return () => {
      clearInterval(gap);
      document.removeEventListener('visibilitychange', onVis);
      window.removeEventListener('blur', onLeave);
      window.removeEventListener('focus', onBack);
      window.removeEventListener('pagehide', onLeave);
      window.removeEventListener('pageshow', onBack);
    };
  },[]);

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
      {warned && !forcedEnd && (
        <div className="fixed inset-0 bg-black/70 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl w-full max-w-sm shadow-2xl p-6 text-center">
            <AlertCircle size={40} className="mx-auto mb-3 text-red-500"/>
            <h3 className="font-bold text-slate-800 text-lg mb-2">Cảnh báo!</h3>
            <p className="text-sm text-slate-600 mb-5">Bạn đã rời khỏi màn hình làm bài <b>{leaveCount.current}/{MAX_LEAVES}</b> lần. Nếu rời đi đến lần thứ <b>{MAX_LEAVES}</b>, bài thi sẽ bị <b>kết thúc và nộp ngay</b> (còn {MAX_LEAVES-leaveCount.current} lần cảnh báo).</p>
            <button onClick={()=>setWarned(false)} className="w-full py-2.5 bg-emerald-600 text-white rounded-xl text-sm hover:bg-emerald-700">Tôi đã hiểu, tiếp tục làm bài</button>
          </div>
        </div>
      )}
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
      <div className="flex flex-col md:flex-row flex-1 gap-4 p-3 md:p-6">
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
        <div className="w-full md:w-48 md:flex-shrink-0">
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
// Nộp bài xong chỉ báo ngắn gọn số câu đúng — không hiển thị lại đáp án đúng
// để tránh lộ nội dung ngân hàng câu hỏi cho các lượt thi sau.
const ResultScreen = ({result, exam, questions, onBack}) => {
  const total = exam.qIds.filter(id => questions.some(q => q.id === id)).length;
  const unanswered = result.answers.filter(x => x === -1).length;
  const wrong = total - result.correct - unanswered;
  const ok = result.score>=exam.pass;

  // Đạt yêu cầu thì bắn pháo giấy chúc mừng trong ~3 giây
  useEffect(()=>{
    if(!ok) return;
    const end = Date.now() + 3000;
    const colors = ['#10b981','#fbbf24','#ef4444','#3b82f6','#a855f7'];
    let raf;
    const frame = () => {
      confetti({particleCount:4, angle:60, spread:60, origin:{x:0, y:0.7}, colors});
      confetti({particleCount:4, angle:120, spread:60, origin:{x:1, y:0.7}, colors});
      if(Date.now() < end) raf = requestAnimationFrame(frame);
    };
    confetti({particleCount:120, spread:90, origin:{y:0.6}, colors});
    frame();
    return () => { cancelAnimationFrame(raf); confetti.reset(); };
  },[]);

  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
      <div className="w-full max-w-md">
        <div className={`rounded-2xl px-6 py-8 text-center text-white font-medium mb-4 ${ok?'bg-emerald-500':'bg-red-500'}`}>
          <div className="text-3xl font-bold mb-2">{result.score}%</div>
          <div className="text-sm font-normal">Cần đạt {exam.pass}%</div>
        </div>
        <div className="bg-white rounded-2xl p-6 mb-4 shadow-sm border border-slate-100">
          <div className="grid grid-cols-3 gap-3">
            <div className="text-center">
              <div className="text-2xl font-bold text-emerald-600">{result.correct}</div>
              <div className="text-xs text-slate-500 mt-1">Câu đúng</div>
            </div>
            <div className="text-center">
              <div className="text-2xl font-bold text-red-500">{wrong}</div>
              <div className="text-xs text-slate-500 mt-1">Câu sai</div>
            </div>
            <div className="text-center">
              <div className="text-2xl font-bold text-amber-500">{unanswered}</div>
              <div className="text-xs text-slate-500 mt-1">Chưa làm</div>
            </div>
          </div>
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

// Nhánh nguyệt quế trái: lá và cành bám theo cùng một cung tròn
const LAUREL = (() => {
  const CX = 60, CY = 30, R = 38;
  const at = (deg, r=R) => {
    const t = deg*Math.PI/180;
    return [CX + r*Math.cos(t), CY + r*Math.sin(t)];
  };
  const stem = Array.from({length:20}, (_,i)=>at(97 + i*(90/19)))
    .map(([x,y],i)=>`${i?'L':'M'}${x.toFixed(1)},${y.toFixed(1)}`).join(' ');
  const leaves = Array.from({length:9}, (_,i)=>{
    const deg = 101 + i*10.5;                 // 101° (đáy) → 185° (mũi bên trái)
    const [x,y] = at(deg, R + 3.5);
    return {x, y, rot: deg + 90, s: 1 - i*0.045};
  });
  return {stem, leaves};
})();

// Bản vẽ SVG dự phòng, dùng khi chưa có file ảnh trong public/
const EmblemSvg = ({size=100}) => {
  const branch = (
    <g>
      <path d={LAUREL.stem} fill="none" stroke="#15803d" strokeWidth="2.4" strokeLinecap="round"/>
      {LAUREL.leaves.map((l,i)=>(
        <ellipse key={i} cx={l.x} cy={l.y} rx={7.5*l.s} ry={3.4*l.s} fill={i%2?'#15803d':'#12693a'}
                 transform={`rotate(${l.rot} ${l.x} ${l.y})`}/>
      ))}
    </g>
  );
  return (
    <svg width={size} height={size*0.63} viewBox="0 0 120 76" className="drop-shadow-sm">
      <defs>
        <linearGradient id="em-shield" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor="#d33b2c"/><stop offset="100%" stopColor="#a92418"/>
        </linearGradient>
      </defs>
      {branch}
      <g transform="translate(120,0) scale(-1,1)">{branch}</g>
      <path d="M39 6 H81 A3 3 0 0 1 84 9 V38 C84 53 74 64 60 70 C46 64 36 53 36 38 V9 A3 3 0 0 1 39 6 Z"
            fill="url(#em-shield)" stroke="#8f2016" strokeWidth="3" strokeLinejoin="round"/>
      <polygon points={starPts(60,34,16,6.6)} fill="#f7c744"/>
    </svg>
  );
};

// Quốc huy: ưu tiên file ảnh trong public/, không có thì tự vẽ bằng SVG.
// Đặt file của bạn vào public/ với một trong các tên dưới đây (hoặc sửa lại danh sách).
const EMBLEM_SRCS = ['/bvqy4.png'];
const Emblem = ({size=100, className=''}) => {
  const [tried, setTried] = useState(0);
  if (tried >= EMBLEM_SRCS.length) return <EmblemSvg size={size}/>;
  return (
    <img
      key={tried}
      src={EMBLEM_SRCS[tried]}
      onError={()=>setTried(t=>t+1)}
      alt="Quốc huy"
      width={size}
      draggable={false}
      className={`h-auto select-none ${className}`}
    />
  );
};

// Nền: sóng xanh + watermark trụ sở và ngôi sao
const Login = ({onLogin, externalErr}) => {
  const [account, setAccount] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [err, setErr] = useState('');
  const [remember, setRemember] = useState(true);
  const [busy, setBusy] = useState(false);

  const handleLogin = async (event) => {
    event.preventDefault();
    const acc = account.trim();
    if (!acc) { setErr('Vui lòng nhập tài khoản.'); return; }
    if (!password) { setErr('Vui lòng nhập mật khẩu.'); return; }
    setBusy(true);
    try {
      await onLogin(acc, password, remember);   // thành công → App tự chuyển màn hình
    } catch (ex) {
      setErr(authErrorMessage(ex));
    }
    setBusy(false);
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-[#fafafa] p-4">
      <div className="w-full max-w-sm">
        <h1 className="font-serif text-3xl font-bold text-slate-900">Đăng nhập</h1>
        <p className="text-slate-500 text-sm mt-1.5">Sử dụng tài khoản khoa/phòng do quản trị cấp.</p>

        <form onSubmit={handleLogin} className="mt-6 space-y-4" noValidate autoComplete="off">
          <div>
            <label htmlFor="login-account" className="block text-sm text-slate-700 mb-1.5">Tên đăng nhập</label>
            <div className="relative">
              <User size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"/>
              <input
                type="text"
                id="login-account"
                name="hvqy4-account"
                autoComplete="off"
                autoFocus
                value={account}
                onChange={e=>{setAccount(e.target.value); setErr('');}}
                className="w-full pl-9 pr-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-slate-800 focus:outline-none focus:bg-white focus:border-[#0B4F32] focus:ring-2 focus:ring-emerald-500/20 transition"
              />
            </div>
          </div>

          <div>
            <label htmlFor="login-password" className="block text-sm text-slate-700 mb-1.5">Mật khẩu</label>
            <div className="relative">
              <Lock size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"/>
              <input
                type={showPassword ? 'text' : 'password'}
                id="login-password"
                name="hvqy4-password"
                autoComplete="new-password"
                value={password}
                onChange={e=>{setPassword(e.target.value); setErr('');}}
                className="w-full pl-9 pr-10 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-slate-800 focus:outline-none focus:bg-white focus:border-[#0B4F32] focus:ring-2 focus:ring-emerald-500/20 transition"
              />
              <button
                type="button"
                onClick={()=>setShowPassword(s=>!s)}
                aria-label={showPassword ? 'Ẩn mật khẩu' : 'Hiện mật khẩu'}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
              >
                {showPassword ? <EyeOff size={16}/> : <Eye size={16}/>}
              </button>
            </div>
          </div>

          <div className="flex items-center justify-between text-sm">
            <label className="flex items-center gap-2 text-slate-600 cursor-pointer">
              <input type="checkbox" checked={remember} onChange={e=>setRemember(e.target.checked)} className="w-4 h-4 rounded border-slate-300 accent-[#0B4F32]"/>
              Ghi nhớ đăng nhập
            </label>
            <button
              type="button"
              onClick={()=>setErr('Vui lòng liên hệ quản trị viên để được cấp lại mật khẩu.')}
              className="text-emerald-700 hover:underline"
            >
              Quên mật khẩu?
            </button>
          </div>

          {(err || externalErr) && <p role="alert" className="text-red-600 text-sm">{err || externalErr}</p>}

          <button type="submit" disabled={busy} className="w-full bg-[#0B4F32] hover:bg-[#0a4429] disabled:opacity-60 text-white py-2.5 rounded-lg font-semibold transition-colors">
            {busy ? 'Đang đăng nhập...' : 'Đăng nhập'}
          </button>
        </form>

        <div className="mt-6 flex justify-center">
          <div className="group relative inline-block text-center text-slate-400 text-xs cursor-default">
            © 2026 Ban công nghệ thông tin · Bệnh viện Quân y 4
            <span className="pointer-events-none absolute left-1/2 bottom-full mb-2 -translate-x-1/2 whitespace-nowrap rounded-md bg-slate-800 px-3 py-1.5 text-xs text-white opacity-0 shadow-lg transition-opacity duration-150 group-hover:opacity-100">
              Designed, built, and maintained by Bui Duc Thang
            </span>
          </div>
        </div>
      </div>
    </div>
  );
};

// ── MAIN APP ──
// ── EXAM RESULTS (admin): pick an exam → list every attempt for it ──
const ExamResults = ({results, exams, employees, questions, onClearAll}) => {
  const [selId, setSelId] = useState(null);
  const [attemptId, setAttemptId] = useState(null);
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
          <div><h1 className="text-lg md:text-xl font-bold text-slate-700">Kết quả thi</h1><p className="text-slate-500 text-xs md:text-sm mt-0.5">Chọn một đề thi để xem danh sách người thi và kết quả</p></div>
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

  const attempt = rows.find(r=>r.id===attemptId);
  if (attempt) {
    const emp = employees.find(e=>e.id===attempt.empId);
    const hasSnapshot = Array.isArray(attempt.questionSnapshot);
    const reviewQuestions = hasSnapshot ? attempt.questionSnapshot : (exam?.qIds || []).map(id=>questions.find(q=>q.id===id));
    const answers = Array.isArray(attempt.answers) ? attempt.answers : [];
    // Legacy submissions stored only answer positions. Missing questions make that mapping ambiguous.
    const canMap = hasSnapshot || (reviewQuestions.every(Boolean) && reviewQuestions.length===answers.length);
    return (
      <div>
        <button onClick={()=>setAttemptId(null)} className="flex items-center gap-1 text-sm text-slate-500 hover:text-emerald-600 mb-3"><ArrowLeft size={15}/>Quay lại danh sách kết quả</button>
        <h1 className="text-lg md:text-xl font-bold text-slate-800">Bài thi của {emp?.name || 'Thí sinh đã xóa'}</h1>
        <p className="text-slate-500 text-sm mt-1 mb-4">{exam?.title || 'Đề thi đã xóa'} • {attempt.date} • {fmtTime(attempt.timeTaken)} • {attempt.score}% • {attempt.correct} câu đúng</p>

        {!canMap || answers.length===0 ? (
          <p className="bg-white border border-slate-200 rounded-xl p-5 text-slate-600">Không đủ dữ liệu để đối chiếu từng câu trả lời của bài thi này.</p>
        ) : (
          <div className="space-y-4">
            {reviewQuestions.map((q,index)=>{
              const selected = answers[index];
              const answered = Number.isInteger(selected) && selected>=0 && selected<q.opts.length;
              const correct = answered && selected===q.ans;
              return (
                <section key={index} className="bg-white border border-slate-200 rounded-xl p-4 sm:p-5">
                  <div className="flex items-center justify-between gap-3 mb-3">
                    <h2 className="font-semibold text-slate-800">Câu {index+1}</h2>
                    <span className={`text-xs font-medium ${correct?'text-emerald-700':answered?'text-red-600':'text-slate-500'}`}>{correct?'Đúng':answered?'Sai':selected===-1?'Chưa trả lời':'Không có dữ liệu trả lời'}</span>
                  </div>
                  <p className="text-slate-800 whitespace-pre-wrap mb-3">{q.text}</p>
                  <div className="space-y-2">
                    {q.opts.map((option,i)=>(
                      <div key={i} className={`border rounded-lg p-3 text-sm whitespace-pre-wrap ${i===q.ans?'bg-emerald-50 border-emerald-300 text-emerald-900':i===selected?'bg-red-50 border-red-300 text-red-800':'border-slate-200 text-slate-600'}`}>
                        <span className="font-semibold">{OPT_LETTERS[i]}. </span>{option}
                        {(i===selected || i===q.ans) && <span className="block text-xs font-semibold mt-1">{[i===selected?'Thí sinh đã chọn':null,i===q.ans?'Đáp án đúng':null].filter(Boolean).join(' • ')}</span>}
                      </div>
                    ))}
                  </div>
                </section>
              );
            })}
          </div>
        )}
      </div>
    );
  }

  const exportExcel = () => {
    const head = ['Họ và tên','Đơn vị','Điểm (%)','Số câu đúng','Tổng câu','Kết quả','Thời gian làm bài','Ngày thi','Số lần rời màn hình'];
    const body = rows.map(r=>{
      const emp = employees.find(e=>e.id===r.empId);
      const ok = exam && r.score >= exam.pass;
      return [emp?.name||'(Đã xóa)', emp?.dept||'', r.score, r.correct, totalQ, ok?'Đạt':'Không đạt', fmtTime(r.timeTaken), r.date, r.leaves ?? '--'];
    });
    const ws = XLSX.utils.aoa_to_sheet([
      ['KẾT QUẢ THI: '+(exam?.title||'')],
      ['Xuất ngày: '+new Date().toLocaleDateString('vi-VN')],
      [],
      head, ...body,
    ]);
    ws['!cols']=[{wch:22},{wch:18},{wch:10},{wch:12},{wch:10},{wch:14},{wch:16},{wch:14},{wch:20}];
    ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:8}}];
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
            <thead><tr className="bg-slate-50">{['Thí sinh','Đơn vị','Điểm','Số câu đúng','Kết quả','Thời gian','Ngày thi','Rời màn hình'].map(h=><th key={h} className={`px-3 py-2 text-xs font-medium text-slate-400 uppercase ${h==='Kết quả'?'text-center':'text-left'}`}>{h}</th>)}</tr></thead>
            <tbody>
              {rows.map(r=>{
                const emp=employees.find(e=>e.id===r.empId);
                const ok=exam&&r.score>=exam.pass;
                return (
                  <tr key={r.id} className="border-t border-slate-50 hover:bg-slate-50/50">
                    <td className="px-3 py-2 text-xs font-medium"><button onClick={()=>setAttemptId(r.id)} className="text-emerald-700 hover:text-emerald-900 hover:underline text-left"><span className="block">{emp?.name || '(Đã xóa)'}</span><span className="text-xs font-normal">Xem bài làm</span></button></td>
                    <td className="px-3 py-2 text-xs text-slate-600">{emp?.dept||'--'}</td>
                    <td className="px-3 py-2 text-xs font-bold text-slate-800">{r.score}%</td>
                    <td className="px-3 py-2 text-xs text-slate-600">{r.correct}/{totalQ}</td>
                    <td className="px-3 py-2 text-center"><Badge ok={ok}/></td>
                    <td className="px-3 py-2 text-xs text-slate-500">{fmtTime(r.timeTaken)}</td>
                    <td className="px-3 py-2 text-xs text-slate-400">{r.date}</td>
                    <td className="px-3 py-2 text-xs">
                      {r.leaves==null ? <span className="text-slate-300">--</span>
                        : r.leaves===0 ? <span className="text-slate-500">0</span>
                        : <span className={`px-2 py-0.5 rounded-full font-medium ${r.leaves>=MAX_LEAVES?'bg-red-100 text-red-700':'bg-amber-100 text-amber-700'}`}>{r.leaves>=MAX_LEAVES?`${r.leaves} • Bị nộp bài`:r.leaves}</span>}
                    </td>
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

// ── ACCOUNTS (Quản lý tài khoản) ──
const Accounts = ({employees, setEmployees, authUsers, onProvision}) => {
  const [modal, setModal] = useState(null);
  const [form, setForm] = useState({name:'', oldAccount:'', newAccount:''});
  const [formErr, setFormErr] = useState('');
  const [creds, setCreds] = useState(null);
  const [busy, setBusy] = useState(false);

  const hasLogin = emp => authUsers.some(u => u.role === 'employee' && String(u.empId) === String(emp.id) && u.account === (emp.account||''));
  const missing = employees.filter(e => (e.account||'').trim() && !hasLogin(e));

  const runProvision = async (emps, replace=false) => {
    setBusy(true);
    try {
      const r = await onProvision(emps, replace);
      if (r.creds.length || r.fails.length) setCreds(r);
    } finally { setBusy(false); }
  };

  const openEdit = (emp) => {
    setForm({name:emp.name, oldAccount:emp.account||'', newAccount:emp.account||''});
    setFormErr('');
    setModal({mode:'edit', emp});
  };

  const handleSave = () => {
    const newAccount = form.newAccount.trim();
    if(!newAccount) return setFormErr('Tài khoản không được để trống');
    const dup = employees.find(e=>e.account===newAccount && e.id!==modal.emp.id);
    if(dup) return setFormErr('Tài khoản này đã được sử dụng');
    if(newAccount === (modal.emp.account||'')) return setFormErr('Tài khoản mới phải khác tài khoản cũ');
    setEmployees(p=>p.map(e=>e.id===modal.emp.id?{...e,account:newAccount}:e));
    // Tên đăng nhập đổi → cấp tài khoản mới (mật khẩu mới), thu hồi tài khoản cũ
    runProvision([{...modal.emp, account:newAccount}], true);
    setModal(null);
  };

  return (
    <div>
      {creds && <CredentialsModal data={creds} onClose={()=>setCreds(null)}/>}
      {modal && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl w-full max-w-md shadow-2xl">
            <div className="flex items-center justify-between p-5 border-b">
              <h2 className="font-bold text-slate-800">Đổi tài khoản</h2>
              <button onClick={()=>setModal(null)} className="text-slate-400 hover:text-slate-600"><X size={18}/></button>
            </div>
            <div className="p-5 space-y-4">
              <div>
                <label className="text-xs font-medium text-slate-600 mb-1.5 block">Thí sinh</label>
                <input className="w-full border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm bg-slate-50" value={form.name} disabled/>
              </div>
              <div>
                <label className="text-xs font-medium text-slate-600 mb-1.5 block">Tài khoản cũ</label>
                <input className="w-full border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm bg-slate-50" value={form.oldAccount} disabled/>
              </div>
              <div>
                <label className="text-xs font-medium text-slate-600 mb-1.5 block">Tài khoản mới <span className="text-red-400">*</span></label>
                <input
                  className="w-full border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm focus:outline-none focus:border-emerald-500"
                  placeholder="Nhập tài khoản mới"
                  value={form.newAccount}
                  onChange={e=>{ setForm(f=>({...f,newAccount:e.target.value})); setFormErr(''); }}
                />
              </div>
              {formErr && (
                <div className="flex items-center gap-2 bg-red-50 border border-red-200 rounded-xl px-3 py-2">
                  <AlertCircle size={14} className="text-red-500 flex-shrink-0"/>
                  <span className="text-red-600 text-xs">{formErr}</span>
                </div>
              )}
              <div className="flex gap-2 pt-1">
                <button onClick={()=>setModal(null)} className="flex-1 py-2.5 border border-slate-200 rounded-xl text-sm text-slate-600 hover:bg-slate-50">Hủy</button>
                <button onClick={handleSave} className="flex-1 py-2.5 bg-emerald-600 text-white rounded-xl text-sm font-medium hover:bg-emerald-700">Cập nhật</button>
              </div>
            </div>
          </div>
        </div>
      )}

      <div className="mb-5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h1 className="text-lg md:text-xl font-bold text-slate-800">Quản lý tài khoản</h1>
          <p className="text-slate-500 text-xs md:text-sm mt-0.5">{employees.length} thí sinh · {missing.length} chưa được cấp tài khoản đăng nhập</p>
        </div>
        {missing.length > 0 && (
          <button disabled={busy} onClick={()=>runProvision(missing)} className="px-4 py-2.5 bg-emerald-600 text-white rounded-xl text-sm font-medium hover:bg-emerald-700 disabled:opacity-60">
            {busy ? 'Đang cấp...' : `Cấp tài khoản cho ${missing.length} thí sinh`}
          </button>
        )}
      </div>

      {employees.length === 0 ? (
        <div className="bg-white rounded-xl p-8 text-center text-slate-400 border border-slate-100">Chưa có thí sinh nào</div>
      ) : (
        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-x-auto">
          <table className="w-full min-w-[640px]">
            <thead>
              <tr className="border-b border-slate-100 bg-slate-50/50">
                {['Thí sinh','Đơn vị','Tài khoản','Đăng nhập','Thao tác'].map(h=>(
                  <th key={h} className="px-5 py-3 text-xs font-semibold text-slate-400 uppercase tracking-wide text-left">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {employees.map(emp=>(
                <tr key={emp.id} className="border-t border-slate-50 hover:bg-slate-50/50 transition-colors">
                  <td className="px-5 py-3 text-sm font-medium text-slate-800">{emp.name}</td>
                  <td className="px-5 py-3 text-xs text-slate-500">{emp.dept}</td>
                  <td className="px-5 py-3 text-sm font-mono text-slate-600">{emp.account||'--'}</td>
                  <td className="px-5 py-3 text-xs">
                    {hasLogin(emp)
                      ? <span className="text-emerald-600 font-medium">Đã cấp</span>
                      : <span className="text-amber-600 font-medium">Chưa cấp</span>}
                  </td>
                  <td className="px-5 py-3">
                    <div className="flex gap-2">
                      <button onClick={()=>openEdit(emp)} className="px-3 py-1.5 bg-emerald-50 text-emerald-600 rounded-lg text-xs font-medium hover:bg-emerald-100">Đổi tài khoản</button>
                      {!hasLogin(emp) && (emp.account||'').trim() && (
                        <button disabled={busy} onClick={()=>runProvision([emp])} className="px-3 py-1.5 bg-slate-100 text-slate-700 rounded-lg text-xs font-medium hover:bg-slate-200 disabled:opacity-60">Cấp tài khoản</button>
                      )}
                    </div>
                  </td>
                </tr>
              ))}
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
  const [rail, setRail] = useState(false);   // sidebar thu gọn
  const [employees, setEmployees] = useState([]);
  const [questions, setQuestions] = useState([]);
  const [exams, setExams] = useState([]);
  const [results, setResults] = useState([]);
  const [authUsers, setAuthUsers] = useState([]);   // collection users (chỉ admin đọc được)
  const [authError, setAuthError] = useState('');
  const [showChangePw, setShowChangePw] = useState(false);
  const [activeExam, setActiveExam] = useState(null);
  const [pendingExam, setPendingExam] = useState(null);
  const [lastResult, setLastResult] = useState(null);
  const [loading, setLoading] = useState(true);
  const [dbError, setDbError] = useState(missingConfig.length
    ? `Thiếu cấu hình Firebase (${missingConfig.join(', ')}). Kiểm tra file .env rồi chạy/build lại — dữ liệu sẽ KHÔNG được lưu.`
    : null);
  const [saving, setSaving] = useState(0);   // số lệnh ghi đang chờ máy chủ xác nhận

  // Bản sao mới nhất của từng collection, dùng để tính diff khi ghi (không đọc từ
  // state trong updater — xem makeSyncSetter).
  const dataRef = useRef({ employees: [], questions: [], exams: [], results: [], users: [] });
  const unsubsRef = useRef([]);

  const stopData = () => { unsubsRef.current.forEach(u => u()); unsubsRef.current = []; };
  const resetData = () => {
    dataRef.current = { employees: [], questions: [], exams: [], results: [], users: [] };
    setEmployees([]); setQuestions([]); setExams([]); setResults([]); setAuthUsers([]);
  };

  // Đăng ký realtime các nguồn dữ liệu; xong lần tải đầu của tất cả thì tắt "Đang tải".
  // sources: [{ key, ref (collection hoặc query), setter }]
  const subscribeAll = (sources) => {
    let loaded = 0;
    sources.forEach(({ key, ref, setter }) => {
      let first = true;
      const unsub = onSnapshot(ref,
        snap => {
          const list = snap.docs.map(d => ({ ...d.data(), id: d.data().id ?? d.id }));
          dataRef.current[key] = list;
          setter(list);
          if (first) { first = false; loaded++; if (loaded >= sources.length) setLoading(false); }
        },
        err => {
          // Không có callback này thì listener chết âm thầm: màn hình đứng ở "Đang tải"
          // hoặc hiện dữ liệu cũ trong cache mà không ai biết là đã mất kết nối.
          console.error(`[Firestore] Lỗi đọc "${key}":`, err);
          setDbError(`Không đọc được dữ liệu "${key}" từ máy chủ (${err.code || err.message}).`);
          setLoading(false);
        });
      unsubsRef.current.push(unsub);
    });
  };

  // ── Đăng nhập: Firebase Auth quyết định ai đang đăng nhập, collection `users/{uid}` quyết định vai trò ──
  useEffect(() => {
    if (missingConfig.length) { setLoading(false); return; }
    let seq = 0;   // bỏ qua kết quả của lượt xử lý cũ nếu trạng thái đăng nhập đổi liên tục
    const unsubAuth = onAuthStateChanged(auth, async fbUser => {
      const mine = ++seq;
      stopData(); resetData();
      if (!fbUser) { setUser(null); setLoading(false); return; }
      setLoading(true);
      try {
        const prof = await getDoc(doc(db, COL.users, fbUser.uid));
        if (mine !== seq) return;
        if (!prof.exists()) throw new Error('NO_PROFILE');
        const { role, empId } = prof.data();
        if (role === 'admin') {
          setUser({ role: 'admin', id: 0, name: 'Quản trị viên', dept: 'Admin', uid: fbUser.uid });
          setView('dashboard');
          subscribeAll([
            { key: 'employees', ref: collection(db, COL.employees), setter: setEmployees },
            { key: 'questions', ref: collection(db, COL.questions), setter: setQuestions },
            { key: 'exams',     ref: collection(db, COL.exams),     setter: setExams },
            { key: 'results',   ref: collection(db, COL.results),   setter: setResults },
            { key: 'users',     ref: collection(db, COL.users),     setter: setAuthUsers },
          ]);
        } else if (role === 'employee') {
          const emp = await getDoc(doc(db, COL.employees, String(empId)));
          if (mine !== seq) return;
          if (!emp.exists()) throw new Error('NO_PROFILE');
          setUser({ role: 'employee', ...emp.data(), id: empId, uid: fbUser.uid });
          setView('home');
          subscribeAll([
            { key: 'questions', ref: collection(db, COL.questions), setter: setQuestions },
            { key: 'exams',     ref: collection(db, COL.exams),     setter: setExams },
            // Chỉ tải kết quả của chính mình (rules cũng chặn đọc của người khác)
            { key: 'results',   ref: query(collection(db, COL.results), where('empId', '==', empId)), setter: setResults },
          ]);
        } else {
          throw new Error('NO_PROFILE');
        }
      } catch (err) {
        if (mine !== seq) return;
        console.error('[Auth] Không tải được hồ sơ:', err);
        setAuthError(err.message === 'NO_PROFILE'
          ? 'Tài khoản này chưa được cấp quyền sử dụng hệ thống. Vui lòng liên hệ quản trị viên.'
          : `Không tải được hồ sơ người dùng (${err.code || err.message}).`);
        setUser(null); setLoading(false);
        signOut(auth).catch(() => {});
      }
    });
    return () => { seq++; unsubAuth(); stopData(); };
  }, []);

  const login = async (account, password, remember) => {
    setAuthError('');
    await setPersistence(auth, remember ? browserLocalPersistence : browserSessionPersistence);
    await signInWithEmailAndPassword(auth, toEmail(account), password);
  };
  const logout = () => { setActiveExam(null); setPendingExam(null); setLastResult(null); signOut(auth).catch(() => {}); };

  // ── Firebase write helpers ──
  // Firestore từ chối giá trị undefined và không nhận mảng lồng mảng → dọn trước khi ghi.
  const clean = obj => JSON.parse(JSON.stringify(obj));

  const onWriteError = (what, colName, err) => {
    console.error(`[Firestore] ${what} "${colName}" thất bại:`, err);
    setDbError(`Lưu lên máy chủ thất bại (${colName}: ${err.code || err.message}). Dữ liệu vừa thao tác CHƯA được lưu — kiểm tra mạng rồi làm lại.`);
  };

  const track = p => {
    setSaving(n => n + 1);
    return p.finally(() => setSaving(n => Math.max(0, n - 1)));
  };

  const fbSet = (colName, item) => {
    try {
      return track(setDoc(doc(db, colName, String(item.id)), clean(item)))
        .catch(err => { onWriteError('Ghi', colName, err); throw err; });
    } catch (err) {                       // setDoc ném đồng bộ khi dữ liệu sai kiểu
      onWriteError('Ghi', colName, err);
      return Promise.reject(err);
    }
  };

  const fbDel = (colName, id) =>
    track(deleteDoc(doc(db, colName, String(id))))
      .catch(err => { onWriteError('Xóa', colName, err); throw err; });

  // Side effect (ghi/xóa Firestore) phải nằm NGOÀI updater của setState: React
  // StrictMode và concurrent render gọi updater nhiều lần → sẽ ghi/xóa lặp.
  const makeSyncSetter = (colName, localSetter) => (updater) => {
    const prev = dataRef.current[colName];
    const next = typeof updater === 'function' ? updater(prev) : updater;
    dataRef.current[colName] = next;
    localSetter(next);
    next.forEach(item => {
      const old = prev.find(e => e.id === item.id);
      if (!old || JSON.stringify(old) !== JSON.stringify(item)) fbSet(colName, item).catch(()=>{});
    });
    prev.forEach(item => { if (!next.find(e => e.id === item.id)) fbDel(colName, item.id).catch(()=>{}); });
  };

  const setQuestionsSync = makeSyncSetter(COL.questions, setQuestions);
  const setExamsSync     = makeSyncSetter(COL.exams, setExams);
  const setResultsSync   = makeSyncSetter(COL.results, setResults);

  // Thu hồi quyền đăng nhập của thí sinh: xoá doc users tương ứng (rules sẽ chặn ngay mọi truy cập).
  const revokeLogins = (empId) => Promise.all(
    dataRef.current.users
      .filter(u => u.role === 'employee' && String(u.empId) === String(empId))
      .map(u => fbDel(COL.users, u.id).catch(() => {}))
  );

  // Xoá thí sinh → thu hồi luôn tài khoản đăng nhập của họ
  const setEmployeesSync = (updater) => {
    const prev = dataRef.current.employees;
    makeSyncSetter(COL.employees, setEmployees)(updater);
    const next = dataRef.current.employees;
    prev.filter(e => !next.some(n => n.id === e.id)).forEach(e => revokeLogins(e.id));
  };

  // Cấp tài khoản đăng nhập (Firebase Auth) cho danh sách thí sinh, mỗi người một mật khẩu ngẫu nhiên.
  // Trả về { creds: [{name,dept,account,password}], fails: [{emp,reason}] } để admin xem/tải về.
  const provisionLogins = async (emps) => {
    const creds = [], fails = [];
    for (const emp of emps) {
      const account = (emp.account || '').trim();
      if (!account) { fails.push({ emp, reason: 'Chưa có tên đăng nhập' }); continue; }
      const password = genPassword();
      try {
        const oldLogins = dataRef.current.users.filter(u => u.role === 'employee' && String(u.empId) === String(emp.id));
        const uid = await createLoginAccount(account, password);
        await fbSet(COL.users, { id: uid, role: 'employee', empId: emp.id, account });
        // Chỉ thu hồi quyền cũ SAU khi tài khoản mới đã tạo xong (lỗi giữa chừng không làm mất đăng nhập đang dùng)
        await Promise.all(oldLogins.map(u => fbDel(COL.users, u.id).catch(() => {})));
        creds.push({ name: emp.name, dept: emp.dept, account, password });
      } catch (err) {
        fails.push({ emp, reason: err.code === 'auth/email-already-in-use'
          ? 'Tên đăng nhập này đã tồn tại trong hệ thống đăng nhập (cần xoá/đặt lại bằng scripts/adminTools.mjs)'
          : (err.code || err.message) });
      }
    }
    return { creds, fails };
  };
  const startExam = exam => {
    // Mỗi thí sinh chỉ được thi mỗi đề một lần
    if (results.some(r => r.empId === user.id && r.examId === exam.id)) return;
    setLastResult(null);
    if (exam.sessions && exam.sessions.length > 0) {
      setPendingExam(exam);
    } else {
      setActiveExam(exam);
    }
  };
  const finishExam = async r => {
    // Ghi hỏng thì vẫn cho thí sinh xem kết quả (banner đỏ sẽ báo chưa lưu được),
    // không để kẹt lại ở màn hình làm bài.
    try { await fbSet(COL.results, r); } catch { /* onWriteError đã hiện banner */ }
    setLastResult(r);
    setActiveExam(null);
  };

  const banner = dbError && (
    <div className="fixed top-0 inset-x-0 z-[100] bg-red-600 text-white text-xs sm:text-sm px-4 py-2.5 flex items-start gap-2 shadow-lg">
      <AlertCircle size={16} className="flex-shrink-0 mt-0.5"/>
      <div className="flex-1">
        <p className="font-semibold">Mất kết nối cơ sở dữ liệu</p>
        <p className="opacity-90 break-words">{dbError}</p>
        <p className="opacity-70 mt-0.5">Project: {projectId || '(chưa cấu hình)'}</p>
      </div>
      <button onClick={()=>setDbError(null)} className="flex-shrink-0 opacity-80 hover:opacity-100"><X size={16}/></button>
    </div>
  );

  if (loading) return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center">
      {banner}
      <div className="text-center">
        <div className="flex justify-center mb-3 animate-pulse"><Emblem size={96}/></div>
        <p className="text-slate-500 text-sm">Đang tải dữ liệu...</p>
      </div>
    </div>
  );

  if (!user) return <>{banner}<Login onLogin={login} externalErr={authError}/></>;
  if (pendingExam) return <>{banner}<SessionCodeModal exam={pendingExam} onConfirm={()=>{setActiveExam(pendingExam);setPendingExam(null);}} onCancel={()=>setPendingExam(null)}/></>;
  if (activeExam) return <>{banner}<ExamScreen user={user} exam={activeExam} questions={questions} onFinish={finishExam}/></>;
  if (lastResult) {
    const exam = exams.find(e=>e.id===lastResult.examId);
    return <>{banner}<ResultScreen result={lastResult} exam={exam} questions={questions} onBack={()=>{setLastResult(null);setView("home");}}/></>;
  }

  const adminViews = {
    dashboard: <Reports results={results} exams={exams} employees={employees}/>,
    questions: <Questions questions={questions} setQuestions={setQuestionsSync}/>,
    exams:     <Exams exams={exams} setExams={setExamsSync} questions={questions}/>,
    results:   <ExamResults results={results} exams={exams} employees={employees} questions={questions} onClearAll={()=>setResultsSync([])}/>,
    employees: <EmployeesView employees={employees} setEmployees={setEmployeesSync} results={results} exams={exams} onProvision={provisionLogins}/>,
    accounts:  <Accounts employees={employees} setEmployees={setEmployeesSync} authUsers={authUsers} onProvision={provisionLogins}/>,
  };
  const empViews = {
    home:    <EmpHome user={user} exams={exams} results={results} onStart={startExam}/>,
    results: <MyResults user={user} exams={exams} questions={questions} results={results}/>,
  };

  return (
    <div className="flex bg-slate-50 min-h-screen">
      {banner}
      {showChangePw && <ChangePasswordModal onClose={()=>setShowChangePw(false)}/>}
      <Sidebar role={user.role} active={view} setActive={setView} user={user} onLogout={logout} onChangePw={()=>setShowChangePw(true)} rail={rail} setRail={setRail}/>
      <div className={`flex-1 ${rail?'md:ml-[78px]':'md:ml-[168px]'} overflow-auto pt-14 md:pt-0 pb-16 md:pb-0 transition-[margin] duration-200`}>
        <div className="p-3 sm:p-4 md:p-7">{user.role==="admin"?adminViews[view]:empViews[view]}</div>
      </div>
      {saving > 0 && (
        <div className="fixed bottom-20 md:bottom-5 right-4 z-[90] bg-slate-800 text-white text-xs px-3 py-2 rounded-lg shadow-lg flex items-center gap-2">
          <span className="w-2 h-2 rounded-full bg-amber-400 animate-pulse"/>Đang lưu lên máy chủ...
        </div>
      )}
    </div>
  );
}
