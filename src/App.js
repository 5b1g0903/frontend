import React, { useState, useMemo } from "react";
import { Bar } from "react-chartjs-2";
import { Chart as ChartJS, CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend } from "chart.js";
import * as XLSX from "xlsx";

ChartJS.register(CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend);

export default function App() {
  const [page, setPage] = useState("home");
  const [courses, setCourses] = useState([
    { id: "1", name: "資料結構", teacher: "王老師", credit: 3, time: "一34", category: "必修", dept: "資工系" },
    { id: "2", name: "演算法", teacher: "李老師", credit: 3, time: "三56", category: "必修", dept: "資工系" },
    { id: "3", name: "經濟學導論", teacher: "張教授", credit: 2, time: "二12", category: "選修", dept: "管院" },
  ]);
  const [schedule, setSchedule] = useState([]);
  const [search, setSearch] = useState("");
  const [deptFilter, setDeptFilter] = useState("全部");

  // --- 智慧校園參數設定 ---
  const CREDIT_LIMIT = 25; // 每學期學分上限
  const GRAD_REQUIRED = 128; // 畢業總學分門檻

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = evt.target.result;
      const workbook = XLSX.read(data, { type: "binary" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const mapped = XLSX.utils.sheet_to_json(sheet).map((row, i) => ({
        id: `ex-${Date.now()}-${i}`,
        name: row.課程名稱 || row.name || "未命名",
        teacher: row.老師 || row.teacher || "未知",
        credit: Number(row.學分 || row.credit || 0),
        time: row.時間 || row.time || "",
        category: row.必選修 || row.category || "選修",
        dept: row.系所 || row.dept || "通識"
      }));
      setCourses(prev => [...prev, ...mapped]);
    };
    reader.readAsBinaryString(file);
  };

  const parseTime = (time) => {
    const weekMap = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5 };
    const day = weekMap[time[0]] || 0;
    const periods = time.slice(1).split("").map(Number).filter(n => !isNaN(n));
    return { day, periods };
  };

  // --- 核心智慧邏輯：計算課表、衝堂、學分進度 ---
  const { table, conflicts, currentCredits, requiredCredits } = useMemo(() => {
    const t = Array.from({ length: 10 }, () => Array(6).fill().map(() => []));
    const conf = [];
    let credits = 0;
    let reqCredits = 0;
    
    schedule.forEach((course) => {
      credits += course.credit;
      if (course.category === "必修") reqCredits += course.credit;
      const { day, periods } = parseTime(course.time);
      periods.forEach((p) => {
        if (p > 0 && p <= 10 && day > 0) {
          t[p - 1][day].push(course);
          if (t[p - 1][day].length > 1) conf.push(`${course.name} (週${course.time[0]}第${p}節)`);
        }
      });
    });
    return { table: t, conflicts: [...new Set(conf)], currentCredits: credits, requiredCredits: reqCredits };
  }, [schedule]);

  const addCourse = (course) => {
    if (currentCredits + course.credit > CREDIT_LIMIT) {
      alert(`⚠️ 已達本學期學分上限 (${CREDIT_LIMIT})！`);
      return;
    }
    if (schedule.find(c => c.id === course.id)) return;
    setSchedule([...schedule, course]);
  };

  const chartData = {
    labels: ["必修", "選修"],
    datasets: [{
      data: [requiredCredits, currentCredits - requiredCredits],
      backgroundColor: ["#C084FC", "#F472B6"],
      borderRadius: 12,
    }],
  };

  return (
    <div className="min-h-screen bg-[#FDFCFE] text-slate-700 font-sans">
      {/* 玻璃擬態導覽列 */}
      <nav className="sticky top-0 z-50 bg-white/70 backdrop-blur-xl border-b border-purple-100 px-6 py-4 flex flex-wrap justify-between items-center shadow-sm">
        <div className="flex items-center gap-2">
          <div className="w-9 h-9 bg-gradient-to-tr from-purple-500 to-pink-400 rounded-xl shadow-lg shadow-purple-200"></div>
          <h1 className="text-xl font-black bg-clip-text text-transparent bg-gradient-to-r from-purple-600 to-pink-500">SMART CAMPUS</h1>
        </div>
        <div className="flex bg-slate-100/50 p-1 rounded-2xl border border-slate-200/30 mt-2 md:mt-0">
          <button onClick={() => setPage("courses")} className={`px-5 py-2 rounded-xl text-sm font-bold transition ${page === "courses" ? "bg-white text-purple-600 shadow-sm" : "text-slate-400 hover:text-purple-400"}`}>課程庫</button>
          <button onClick={() => setPage("schedule")} className={`px-5 py-2 rounded-xl text-sm font-bold transition ${page === "schedule" ? "bg-white text-purple-600 shadow-sm" : "text-slate-400 hover:text-purple-400"}`}>我的課表</button>
        </div>
      </nav>

      <main className="max-w-6xl mx-auto p-4 md:p-8">
        
        {/* 智慧學分儀表板 (Mobile Friendly) */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          <div className="bg-gradient-to-br from-purple-600 to-indigo-500 p-7 rounded-[2.5rem] text-white shadow-xl shadow-purple-200">
            <p className="text-[10px] font-black uppercase tracking-widest opacity-70">Semester Credits</p>
            <h3 className="text-4xl font-black mt-1">{currentCredits} <span className="text-sm font-normal opacity-50">/ {CREDIT_LIMIT}</span></h3>
            <div className="w-full bg-black/10 h-2 rounded-full mt-4 overflow-hidden">
              <div className="bg-white h-full transition-all duration-500" style={{ width: `${(currentCredits/CREDIT_LIMIT)*100}%` }}></div>
            </div>
          </div>
          <div className="bg-white p-7 rounded-[2.5rem] border border-purple-50 shadow-sm">
            <p className="text-[10px] font-black uppercase tracking-widest text-slate-400">畢業達成率</p>
            <h3 className="text-3xl font-black text-slate-800 mt-1">{(currentCredits/GRAD_REQUIRED*100).toFixed(1)}%</h3>
            <p className="text-xs text-slate-400 mt-2 font-medium italic">尚缺 {GRAD_REQUIRED - currentCredits} 學分畢業</p>
          </div>
          <div className={`p-7 rounded-[2.5rem] border shadow-sm transition-colors ${conflicts.length > 0 ? "bg-rose-50 border-rose-100" : "bg-emerald-50 border-emerald-100"}`}>
            <p className={`text-[10px] font-black uppercase tracking-widest ${conflicts.length > 0 ? "text-rose-400" : "text-emerald-400"}`}>系統狀態</p>
            <h3 className={`text-sm font-bold mt-2 ${conflicts.length > 0 ? "text-rose-600" : "text-emerald-600"}`}>
              {conflicts.length > 0 ? `🛑 偵測到 ${conflicts.length} 處衝堂` : "✅ 課表編排正常"}
            </h3>
          </div>
        </div>

        {page === "courses" && (
          <div className="space-y-6 animate-in fade-in duration-700">
            <div className="bg-white/50 backdrop-blur-md p-4 rounded-[2rem] border border-purple-50 flex flex-wrap gap-3 items-center shadow-sm">
              <input type="text" placeholder="搜尋課程..." className="flex-1 min-w-[150px] p-3 rounded-2xl bg-white border border-slate-100 outline-none text-sm focus:ring-2 focus:ring-purple-200 transition" onChange={e => setSearch(e.target.value)} />
              <select className="p-3 rounded-2xl bg-white border border-slate-100 text-sm outline-none" onChange={e => setDeptFilter(e.target.value)}>
                <option value="全部">所有系所</option>
                <option value="資工系">資工系</option>
                <option value="管院">管院</option>
                <option value="通識">通識</option>
              </select>
              <label className="cursor-pointer bg-slate-900 text-white px-6 py-3 rounded-2xl font-bold text-xs hover:bg-purple-600 transition shadow-lg active:scale-95">
                📁 匯入
                <input type="file" onChange={handleFile} className="hidden" />
              </label>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6">
              {courses.filter(c => (deptFilter === "全部" || c.dept === deptFilter) && c.name.includes(search)).map(course => (
                <div key={course.id} className="bg-white p-6 rounded-[2.5rem] border border-slate-100 shadow-sm hover:shadow-xl hover:-translate-y-1 transition-all group">
                  <div className="flex justify-between items-start mb-4">
                    <span className={`px-3 py-1 text-[9px] font-black rounded-lg uppercase tracking-widest ${course.category === '必修' ? 'bg-purple-100 text-purple-600' : 'bg-slate-100 text-slate-400'}`}>{course.category}</span>
                    <span className="text-[10px] font-bold text-slate-300">{course.dept}</span>
                  </div>
                  <h4 className="text-lg font-bold text-slate-800 mb-4 group-hover:text-purple-600 transition-colors">{course.name}</h4>
                  <div className="flex justify-between items-center">
                    <div className="text-[11px] text-slate-400 font-medium">
                      <p>教授：{course.teacher}</p>
                      <p>時段：{course.time} ({course.credit}學分)</p>
                    </div>
                    <button onClick={() => addCourse(course)} className="w-10 h-10 bg-purple-50 text-purple-600 rounded-full flex items-center justify-center hover:bg-purple-600 hover:text-white transition shadow-sm active:scale-90 font-bold">＋</button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {page === "schedule" && (
          <div className="bg-white p-6 md:p-10 rounded-[3rem] shadow-2xl border border-purple-50 animate-in zoom-in-95 duration-500">
             <div className="flex flex-col md:flex-row justify-between items-center gap-4 mb-10">
                <h2 className="text-2xl font-black text-slate-800 tracking-tighter">WEEKLY PLANNER</h2>
                <div className="flex gap-2">
                   <button onClick={() => setSchedule([])} className="px-4 py-2 bg-rose-50 text-rose-500 rounded-xl text-xs font-bold hover:bg-rose-100 transition">清空</button>
                   <button className="px-4 py-2 bg-purple-600 text-white rounded-xl text-xs font-bold shadow-lg shadow-purple-100">匯出 PDF</button>
                </div>
             </div>
             
             <div className="overflow-x-auto -mx-6 px-6 hide-scrollbar">
                <table className="w-full min-w-[700px] border-separate border-spacing-2 text-xs">
                  <thead>
                    <tr>
                      <th className="w-12 py-4 text-slate-200 font-black">#</th>
                      {['一','二','三','四','五'].map(d => <th key={d} className="py-4 font-black text-slate-500 border-b-2 border-slate-50 uppercase tracking-widest">週{d}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {table.map((row, i) => (
                      <tr key={i}>
                        <td className="text-center font-black text-slate-200">{i + 1}</td>
                        {row.slice(1).map((cell, j) => {
                          const hasConflict = cell.length > 1;
                          return (
                            <td key={j} className={`p-1.5 h-24 rounded-[1.8rem] border transition-all ${hasConflict ? 'bg-rose-50 border-rose-100 shadow-[inset_0_0_10px_rgba(244,63,94,0.05)]' : 'bg-slate-50/40 border-slate-100/50'}`}>
                              {cell.map(c => (
                                <div key={c.id} className={`p-2.5 mb-1.5 rounded-2xl text-[10px] font-black shadow-sm border animate-in fade-in slide-in-from-top-1 ${c.category === '必修' ? 'bg-purple-600 text-white border-purple-400' : 'bg-white text-purple-600 border-purple-100'}`}>
                                  {c.name}
                                </div>
                              ))}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
             </div>

             {/* 智慧學分分佈圖 */}
             <div className="mt-16 grid grid-cols-1 md:grid-cols-2 gap-10 items-center">
                <div className="h-48 w-full">
                  <Bar data={chartData} options={{ responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, grid: { display: false } } } }} />
                </div>
                <div className="space-y-4">
                   <div className="p-5 bg-purple-50/50 rounded-3xl border border-purple-100/50">
                      <p className="text-[10px] font-black text-purple-400 uppercase tracking-widest">必修進度</p>
                      <p className="text-sm font-bold text-purple-700 mt-1">目前已選 {requiredCredits} 學分必修課程</p>
                   </div>
                   <div className="p-5 bg-pink-50/50 rounded-3xl border border-pink-100/50">
                      <p className="text-[10px] font-black text-pink-400 uppercase tracking-widest">選修進度</p>
                      <p className="text-sm font-bold text-pink-700 mt-1">目前已選 {currentCredits - requiredCredits} 學分選修課程</p>
                   </div>
                </div>
             </div>
          </div>
        )}
      </main>

      <style dangerouslySetInnerHTML={{__html: `
        .hide-scrollbar::-webkit-scrollbar { display: none; }
        .hide-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }
      `}} />
    </div>
  );
}