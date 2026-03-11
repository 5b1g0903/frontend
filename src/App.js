import React, { useState, useMemo, useRef } from "react";
import { Bar } from "react-chartjs-2";
import { Chart as ChartJS, CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend } from "chart.js";
import * as XLSX from "xlsx";
import { useReactToPrint } from "react-to-print";

ChartJS.register(CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend);

// --- 智慧欄位字典：解決各校 Excel 表頭不一的問題 ---
const columnDictionary = {
  name: ['課程名稱', '科目', '課名', '課程', '科目名稱', 'course'],
  teacher: ['老師', '教師', '授課教師', '教授', '任課教師', 'teacher'],
  credit: ['學分', '學分數', 'credit'],
  category: ['必選修', '修別', '必/選修', '屬性', '選必修', '必選', 'type'],
  time: ['時間', '節次', '星期', '上課時間', '星期/節次', 'time'],
  dept: ['系所', '開課單位', '系級', 'dept']
};

const getStandardKey = (excelKey) => {
  for (const [standardKey, keywords] of Object.entries(columnDictionary)) {
    if (keywords.some(k => excelKey.includes(k))) return standardKey;
  }
  return null;
};

export default function App() {
  const [page, setPage] = useState("courses");
  const [courses, setCourses] = useState([
    { id: "1", name: "資料結構", teacher: "王老師", credit: 3, time: "一34", category: "必修", dept: "資工系" },
    { id: "2", name: "演算法", teacher: "李老師", credit: 3, time: "三56", category: "必修", dept: "資工系" },
    { id: "3", name: "經濟學導論", teacher: "張教授", credit: 2, time: "二12", category: "選修", dept: "管院" },
  ]);
  const [schedule, setSchedule] = useState([]);
  const [search, setSearch] = useState("");
  const [deptFilter, setDeptFilter] = useState("全部");

  // PDF 匯出用的 Ref
  const componentRef = useRef();
  const handlePrint = useReactToPrint({
    content: () => componentRef.current,
    documentTitle: '我的大學課表',
  });

  const CREDIT_LIMIT = 25;
  const GRAD_REQUIRED = 128;

  // --- 智慧 Excel 匯入邏輯 ---
  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = evt.target.result;
      const workbook = XLSX.read(data, { type: "binary" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawData = XLSX.utils.sheet_to_json(sheet);
      
      const mapped = rawData.map((row, i) => {
        let standardizedRow = { id: `ex-${Date.now()}-${i}` };
        // 遍歷 Excel 的每一欄，透過字典比對轉換
        Object.entries(row).forEach(([key, value]) => {
          const standardKey = getStandardKey(key);
          if (standardKey) {
            standardizedRow[standardKey] = (standardKey === 'credit') ? Number(value) : value;
          }
        });
        // 補足漏掉的欄位預設值
        return {
          ...standardizedRow,
          name: standardizedRow.name || "未命名課程",
          teacher: standardizedRow.teacher || "未知",
          credit: standardizedRow.credit || 0,
          time: standardizedRow.time || "",
          category: standardizedRow.category || "選修",
          dept: standardizedRow.dept || "未知"
        };
      });
      
      setCourses(prev => [...prev, ...mapped]);
      alert(`成功匯入 ${mapped.length} 門課程！`);
    };
    reader.readAsBinaryString(file);
  };

  const parseTime = (time) => {
    if (!time || typeof time !== 'string') return { day: 0, periods: [] };
    const weekMap = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5 };
    const day = weekMap[time[0]] || 0;
    const periods = time.slice(1).split("").map(Number).filter(n => !isNaN(n));
    return { day, periods };
  };

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
        
        {/* 智慧學分儀表板 */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          <div className="bg-gradient-to-br from-purple-600 to-indigo-500 p-7 rounded-[2.5rem] text-white shadow-xl shadow-purple-200">
            <p className="text-[10px] font-black uppercase tracking-widest opacity-70">學期學分進度</p>
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
          <div className="space-y-6">
            <div className="bg-white/50 backdrop-blur-md p-4 rounded-[2rem] border border-purple-50 flex flex-wrap gap-3 items-center shadow-sm">
              <input type="text" placeholder="搜尋課程..." className="flex-1 min-w-[150px] p-3 rounded-2xl bg-white border border-slate-100 outline-none text-sm focus:ring-2 focus:ring-purple-200 transition" onChange={e => setSearch(e.target.value)} />
              <label className="cursor-pointer bg-slate-900 text-white px-6 py-3 rounded-2xl font-bold text-xs hover:bg-purple-600 transition shadow-lg active:scale-95">
                📁 匯入 Excel
                <input type="file" onChange={handleFile} className="hidden" accept=".xlsx, .xls" />
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
          <div ref={componentRef} className="bg-white p-6 md:p-10 rounded-[3rem] shadow-2xl border border-purple-50">
             <div className="flex flex-col md:flex-row justify-between items-center gap-4 mb-10">
                <h2 className="text-2xl font-black text-slate-800 tracking-tighter">WEEKLY PLANNER</h2>
                <div className="flex gap-2">
                   <button onClick={() => setSchedule([])} className="px-4 py-2 bg-rose-50 text-rose-500 rounded-xl text-xs font-bold hover:bg-rose-100 transition no-print">清空</button>
                   <button onClick={handlePrint} className="px-4 py-2 bg-purple-600 text-white rounded-xl text-xs font-bold shadow-lg shadow-purple-100 no-print">匯出 PDF</button>
                </div>
             </div>
             
             <div className="overflow-x-auto -mx-6 px-6">
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
                        {row.slice(1).map((cell, j) => (
                          <td key={j} className={`p-1.5 h-24 rounded-[1.8rem] border ${cell.length > 1 ? 'bg-rose-50 border-rose-100' : 'bg-slate-50/40 border-slate-100/50'}`}>
                            {cell.map(c => (
                              <div key={c.id} className={`p-2.5 mb-1.5 rounded-2xl text-[10px] font-black shadow-sm border ${c.category === '必修' ? 'bg-purple-600 text-white border-purple-400' : 'bg-white text-purple-600 border-purple-100'}`}>
                                {c.name}
                              </div>
                            ))}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
             </div>

             <div className="mt-16 grid grid-cols-1 md:grid-cols-2 gap-10 items-center">
                <div className="h-48 w-full"><Bar data={chartData} options={{ responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }} /></div>
                <div className="space-y-4">
                   <div className="p-5 bg-purple-50/50 rounded-3xl border border-purple-100/50">
                      <p className="text-xs font-bold text-purple-700">必修進度：已選 {requiredCredits} 學分</p>
                   </div>
                </div>
             </div>
          </div>
        )}
      </main>

      <style dangerouslySetInnerHTML={{__html: `
        @media print {
          .no-print { display: none !important; }
          body { background: white !important; }
          .min-h-screen { height: auto !important; }
        }
      `}} />
    </div>
  );
}