import React, { useState, useMemo, useRef } from "react";
import { Bar } from "react-chartjs-2";
import { Chart as ChartJS, CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend } from "chart.js";
import * as XLSX from "xlsx";
import { useReactToPrint } from "react-to-print";

ChartJS.register(CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend);

// --- 智慧字典：擴充匹配關鍵字 ---
const columnDictionary = {
  name: ['課程名稱', '科目', '課名', '課程', '科目名稱', 'subject', 'course'],
  teacher: ['老師', '教師', '授課教師', '教授', '任課教師', 'professor'],
  credit: ['學分', '學分數', '學分/時數', 'units'],
  category: ['必選修', '修別', '必/選修', '屬性', '選必修', 'type'],
  time: ['時間', '節次', '星期', '上課時間', '星期/節次', 'schedule'],
  dept: ['系所', '開課單位', '系級', 'dept']
};

const getStandardKey = (excelKey) => {
  const key = String(excelKey || "").toLowerCase();
  for (const [standardKey, keywords] of Object.entries(columnDictionary)) {
    if (keywords.some(k => key.includes(k))) return standardKey;
  }
  return null;
};

export default function App() {
  const [page, setPage] = useState("courses");
  const [courses, setCourses] = useState([]);
  const [schedule, setSchedule] = useState([]);
  const [search, setSearch] = useState("");

  const componentRef = useRef();
  const handlePrint = useReactToPrint({
    content: () => componentRef.current,
    documentTitle: '我的智慧課表',
  });

  const CREDIT_LIMIT = 25;
  const GRAD_REQUIRED = 128;

  // --- 核心：報表級 Excel 智慧解析 ---
  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = evt.target.result;
      const workbook = XLSX.read(data, { type: "binary" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      
      // 1. 轉成原始二維陣列 (含空白格)
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      
      // 2. 智慧搜尋真正的「標頭列」位置 (跳過學校名稱等標題)
      const headerIndex = rows.findIndex(row => 
        row.some(cell => {
          const c = String(cell || "");
          return c.includes("科目") || c.includes("課程") || c.includes("學分");
        })
      );

      if (headerIndex === -1) {
        alert("找不到有效的課程欄位！請確認 Excel 內包含「科目」或「學分」等字樣。");
        return;
      }

      const headerRow = rows[headerIndex];
      const dataRows = rows.slice(headerIndex + 1);

      // 3. 轉換與過濾雜訊
      const mapped = dataRows.map((row, i) => {
        let stdRow = { id: `ex-${Date.now()}-${i}` };
        row.forEach((cell, cellIndex) => {
          const excelKey = headerRow[cellIndex];
          const stdKey = getStandardKey(excelKey);
          if (stdKey) {
            if (stdKey === 'credit') {
              // 處理學分 (洗掉括號或非數字符號)
              const val = String(cell || "").replace(/[^\d.]/g, '');
              stdRow[stdKey] = val ? Number(val) : 0;
            } else {
              stdRow[stdKey] = cell ? String(cell).trim() : "";
            }
          }
        });
        return stdRow;
      }).filter(c => {
        // ✨ 強力過濾器：過濾標題列、小計列、說明文字
        const isTrash = !c.name || c.name.includes("學年") || c.name.includes("學期") || c.name.includes("合計") || c.name.length < 2;
        return !isTrash && c.credit > 0;
      });

      setCourses(prev => [...prev, ...mapped]);
      alert(`智慧解析完成！成功跳過報表標題，匯入 ${mapped.length} 門有效課程。`);
    };
    reader.readAsBinaryString(file);
  };

  // --- 時間解析：支援各種奇怪的格式 ---
  const parseTime = (timeStr) => {
    if (!timeStr) return { day: 0, periods: [] };
    const weekMap = { "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "1": 1, "2": 2, "3": 3, "4": 4, "5": 5 };
    
    let day = 0;
    // 優先找括號內的星期，或是第一個出現的星期字眼
    for (let char of timeStr) {
      if (weekMap[char]) {
        day = weekMap[char];
        break;
      }
    }
    // 提取數字節次
    const periods = timeStr.replace(/[^\d]/g, "").split("").map(Number).filter(n => n > 0 && n <= 10);
    return { day, periods };
  };

  const { table, conflicts, currentCredits, requiredCredits } = useMemo(() => {
    const t = Array.from({ length: 10 }, () => Array(6).fill().map(() => []));
    let conf = [], credits = 0, reqCredits = 0;
    
    schedule.forEach((course) => {
      credits += course.credit;
      if (String(course.category).includes("必")) reqCredits += course.credit;
      const { day, periods } = parseTime(course.time);
      periods.forEach((p) => {
        if (p > 0 && p <= 10 && day > 0) {
          t[p - 1][day].push(course);
          if (t[p - 1][day].length > 1) conf.push(`${course.name} (週${day}第${p}節)`);
        }
      });
    });
    return { table: t, conflicts: [...new Set(conf)], currentCredits: credits, requiredCredits: reqCredits };
  }, [schedule]);

  const addCourse = (course) => {
    if (currentCredits + course.credit > CREDIT_LIMIT) return alert(`⚠️ 學分上限為 ${CREDIT_LIMIT}`);
    if (schedule.some(c => c.id === course.id)) return;
    setSchedule([...schedule, course]);
  };

  const chartData = {
    labels: ["必修", "選修"],
    datasets: [{ data: [requiredCredits, currentCredits - requiredCredits], backgroundColor: ["#C084FC", "#F472B6"], borderRadius: 12 }],
  };

  return (
    <div className="min-h-screen bg-[#FDFCFE] text-slate-700 font-sans">
      <nav className="sticky top-0 z-50 bg-white/70 backdrop-blur-xl border-b border-purple-100 px-6 py-4 flex justify-between items-center shadow-sm">
        <div className="flex items-center gap-2">
          <div className="w-9 h-9 bg-gradient-to-tr from-purple-500 to-pink-400 rounded-xl shadow-lg"></div>
          <h1 className="text-xl font-black bg-clip-text text-transparent bg-gradient-to-r from-purple-600 to-pink-500 tracking-tighter">SMART CAMPUS</h1>
        </div>
        <div className="flex bg-slate-100/50 p-1 rounded-2xl border border-slate-200/30">
          <button onClick={() => setPage("courses")} className={`px-5 py-2 rounded-xl text-sm font-bold transition ${page === "courses" ? "bg-white text-purple-600 shadow-sm" : "text-slate-400"}`}>課程庫</button>
          <button onClick={() => setPage("schedule")} className={`px-5 py-2 rounded-xl text-sm font-bold transition ${page === "schedule" ? "bg-white text-purple-600 shadow-sm" : "text-slate-400"}`}>我的課表</button>
        </div>
      </nav>

      <main className="max-w-6xl mx-auto p-4 md:p-8">
        {/* 儀表板區域 */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          <div className="bg-gradient-to-br from-purple-600 to-indigo-500 p-7 rounded-[2.5rem] text-white shadow-xl">
            <p className="text-[10px] font-black uppercase tracking-widest opacity-70">學期學分</p>
            <h3 className="text-4xl font-black mt-1">{currentCredits} <span className="text-sm font-normal opacity-50">/ {CREDIT_LIMIT}</span></h3>
            <div className="w-full bg-black/10 h-2 rounded-full mt-4 overflow-hidden"><div className="bg-white h-full transition-all duration-500" style={{ width: `${(currentCredits/CREDIT_LIMIT)*100}%` }}></div></div>
          </div>
          <div className="bg-white p-7 rounded-[2.5rem] border border-purple-50 shadow-sm">
            <p className="text-[10px] font-black uppercase tracking-widest text-slate-400">畢業進度</p>
            <h3 className="text-3xl font-black text-slate-800 mt-1">{(currentCredits/GRAD_REQUIRED*100).toFixed(1)}%</h3>
          </div>
          <div className={`p-7 rounded-[2.5rem] border shadow-sm ${conflicts.length > 0 ? "bg-rose-50 border-rose-100" : "bg-emerald-50 border-emerald-100"}`}>
            <p className={`text-[10px] font-black uppercase tracking-widest ${conflicts.length > 0 ? "text-rose-400" : "text-emerald-400"}`}>狀態</p>
            <h3 className="text-sm font-bold mt-2">{conflicts.length > 0 ? `🛑 衝堂：${conflicts.length} 處` : "✅ 正常"}</h3>
          </div>
        </div>

        {page === "courses" && (
          <div className="space-y-6">
            <div className="bg-white/50 backdrop-blur-md p-4 rounded-[2rem] border border-purple-50 flex flex-wrap gap-3 items-center shadow-sm">
              <input type="text" placeholder="搜尋..." className="flex-1 p-3 rounded-2xl bg-white border border-slate-100 outline-none text-sm" onChange={e => setSearch(e.target.value)} />
              <label className="cursor-pointer bg-slate-900 text-white px-6 py-3 rounded-2xl font-bold text-xs hover:bg-purple-600 transition shadow-lg">
                📁 匯入檔案 <input type="file" onChange={handleFile} className="hidden" accept=".xlsx, .xls, .csv" />
              </label>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6">
              {courses.filter(c => c.name.includes(search)).map(course => (
                <div key={course.id} className="bg-white p-6 rounded-[2.5rem] border border-slate-100 shadow-sm hover:shadow-xl hover:-translate-y-1 transition-all group">
                  <div className="flex justify-between items-start mb-4">
                    <span className={`px-3 py-1 text-[9px] font-black rounded-lg ${String(course.category).includes('必') ? 'bg-purple-100 text-purple-600' : 'bg-slate-100 text-slate-400'}`}>{course.category}</span>
                    <span className="text-[10px] font-bold text-slate-300">{course.dept}</span>
                  </div>
                  <h4 className="text-lg font-bold text-slate-800 mb-4">{course.name}</h4>
                  <div className="flex justify-between items-center">
                    <div className="text-[11px] text-slate-400 font-medium">
                      <p>教授：{course.teacher}</p>
                      <p>時段：{course.time} ({course.credit}學分)</p>
                    </div>
                    <button onClick={() => addCourse(course)} className="w-10 h-10 bg-purple-50 text-purple-600 rounded-full flex items-center justify-center hover:bg-purple-600 hover:text-white transition font-bold">＋</button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {page === "schedule" && (
          <div ref={componentRef} className="bg-white p-6 md:p-10 rounded-[3rem] shadow-2xl border border-purple-50">
             <div className="flex justify-between items-center mb-10 no-print">
                <h2 className="text-2xl font-black text-slate-800 tracking-tighter">WEEKLY PLANNER</h2>
                <div className="flex gap-2">
                   <button onClick={() => setSchedule([])} className="px-4 py-2 bg-rose-50 text-rose-500 rounded-xl text-xs font-bold">清空</button>
                   <button onClick={handlePrint} className="px-4 py-2 bg-purple-600 text-white rounded-xl text-xs font-bold shadow-lg shadow-purple-200">匯出 PDF</button>
                </div>
             </div>
             <div className="overflow-x-auto">
                <table className="w-full border-separate border-spacing-2 text-xs">
                  <thead>
                    <tr>
                      <th className="w-12 py-4 text-slate-200">#</th>
                      {['一','二','三','四','五'].map(d => <th key={d} className="py-4 font-black text-slate-500 border-b-2 border-slate-50">週{d}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {table.map((row, i) => (
                      <tr key={i}>
                        <td className="text-center font-black text-slate-200">{i + 1}</td>
                        {row.slice(1).map((cell, j) => (
                          <td key={j} className={`p-1.5 h-24 rounded-[1.8rem] border ${cell.length > 1 ? 'bg-rose-50 border-rose-100 shadow-[inset_0_0_10px_rgba(244,63,94,0.05)]' : 'bg-slate-50/40 border-slate-100/50'}`}>
                            {cell.map(c => (
                              <div key={c.id} className={`p-2.5 mb-1.5 rounded-2xl text-[10px] font-black shadow-sm border ${String(c.category).includes('必') ? 'bg-purple-600 text-white' : 'bg-white text-purple-600 border-purple-100'}`}>
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
          </div>
        )}
      </main>

      <style>{`
        @media print { .no-print { display: none !important; } .shadow-2xl { border: none !important; box-shadow: none !important; } body { background: white !important; } }
      `}</style>
    </div>
  );
}