"use client";

import { useCallback, useMemo, useRef, useState } from "react";
import { toast } from "sonner";
import { processExcelWithPreview, InstitutionRecord } from "@/app/lib/processor";

type ProcessState = "idle" | "loading" | "done";
type SortDir = "asc" | "desc";

const COLUMNS: { key: keyof InstitutionRecord; label: string }[] = [
  { key: "code", label: "医療機関コード" },
  { key: "name", label: "医療機関名" },
  { key: "postalCode", label: "郵便番号" },
  { key: "address", label: "所在地" },
  { key: "phone", label: "電話番号" },
  { key: "category", label: "種別" },
  { key: "status", label: "状態" },
  { key: "bedsGeneral", label: "一般病床" },
  { key: "bedsPsychiatric", label: "精神病床" },
  { key: "bedsNursing", label: "療養病床" },
  { key: "bedsTuberculosis", label: "結核病床" },
  { key: "fullTimeDoctors", label: "常勤医数" },
  { key: "partTimeDoctors", label: "非常勤医数" },
  { key: "departments", label: "診療科目" },
  { key: "founder", label: "開設者" },
  { key: "manager", label: "管理者" },
  { key: "designatedDate", label: "指定年月日" },
  { key: "renewalDate", label: "指定更新日" },
];

export default function HomePage() {
  const [dragOver, setDragOver] = useState(false);
  const [processState, setProcessState] = useState<ProcessState>("idle");
  const [fileName, setFileName] = useState<string>("");
  const [outputBlob, setOutputBlob] = useState<Blob | null>(null);
  const [records, setRecords] = useState<InstitutionRecord[]>([]);
  const [sortKey, setSortKey] = useState<keyof InstitutionRecord | null>(null);
  const [sortDir, setSortDir] = useState<SortDir>("asc");
  const inputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback(async (file: File) => {
    if (!file.name.endsWith(".xlsx")) {
      toast.error("ファイル形式エラー", {
        description: ".xlsx ファイルのみ受け付けています。",
      });
      return;
    }

    setFileName(file.name);
    setProcessState("loading");
    setOutputBlob(null);
    setRecords([]);
    setSortKey(null);

    try {
      const buffer = await file.arrayBuffer();

      await new Promise<void>((resolve, reject) => {
        setTimeout(() => {
          try {
            const { output, records: recs } = processExcelWithPreview(buffer);
            const blob = new Blob([output], {
              type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            });
            setOutputBlob(blob);
            setRecords(recs);
            setProcessState("done");
            resolve();
          } catch (err) {
            reject(err);
          }
        }, 50);
      });

      toast.success("処理完了", {
        description: "Excelファイルの整形が完了しました。",
      });
    } catch (err) {
      setProcessState("idle");
      const message =
        err instanceof Error ? err.message : "不明なエラーが発生しました。";
      toast.error("処理エラー", { description: message });
    }
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      setDragOver(false);
      const file = e.dataTransfer.files[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  const handleDragOver = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setDragOver(true);
  }, []);

  const handleDragLeave = useCallback(() => {
    setDragOver(false);
  }, []);

  const handleInputChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) handleFile(file);
      e.target.value = "";
    },
    [handleFile]
  );

  const handleDownload = useCallback(() => {
    if (!outputBlob) return;
    const url = URL.createObjectURL(outputBlob);
    const a = document.createElement("a");
    a.href = url;
    const baseName = fileName.replace(/\.xlsx$/i, "");
    a.download = `${baseName}_整形済み.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, [outputBlob, fileName]);

  const handleReset = useCallback(() => {
    setProcessState("idle");
    setOutputBlob(null);
    setFileName("");
    setRecords([]);
    setSortKey(null);
  }, []);

  const handleSort = useCallback(
    (key: keyof InstitutionRecord) => {
      if (sortKey === key) {
        setSortDir((d) => (d === "asc" ? "desc" : "asc"));
      } else {
        setSortKey(key);
        setSortDir("asc");
      }
    },
    [sortKey]
  );

  const sortedRecords = useMemo(() => {
    if (!sortKey) return records;
    return [...records].sort((a, b) => {
      const av = a[sortKey];
      const bv = b[sortKey];
      const aStr = av === "" || av === undefined ? "" : String(av);
      const bStr = bv === "" || bv === undefined ? "" : String(bv);
      const aNum = Number(av);
      const bNum = Number(bv);
      let cmp: number;
      if (!isNaN(aNum) && !isNaN(bNum) && av !== "" && bv !== "") {
        cmp = aNum - bNum;
      } else {
        cmp = aStr.localeCompare(bStr, "ja");
      }
      return sortDir === "asc" ? cmp : -cmp;
    });
  }, [records, sortKey, sortDir]);

  return (
    <main className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50 p-6">
      <div className="w-full max-w-7xl mx-auto">
        {/* Header */}
        <div className="text-center mb-8">
          <h1 className="text-2xl font-bold text-slate-800 mb-2">
            厚生局 医療機関データ整形ツール
          </h1>
          <a
            href="https://kouseikyoku.mhlw.go.jp/tokaihokuriku/newpage_00287.html"
            target="_blank"
            rel="noopener noreferrer"
            className="text-sm text-blue-500 hover:underline"
          >
            東海北陸厚生局 医療機関一覧ページ
          </a>
          <p className="text-xs text-slate-400 mt-2">v0.1.0</p>
        </div>

        {/* Drop Zone */}
        <div className="max-w-2xl mx-auto mb-6">
          <div
            onClick={() =>
              processState !== "loading" && inputRef.current?.click()
            }
            onDrop={handleDrop}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            className={[
              "relative rounded-2xl border-2 border-dashed transition-all duration-200",
              "flex flex-col items-center justify-center",
              "min-h-72 p-10 text-center select-none",
              dragOver
                ? "border-blue-500 bg-blue-50 scale-[1.01] cursor-copy"
                : processState === "done"
                ? "border-green-400 bg-green-50 cursor-default"
                : "border-slate-300 bg-white hover:border-blue-400 hover:bg-blue-50/40 cursor-pointer",
              processState === "loading" ? "cursor-not-allowed opacity-80" : "",
            ].join(" ")}
          >
            <input
              ref={inputRef}
              type="file"
              accept=".xlsx"
              className="hidden"
              onChange={handleInputChange}
            />

            {/* Idle state */}
            {processState === "idle" && (
              <>
                <div className="mb-4 p-4 rounded-full bg-blue-100">
                  <svg
                    className="w-10 h-10 text-blue-500"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={1.5}
                      d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                    />
                  </svg>
                </div>
                <p className="text-slate-700 font-semibold text-lg mb-1">
                  Excelファイルをドロップ
                </p>
                <p className="text-slate-400 text-sm mb-4">
                  またはクリックしてファイルを選択
                </p>
                <span className="inline-block px-3 py-1 rounded-full bg-slate-100 text-slate-500 text-xs font-mono">
                  .xlsx のみ対応
                </span>
              </>
            )}

            {/* Loading state */}
            {processState === "loading" && (
              <>
                <div className="mb-4">
                  <svg
                    className="w-12 h-12 text-blue-500 animate-spin"
                    fill="none"
                    viewBox="0 0 24 24"
                  >
                    <circle
                      className="opacity-25"
                      cx="12"
                      cy="12"
                      r="10"
                      stroke="currentColor"
                      strokeWidth="4"
                    />
                    <path
                      className="opacity-75"
                      fill="currentColor"
                      d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"
                    />
                  </svg>
                </div>
                <p className="text-slate-700 font-semibold text-lg mb-1">
                  処理中...
                </p>
                <p className="text-slate-400 text-sm">
                  {fileName} を解析しています
                </p>
              </>
            )}

            {/* Done state */}
            {processState === "done" && (
              <>
                <div className="mb-4 p-4 rounded-full bg-green-100">
                  <svg
                    className="w-10 h-10 text-green-500"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={2}
                      d="M5 13l4 4L19 7"
                    />
                  </svg>
                </div>
                <p className="text-slate-700 font-semibold text-lg mb-1">
                  処理完了 — {records.length.toLocaleString()} 件
                </p>
                <p className="text-slate-400 text-sm mb-6">
                  {fileName} の整形が完了しました
                </p>
                <div className="flex gap-3">
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      handleDownload();
                    }}
                    className="flex items-center gap-2 px-6 py-2.5 rounded-lg bg-green-600 text-white font-medium text-sm hover:bg-green-700 transition-colors shadow-sm"
                  >
                    <svg
                      className="w-4 h-4"
                      fill="none"
                      stroke="currentColor"
                      viewBox="0 0 24 24"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        strokeWidth={2}
                        d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"
                      />
                    </svg>
                    ダウンロード
                  </button>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      handleReset();
                    }}
                    className="flex items-center gap-2 px-6 py-2.5 rounded-lg bg-slate-100 text-slate-600 font-medium text-sm hover:bg-slate-200 transition-colors"
                  >
                    <svg
                      className="w-4 h-4"
                      fill="none"
                      stroke="currentColor"
                      viewBox="0 0 24 24"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        strokeWidth={2}
                        d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"
                      />
                    </svg>
                    別のファイルを処理
                  </button>
                </div>
              </>
            )}
          </div>
        </div>

        {/* Preview Table */}
        {processState === "done" && sortedRecords.length > 0 && (
          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
            <div className="overflow-auto max-h-[60vh]">
              <table className="w-full text-xs border-collapse">
                <thead className="sticky top-0 z-10">
                  <tr>
                    {COLUMNS.map((col) => (
                      <th
                        key={col.key}
                        onClick={() => handleSort(col.key)}
                        className="bg-[#1E3A5F] text-white font-semibold px-3 py-2 text-left whitespace-nowrap cursor-pointer select-none hover:bg-[#2a4f7c] transition-colors"
                      >
                        <span className="flex items-center gap-1">
                          {col.label}
                          {sortKey === col.key ? (
                            <span className="text-blue-300">
                              {sortDir === "asc" ? "↑" : "↓"}
                            </span>
                          ) : (
                            <span className="text-slate-500 opacity-50">↕</span>
                          )}
                        </span>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {sortedRecords.map((rec, i) => (
                    <tr
                      key={i}
                      className={
                        i % 2 === 0 ? "bg-[#EBF3FB]" : "bg-white"
                      }
                    >
                      {COLUMNS.map((col) => (
                        <td
                          key={col.key}
                          className="px-3 py-1.5 text-slate-700 whitespace-nowrap border-b border-slate-100"
                        >
                          {rec[col.key] === "" || rec[col.key] === undefined
                            ? ""
                            : String(rec[col.key])}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </main>
  );
}
