"use client";

import { useCallback, useRef, useState } from "react";
import { toast } from "sonner";
import { processExcel } from "@/app/lib/processor";

type ProcessState = "idle" | "loading" | "done";

export default function HomePage() {
  const [dragOver, setDragOver] = useState(false);
  const [processState, setProcessState] = useState<ProcessState>("idle");
  const [fileName, setFileName] = useState<string>("");
  const [outputBlob, setOutputBlob] = useState<Blob | null>(null);
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

    try {
      const buffer = await file.arrayBuffer();

      // Process in a setTimeout to allow UI to update first
      await new Promise<void>((resolve, reject) => {
        setTimeout(() => {
          try {
            const result = processExcel(buffer);
            const blob = new Blob([result], {
              type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            });
            setOutputBlob(blob);
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
  }, []);

  return (
    <main className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50 flex items-center justify-center p-6">
      <div className="w-full max-w-2xl">
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
        </div>

        {/* Drop Zone */}
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
            processState === "loading"
              ? "cursor-not-allowed opacity-80"
              : "",
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
              <p className="text-slate-400 text-sm">{fileName} を解析しています</p>
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
                処理完了
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
    </main>
  );
}
