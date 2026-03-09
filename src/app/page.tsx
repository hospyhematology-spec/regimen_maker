"use client";
import React, { useCallback, useRef, useState } from "react";
import { useAppStore } from "@/store/appStore";
import type { AppStep } from "@/store/appStore";
import type { RegimenMaster, RichBulletItem, AdministrationStep, RegimenBlock } from "@/types/regimen";

// ─── Progress Steps ─────────────────────────────────────────────────
const STEPS: { id: AppStep; label: string }[] = [
  { id: "upload", label: "アップロード" },
  { id: "extracting", label: "自動抽出" },
  { id: "review-basic", label: "基本情報" },
  { id: "review-regimen", label: "投与方法" },
  { id: "export", label: "出力" },
];

function ProgressBar({ current }: { current: AppStep }) {
  const idx = STEPS.findIndex((s) => s.id === current);
  return (
    <div className="progress-steps">
      {STEPS.map((s, i) => (
        <div
          key={s.id}
          className={`step-item ${i < idx ? "completed" : ""} ${i === idx ? "active" : ""}`}
        >
          <div className="step-circle">{i < idx ? "✓" : i + 1}</div>
          <div className="step-label">{s.label}</div>
        </div>
      ))}
    </div>
  );
}

// ─── Confidence Badge ─────────────────────────────────────────────
function ConfBadge({ value }: { value: number }) {
  const pct = Math.round(value * 100);
  const cls = pct >= 85 ? "confidence-high" : pct >= 65 ? "confidence-mid" : "confidence-low";
  return <span className={`confidence-badge ${cls}`}>信頼度 {pct}%</span>;
}

// ─── Source Modal ─────────────────────────────────────────────────
function SourceModal({ traces, onClose }: { traces: import("@/types/regimen").SourceTrace[]; onClose: () => void }) {
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-box fade-in" onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.25rem" }}>
          <h3 style={{ fontSize: "1rem", fontWeight: 700 }}>📄 抽出元データ</h3>
          <button className="btn btn-secondary" onClick={onClose} style={{ padding: "0.3rem 0.75rem", fontSize: "0.8rem" }}>✕</button>
        </div>
        {traces.map((t, i) => (
          <div key={i} style={{ marginBottom: "1rem", padding: "0.875rem", background: "var(--bg-secondary)", borderRadius: 10, border: "1px solid var(--border-subtle)" }}>
            <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap", marginBottom: "0.5rem" }}>
              <span className="source-pill">📁 {t.fileName ?? t.sourceType}</span>
              {t.sheetName && <span className="source-pill">🗂 {t.sheetName}</span>}
              {t.cellRange && <span className="source-pill">🔲 {t.cellRange}</span>}
              {t.page && <span className="source-pill">📄 P.{t.page}</span>}
              <ConfBadge value={t.confidence} />
            </div>
            <div style={{ fontSize: "0.8rem", color: "var(--text-muted)", marginBottom: "0.35rem" }}>引用テキスト</div>
            <div style={{ fontSize: "0.85rem", color: "var(--text-secondary)", background: "var(--bg-card)", padding: "0.5rem 0.75rem", borderRadius: 6, fontFamily: "monospace" }}>
              {t.quotedText}
            </div>
            {t.aiInterpretation && t.aiInterpretation !== t.quotedText && (
              <>
                <div style={{ fontSize: "0.8rem", color: "var(--text-muted)", margin: "0.5rem 0 0.25rem" }}>AI解釈</div>
                <div style={{ fontSize: "0.85rem", color: "var(--accent)" }}>{t.aiInterpretation}</div>
              </>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── History Modal ─────────────────────────────────────────────────
function HistoryModal({ onClose, onSelect }: { onClose: () => void; onSelect: (r: Partial<RegimenMaster>) => void }) {
  const [history, setHistory] = useState<RegimenMaster[]>([]);
  const [loading, setLoading] = useState(true);

  React.useEffect(() => {
    try {
      const stored = localStorage.getItem("regimen_history");
      if (stored) {
        setHistory(JSON.parse(stored));
      }
    } catch (e) {
      console.error(e);
    } finally {
      setLoading(false);
    }
  }, []);

  const handleSelect = (id: string) => {
    const data = history.find(h => h.id === id);
    if (data) onSelect(data);
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-box fade-in" onClick={(e) => e.stopPropagation()} style={{ maxWidth: 600 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.25rem" }}>
          <h3 style={{ fontSize: "1.1rem", fontWeight: 700 }}>⏳ 過去のレジメンと比較・反映</h3>
          <button className="btn btn-secondary" onClick={onClose} style={{ padding: "0.3rem 0.75rem", fontSize: "0.8rem" }}>✕</button>
        </div>
        {loading ? (
          <div style={{ padding: "2rem", textAlign: "center", color: "var(--text-muted)" }}><span className="spin" style={{ display: "inline-block" }}>⟳</span> 読み込み中...</div>
        ) : history.length === 0 ? (
          <div style={{ padding: "2rem", textAlign: "center", color: "var(--text-muted)" }}>過去のレジメンデータがありません</div>
        ) : (
          <div style={{ display: "grid", gap: "0.75rem", maxHeight: "60vh", overflowY: "auto" }}>
            {history.map(h => (
              <div key={h.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "0.875rem", background: "var(--bg-secondary)", borderRadius: 10, border: "1px solid var(--border-subtle)" }}>
                <div>
                  <div style={{ fontWeight: 600, color: "var(--text-primary)" }}>{h.basicInfo.regimenName || "名称未設定"}</div>
                  <div style={{ fontSize: "0.8rem", color: "var(--text-muted)", marginTop: "0.2rem" }}>{h.basicInfo.cancerType || "がん種不明"} | {new Date(h.createdAt).toLocaleDateString()}</div>
                </div>
                <button className="btn btn-primary" onClick={() => handleSelect(h.id)} style={{ padding: "0.4rem 0.8rem", fontSize: "0.8rem" }}>
                  このレジメンを反映
                </button>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Dictionaries ─────────────────────────────────────────────────
const COMMON_CANCERS = ["肺癌", "乳癌", "胃癌", "大腸癌", "肝癌", "膵癌", "前立腺癌", "悪性リンパ腫"];
const COMMON_DRUGS = ["オプジーボ", "キイトルーダ", "アバスチン", "パクリタキセル", "カルボプラチン", "シスプラチン"];

// ─── Upload Page ──────────────────────────────────────────────────
function UploadPage() {
  const { setUploadedFiles, setStep, setRegimenMaster, setLoading, setError, isLoading, error } = useAppStore();
  const [dragging, setDragging] = useState(false);
  const [files, setFiles] = useState<File[]>([]);
  const fileRef = useRef<HTMLInputElement>(null);

  const addFiles = (newFiles: File[]) => {
    setFiles((prev) => {
      const names = new Set(prev.map((f) => f.name));
      return [...prev, ...newFiles.filter((f) => !names.has(f.name))];
    });
  };

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    addFiles(Array.from(e.dataTransfer.files));
  }, []);

  const removeFile = (name: string) => setFiles((prev) => prev.filter((f) => f.name !== name));

  const handleParse = async () => {
    if (files.length === 0) return;
    setLoading(true);
    setError(null);
    setStep("extracting");

    try {
      // 動的インポートでフロント側のパーサーを読み込む
      const { parseExcelFile } = await import("@/lib/parsers/excelParser");
      const { parsePdfFile } = await import("@/lib/parsers/pdfParser");
      const { parseWordFile, extractBulletItemsFromParagraphs } = await import("@/lib/parsers/wordParser");
      const { extractBasicInfo } = await import("@/lib/extractors/basicInfoExtractor");
      const { extractRegimenBlocks } = await import("@/lib/extractors/regimenExtractor");
      const { extractReferencesFromText } = await import("@/lib/extractors/referenceExtractor");
      const { v4: uuidv4 } = await import("uuid");

      let combinedBasicInfo: Partial<RegimenMaster["basicInfo"]> = {};
      let combinedRegimenBlocks: RegimenBlock[] = [];
      let combinedTextBlocks: RichBulletItem[] = [];

      for (const file of files) {
        const buffer = await file.arrayBuffer();
        const ext = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();

        if (ext === ".xlsx" || ext === ".xls") {
          const sheetsData = await parseExcelFile(buffer);
          // 最初のシートを基本情報シートと仮定して渡す
          if (sheetsData.length > 0) {
            const basicInfoPartial = extractBasicInfo(sheetsData[0], file.name);
            combinedBasicInfo = { ...combinedBasicInfo, ...basicInfoPartial.value };
          }

          const blocks = extractRegimenBlocks(sheetsData, file.name);
          combinedRegimenBlocks = [...combinedRegimenBlocks, ...blocks];

          for (const sheet of sheetsData) {
            const cells = Array.from(sheet.cells.values());
            const textCells = cells.filter((c) => typeof c.value === "string" && c.value.length > 5);
            const items = textCells.map((c) => ({
              text: c.value,
              sourceTrace: [{
                sourceType: "excel" as const,
                fileName: file.name,
                sheetName: sheet.name,
                cellRange: c.address,
                quotedText: c.value.substring(0, 200),
                aiInterpretation: "Excelセルからの抽出テキスト",
                confidence: 0.8,
              }],
            }));
            combinedTextBlocks = [...combinedTextBlocks, ...items];
          }

        } else if (ext === ".pdf") {
          const result = await parsePdfFile(buffer, file.name);
          const paragraphs = result.pages.map((p) => ({ text: p }));
          combinedTextBlocks = [...combinedTextBlocks, ...paragraphs];

        } else if (ext === ".docx" || ext === ".doc") {
          const result = await parseWordFile(buffer, file.name);
          const paragraphs = result.paragraphs.map((p) => ({ text: p }));
          combinedTextBlocks = [...combinedTextBlocks, ...paragraphs];
        }
      }

      let extractedReferences: import("@/types/regimen").ReferenceItem[] = [];
      if (combinedTextBlocks.length > 0) {
        const fileExt = files[0]?.name.toLowerCase();
        const detectedType: import("@/types/regimen").FileType = 
          fileExt?.endsWith(".pdf") ? "pdf" : 
          fileExt?.endsWith(".docx") ? "word" : "excel";

        extractedReferences = extractReferencesFromText(
          combinedTextBlocks,
          "テキスト解析",
          detectedType
        );
      }

      const regimenData: RegimenMaster = {
        id: uuidv4(),
        templateVersion: "1.0",
        basicInfo: {
          applicationDate: combinedBasicInfo.applicationDate ?? null,
          applicantDoctor: combinedBasicInfo.applicantDoctor ?? null,
          departmentChief: combinedBasicInfo.departmentChief ?? null,
          cancerType: combinedBasicInfo.cancerType ?? null,
          regimenName: combinedBasicInfo.regimenName ?? null,
          courseLengthDays: combinedBasicInfo.courseLengthDays ?? null,
          totalCourses: combinedBasicInfo.totalCourses ?? null,
          treatmentPurpose: combinedBasicInfo.treatmentPurpose ?? { selected: null },
          category: combinedBasicInfo.category ?? null,
          contraindications: combinedBasicInfo.contraindications ?? [],
          eligibilityCriteria: combinedBasicInfo.eligibilityCriteria ?? [],
          stopCriteria: combinedBasicInfo.stopCriteria ?? [],
          doseReductionCriteria: combinedBasicInfo.doseReductionCriteria ?? [],
          precautions: combinedBasicInfo.precautions ?? [],
          popupNotes: combinedBasicInfo.popupNotes ?? [],
        },
        regimenBlocks: combinedRegimenBlocks,
        references: extractedReferences,
        sourceFiles: files.map(f => ({ 
          name: f.name, 
          type: (f.name.endsWith(".pdf") ? "pdf" : f.name.endsWith(".docx") ? "word" : "excel") as import("@/types/regimen").FileType, 
          url: null 
        })),
        createdAt: new Date().toISOString(),
      };

      setUploadedFiles(files);
      setRegimenMaster(regimenData);
      setStep("review-basic");
    } catch (err) {
      console.error(err);
      setError(String(err));
      setStep("upload");
    } finally {
      setLoading(false);
    }
  };

  const fileIcon = (name: string) => {
    if (/\.xlsx?$/i.test(name)) return "📊";
    if (/\.pdf$/i.test(name)) return "📕";
    if (/\.docx?$/i.test(name)) return "📝";
    return "📄";
  };

  return (
    <div className="fade-in">
      <div style={{ marginBottom: "2rem" }}>
        <h1 style={{ fontSize: "1.75rem", fontWeight: 700, color: "var(--text-primary)", marginBottom: "0.35rem" }}>
          レジメン情報の読み込み
        </h1>
        <p style={{ color: "var(--text-muted)", fontSize: "0.875rem" }}>
          テンプレートExcel・元資料（Excel / PDF / Word）をアップロードしてください
        </p>
      </div>

      <div
        className={`upload-zone ${dragging ? "dragging" : ""}`}
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={onDrop}
        onClick={() => fileRef.current?.click()}
      >
        <div style={{ fontSize: "3rem", marginBottom: "1rem" }}>☁️</div>
        <div style={{ fontSize: "1.1rem", fontWeight: 600, color: "var(--text-primary)", marginBottom: "0.4rem" }}>
          ファイルをドラッグ＆ドロップ
        </div>
        <div style={{ color: "var(--text-muted)", fontSize: "0.875rem" }}>
          .xlsx / .xls / .pdf / .docx に対応
        </div>
        <button className="btn btn-secondary" style={{ marginTop: "1.25rem" }} onClick={(e) => { e.stopPropagation(); fileRef.current?.click(); }}>
          ファイルを選択
        </button>
        <input
          ref={fileRef}
          type="file"
          multiple
          accept=".xlsx,.xls,.pdf,.docx,.doc"
          style={{ display: "none" }}
          onChange={(e) => addFiles(Array.from(e.target.files ?? []))}
        />
      </div>

      {files.length > 0 && (
        <div className="card fade-in" style={{ marginTop: "1.5rem" }}>
          <div className="section-header">
            <div className="section-title">選択済みソース</div>
            <span style={{ fontSize: "0.8rem", color: "var(--text-muted)" }}>{files.length} 件</span>
          </div>
          {files.map((f) => (
            <div key={f.name} style={{ display: "flex", alignItems: "center", gap: "0.75rem", padding: "0.6rem 0", borderBottom: "1px solid var(--border-subtle)" }}>
              <span style={{ fontSize: "1.25rem" }}>{fileIcon(f.name)}</span>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: "0.875rem", color: "var(--text-primary)" }}>{f.name}</div>
                <div style={{ fontSize: "0.72rem", color: "var(--text-muted)" }}>{(f.size / 1024).toFixed(1)} KB</div>
              </div>
              <button className="btn btn-danger" style={{ padding: "0.2rem 0.6rem", fontSize: "0.75rem" }} onClick={() => removeFile(f.name)}>削除</button>
            </div>
          ))}
          {error && <div style={{ color: "var(--danger)", fontSize: "0.85rem", marginTop: "1rem", padding: "0.75rem", background: "rgba(239,68,68,0.08)", borderRadius: 8 }}>⚠️ {error}</div>}
          <div style={{ marginTop: "1.25rem", display: "flex", justifyContent: "flex-end" }}>
            <button className="btn btn-primary" onClick={handleParse} disabled={isLoading}>
              {isLoading ? <><span className="spin" style={{ display: "inline-block" }}>⟳</span> 解析中...</> : "🔍 自動抽出を開始"}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Extracting Page ──────────────────────────────────────────────
function ExtractingPage() {
  return (
    <div className="fade-in" style={{ textAlign: "center", padding: "4rem 0" }}>
      <div style={{ fontSize: "3.5rem", marginBottom: "1.5rem" }}>
        <span className="spin" style={{ display: "inline-block" }}>⚙️</span>
      </div>
      <h2 style={{ fontSize: "1.4rem", fontWeight: 700, marginBottom: "0.5rem" }}>ファイルを解析中...</h2>
      <p style={{ color: "var(--text-muted)", fontSize: "0.875rem" }}>
        シート構造・結合セル・投与ステップを自動的に解析しています
      </p>
    </div>
  );
}

// ─── Basic Info Review ────────────────────────────────────────────
function BasicInfoReview() {
  const { regimenMaster, updateBasicInfo, setStep } = useAppStore();
  const [sourceModal, setSourceModal] = useState<import("@/types/regimen").SourceTrace[] | null>(null);
  const [historyModalOpen, setHistoryModalOpen] = useState(false);

  if (!regimenMaster) return null;
  const { basicInfo } = regimenMaster;

  const getStringField = (key: string): string => {
    const val = (basicInfo as unknown as Record<string, unknown>)[key];
    return typeof val === "string" ? val : typeof val === "number" ? String(val) : "";
  };

  const renderBullets = (items: RichBulletItem[], label: string, _field?: string) => (
    <div style={{ marginBottom: "1.25rem" }}>
      <label className="form-label">{label}</label>
      {items.map((item, i) => (
        <div key={i} className="bullet-item">
          <span style={{ color: "var(--accent)", marginTop: "0.1rem" }}>•</span>
          <span className="bullet-item-text">{item.text}</span>
          {item.sourceTrace && item.sourceTrace.length > 0 && (
            <button className="source-pill" onClick={() => setSourceModal(item.sourceTrace!)}>📍 出典</button>
          )}
        </div>
      ))}
      {items.length === 0 && (
        <div style={{ color: "var(--text-muted)", fontSize: "0.8rem", padding: "0.5rem 0" }}>（抽出なし）</div>
      )}
    </div>
  );

  return (
    <div className="fade-in">
      <datalist id="cancer-dictionary">
        {COMMON_CANCERS.map(c => <option key={c} value={c} />)}
      </datalist>
      <datalist id="drug-dictionary">
        {COMMON_DRUGS.map(c => <option key={c} value={c} />)}
      </datalist>

      {sourceModal && <SourceModal traces={sourceModal} onClose={() => setSourceModal(null)} />}
      {historyModalOpen && (
        <HistoryModal 
          onClose={() => setHistoryModalOpen(false)} 
          onSelect={(pastData) => {
            if (pastData.basicInfo && confirm("過去のレジメンの基本情報を反映しますか？（現在の内容は上書きされます）")) {
              Object.entries(pastData.basicInfo).forEach(([k, v]) => {
                const key = k as keyof import("@/types/regimen").BasicInfo;
                if (v !== undefined) {
                  updateBasicInfo({ [key]: v } as Partial<import("@/types/regimen").BasicInfo>);
                }
              });
            }
            setHistoryModalOpen(false);
          }} 
        />
      )}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: "2rem" }}>
        <div>
          <h2 style={{ fontSize: "1.4rem", fontWeight: 700 }}>基本情報レビュー</h2>
          <p style={{ color: "var(--text-muted)", fontSize: "0.875rem", marginTop: "0.25rem" }}>抽出された情報を確認・修正してください</p>
        </div>
        <button className="btn btn-secondary" onClick={() => setHistoryModalOpen(true)}>⏳ 過去のレジメンから反映</button>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "1.5rem" }}>
        {/* Left */}
        <div>
          <div className="card" style={{ marginBottom: "1.25rem" }}>
            <div className="section-title" style={{ marginBottom: "1rem" }}>申請情報</div>
            <div style={{ display: "grid", gap: "0.875rem" }}>
              {[
                { label: "申請日", key: "applicationDate" },
                { label: "申請医師", key: "applicantDoctor" },
                { label: "診療科長", key: "departmentChief" },
                { label: "がん種", key: "cancerType", list: "cancer-dictionary" },
                { label: "レジメン名", key: "regimenName" },
              ].map(({ label, key, list }) => (
                <div key={key}>
                  <label className="form-label">{label}</label>
                  <input
                    className="form-input"
                    list={list}
                    value={getStringField(key)}
                    onChange={(e) => updateBasicInfo({ [key]: e.target.value } as Partial<RegimenMaster["basicInfo"]>)}
                  />
                </div>
              ))}
              <div>
                <label className="form-label">コース日数</label>
                <input
                  className="form-input"
                  type="number"
                  value={basicInfo.courseLengthDays ?? ""}
                  onChange={(e) => updateBasicInfo({ courseLengthDays: e.target.value ? Number(e.target.value) : null })}
                />
              </div>
              <div>
                <label className="form-label">総コース数</label>
                <input
                  className="form-input"
                  value={basicInfo.totalCourses ?? ""}
                  onChange={(e) => updateBasicInfo({ totalCourses: e.target.value })}
                />
              </div>
            </div>
          </div>

          <div className="card">
            <div className="section-title" style={{ marginBottom: "1rem" }}>区分・目的</div>
            <div style={{ display: "grid", gap: "0.875rem" }}>
              <div>
                <label className="form-label">区分</label>
                <select
                  className="form-select"
                  value={basicInfo.category ?? ""}
                  onChange={(e) => updateBasicInfo({ category: (e.target.value || null) as RegimenMaster["basicInfo"]["category"] })}
                >
                  <option value="">選択してください</option>
                  <option>入院</option>
                  <option>外来</option>
                  <option>入院・外来</option>
                </select>
              </div>
              <div>
                <label className="form-label">治療目的</label>
                <select
                  className="form-select"
                  value={basicInfo.treatmentPurpose.selected ?? ""}
                  onChange={(e) => updateBasicInfo({
                    treatmentPurpose: {
                      ...basicInfo.treatmentPurpose,
                      selected: (e.target.value || null) as RegimenMaster["basicInfo"]["treatmentPurpose"]["selected"],
                    }
                  })}
                >
                  <option value="">選択してください</option>
                  <option>進行再発</option>
                  <option>術前補助化学療法</option>
                  <option>術後補助化学療法</option>
                  <option>その他</option>
                </select>
              </div>
              {basicInfo.treatmentPurpose.selected === "その他" && (
                <div>
                  <label className="form-label">その他（自由記載）</label>
                  <input
                    className="form-input"
                    value={basicInfo.treatmentPurpose.otherText ?? ""}
                    onChange={(e) => updateBasicInfo({
                      treatmentPurpose: { ...basicInfo.treatmentPurpose, otherText: e.target.value }
                    })}
                  />
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Right */}
        <div>
          <div className="card">
            <div className="section-title" style={{ marginBottom: "1rem" }}>基準・注意事項</div>
            {renderBullets(basicInfo.contraindications, "禁忌")}
            {renderBullets(basicInfo.eligibilityCriteria, "投与基準")}
            {renderBullets(basicInfo.stopCriteria, "中止基準")}
            {renderBullets(basicInfo.doseReductionCriteria, "減量基準")}
            {renderBullets(basicInfo.precautions, "注意事項")}
          </div>
        </div>
      </div>

      <div style={{ display: "flex", justifyContent: "space-between", marginTop: "2rem" }}>
        <button className="btn btn-secondary" onClick={() => setStep("upload")}>← 戻る</button>
        <button className="btn btn-primary" onClick={() => setStep("review-regimen")}>投与方法レビューへ →</button>
      </div>
    </div>
  );
}

// ─── Regimen Review ────────────────────────────────────────────────
function RegimenReview() {
  const { regimenMaster, setStep, setRegimenMaster } = useAppStore();
  const [sourceModal, setSourceModal] = useState<import("@/types/regimen").SourceTrace[] | null>(null);
  const [expandedBlock, setExpandedBlock] = useState<string | null>(null);

  if (!regimenMaster) return null;

  const updateBlock = (blockId: string, patch: Partial<RegimenBlock>) => {
    setRegimenMaster({
      ...regimenMaster,
      regimenBlocks: regimenMaster.regimenBlocks.map((b) =>
        b.regimenId === blockId ? { ...b, ...patch } : b
      ),
    });
  };

  const addBlock = () => {
    const newBlock: RegimenBlock = {
      regimenId: crypto.randomUUID(),
      regimenLabel: `レジメン${regimenMaster.regimenBlocks.length + 1}`,
      administrationSteps: [],
    };
    setRegimenMaster({ ...regimenMaster, regimenBlocks: [...regimenMaster.regimenBlocks, newBlock] });
  };

  const removeBlock = (id: string) => {
    setRegimenMaster({
      ...regimenMaster,
      regimenBlocks: regimenMaster.regimenBlocks.filter((b) => b.regimenId !== id),
    });
  };

  const renderStep = (step: AdministrationStep, blockId: string) => (
    <div key={step.stepNo} className="step-row">
      <div style={{ display: "flex", gap: "0.75rem", alignItems: "center", flexWrap: "wrap", marginBottom: "0.5rem" }}>
        <span style={{
          background: "var(--accent)", color: "white",
          borderRadius: 6, width: 24, height: 24, display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: "0.75rem", fontWeight: 700, flexShrink: 0,
        }}>{step.stepNo}</span>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: "0.875rem", fontWeight: 600, color: "var(--text-primary)" }}>
            {step.mainDrugs.map((d) => d.drugName).join(" / ") || "（薬剤未登録）"}
          </div>
          <div style={{ fontSize: "0.75rem", color: "var(--text-muted)", marginTop: "0.1rem" }}>
            {[
              step.mainDrugs.map((d) => `${d.dose ?? ""}${d.doseUnit ?? ""}`).join(", "),
              step.route && `投与経路: ${step.route}`,
              step.infusionRate && `速度: ${step.infusionRate}${step.rateUnit ?? ""}`,
              step.administrationDays.length > 0 && `Day: ${step.administrationDays.join(", ")}`,
            ].filter(Boolean).join(" | ")}
          </div>
        </div>
        {step.sourceTrace && step.sourceTrace.length > 0 && (
          <button className="source-pill" onClick={() => setSourceModal(step.sourceTrace!)}>📍 出典</button>
        )}
      </div>
      {step.diluents.length > 0 && (
        <div style={{ fontSize: "0.75rem", color: "var(--text-muted)", paddingLeft: "2rem" }}>
          💧 溶解液: {step.diluents.map((d) => `${d.name} ${d.volume ?? ""}${d.unit ?? ""}`).join(", ")}
        </div>
      )}
    </div>
  );

  return (
    <div className="fade-in">
      {sourceModal && <SourceModal traces={sourceModal} onClose={() => setSourceModal(null)} />}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: "2rem" }}>
        <div>
          <h2 style={{ fontSize: "1.4rem", fontWeight: 700 }}>投与方法レビュー</h2>
          <p style={{ color: "var(--text-muted)", fontSize: "0.875rem", marginTop: "0.25rem" }}>
            レジメンブロックと投与ステップを確認・編集してください
          </p>
        </div>
        <button className="btn btn-secondary" onClick={addBlock}>＋ ブロック追加</button>
      </div>

      {regimenMaster.regimenBlocks.length === 0 && (
        <div className="card" style={{ textAlign: "center", padding: "3rem", color: "var(--text-muted)" }}>
          レジメンブロックが見つかりません。「ブロック追加」から手動で追加してください。
        </div>
      )}

      {regimenMaster.regimenBlocks.map((block) => (
        <div key={block.regimenId} className="regimen-block-card">
          <div className="regimen-block-header">
            <span style={{
              background: "linear-gradient(135deg, var(--accent), var(--accent-secondary))",
              color: "white", borderRadius: 8, padding: "0.2rem 0.6rem", fontSize: "0.75rem", fontWeight: 700,
            }}>Block</span>
            <input
              className="form-input"
              style={{ flex: 1, maxWidth: 300 }}
              value={block.regimenLabel}
              onChange={(e) => updateBlock(block.regimenId, { regimenLabel: e.target.value })}
            />
            <input
              className="form-input"
              style={{ flex: 1, maxWidth: 200 }}
              placeholder="コース範囲（例：1コース目）"
              value={block.applicableCourseRange ?? ""}
              onChange={(e) => updateBlock(block.regimenId, { applicableCourseRange: e.target.value || null })}
            />
            <button
              style={{ background: "none", border: "none", cursor: "pointer", fontSize: "1rem", color: "var(--text-muted)", padding: "0.25rem" }}
              onClick={() => setExpandedBlock(expandedBlock === block.regimenId ? null : block.regimenId)}
            >{expandedBlock === block.regimenId ? "▲" : "▼"}</button>
            <button className="btn btn-danger" style={{ padding: "0.25rem 0.6rem", fontSize: "0.75rem" }} onClick={() => removeBlock(block.regimenId)}>削除</button>
          </div>

          {expandedBlock === block.regimenId && (
            <div className="regimen-block-body">
              {block.sourceSheetName && (
                <div style={{ fontSize: "0.75rem", color: "var(--text-muted)", marginBottom: "0.75rem" }}>
                  🗂 ソースシート: {block.sourceSheetName}
                </div>
              )}
              {block.administrationSteps.length === 0 ? (
                <div style={{ color: "var(--text-muted)", fontSize: "0.85rem", padding: "1rem 0" }}>投与ステップが検出されませんでした</div>
              ) : (
                block.administrationSteps.map((s) => renderStep(s, block.regimenId))
              )}
            </div>
          )}
        </div>
      ))}

      <div style={{ display: "flex", justifyContent: "space-between", marginTop: "2rem" }}>
        <button className="btn btn-secondary" onClick={() => setStep("review-basic")}>← 基本情報へ</button>
        <button className="btn btn-primary" onClick={() => setStep("export")}>出力画面へ →</button>
      </div>
    </div>
  );
}

// ─── Export Page ──────────────────────────────────────────────────
function ExportPage() {
  const { regimenMaster, setStep } = useAppStore();
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [isExporting, setIsExporting] = useState(false);
  const [downloadLinks, setDownloadLinks] = useState<{ excel?: string; json?: string; auditLog?: string } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);

  if (!regimenMaster) return null;

  const handleExport = async () => {
    if (!templateFile) return;
    setIsExporting(true);
    setError(null);

    try {
      // 履歴をローカルストレージに保存
      try {
        const stored = localStorage.getItem("regimen_history");
        const history: RegimenMaster[] = stored ? JSON.parse(stored) : [];
        const newHistory = [regimenMaster, ...history.filter(h => h.id !== regimenMaster.id)].slice(0, 50);
        localStorage.setItem("regimen_history", JSON.stringify(newHistory));
      } catch (e) {
        console.warn("ローカル履歴の保存に失敗:", e);
      }

      // JSONとAuditLogのBlob生成
      const jsonBlob = new Blob([JSON.stringify(regimenMaster, null, 2)], { type: "application/json" });
      const jsonUrl = URL.createObjectURL(jsonBlob);

      const auditLog = {
        action: "REGIMEN_FINALIZED",
        timestamp: new Date().toISOString(),
        regimenId: regimenMaster.id,
        summary: `${regimenMaster.basicInfo.regimenName} (Blocks: ${regimenMaster.regimenBlocks.length}, Validation: N/A)`,
      };
      const auditBlob = new Blob([JSON.stringify(auditLog, null, 2)], { type: "application/json" });
      const auditUrl = URL.createObjectURL(auditBlob);

      // 動的インポートでExcelJSとエクスポート処理を呼び出す
      const ExcelJS = (await import("exceljs")).default;
      
      const buffer = await templateFile.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      // Template writing logic (simplified for browser environment)
      const sheet = workbook.worksheets[0];
      if (sheet) {
        if (regimenMaster.basicInfo.regimenName) sheet.getCell("B2").value = regimenMaster.basicInfo.regimenName;
        if (regimenMaster.basicInfo.cancerType) sheet.getCell("B3").value = regimenMaster.basicInfo.cancerType;
        if (regimenMaster.basicInfo.treatmentPurpose?.selected) sheet.getCell("B4").value = regimenMaster.basicInfo.treatmentPurpose.selected;
        if (regimenMaster.basicInfo.category) sheet.getCell("B5").value = regimenMaster.basicInfo.category;
        sheet.getCell("B6").value = `Total ${regimenMaster.regimenBlocks.length} blocks`;
        sheet.getCell("B7").value = "Generated by browser (Static Export)";
      }

      // BufferをBlobに変換してダウンロードリンク生成
      const outBuffer = await workbook.xlsx.writeBuffer();
      const excelBlob = new Blob([outBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const excelUrl = URL.createObjectURL(excelBlob);

      setDownloadLinks({ excel: excelUrl, json: jsonUrl, auditLog: auditUrl });
    } catch (e) {
      console.error(e);
      setError("エクスポート中にエラーが発生しました");
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="fade-in">
      <div style={{ marginBottom: "2rem" }}>
        <h2 style={{ fontSize: "1.4rem", fontWeight: 700 }}>出力</h2>
        <p style={{ color: "var(--text-muted)", fontSize: "0.875rem", marginTop: "0.25rem" }}>
          Excelテンプレートに書き戻し、正規化JSON・監査ログをダウンロードします
        </p>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "1.5rem" }}>
        <div className="card">
          <div className="section-title" style={{ marginBottom: "1rem" }}>確認済みデータ</div>
          <div style={{ display: "grid", gap: "0.5rem" }}>
            {[
              ["レジメン名", regimenMaster.basicInfo.regimenName ?? "（未記入）"],
              ["がん種", regimenMaster.basicInfo.cancerType ?? "（未記入）"],
              ["区分", regimenMaster.basicInfo.category ?? "（未記入）"],
              ["治療目的", regimenMaster.basicInfo.treatmentPurpose.selected ?? "（未記入）"],
              ["レジメンブロック数", `${regimenMaster.regimenBlocks.length} ブロック`],
              ["総投与ステップ数", `${regimenMaster.regimenBlocks.reduce((acc, b) => acc + b.administrationSteps.length, 0)} ステップ`],
            ].map(([k, v]) => (
              <div key={k} style={{ display: "flex", gap: "0.75rem", fontSize: "0.875rem", padding: "0.4rem 0", borderBottom: "1px solid var(--border-subtle)" }}>
                <span style={{ color: "var(--text-muted)", minWidth: 160 }}>{k}</span>
                <span style={{ color: "var(--text-primary)", fontWeight: 500 }}>{v}</span>
              </div>
            ))}
          </div>
        </div>

        <div className="card">
          <div className="section-title" style={{ marginBottom: "1rem" }}>書き戻し設定</div>
          <label className="form-label">テンプレートExcelを選択</label>
          <div
            style={{ border: "2px dashed var(--border)", borderRadius: 10, padding: "1.25rem", textAlign: "center", cursor: "pointer", marginBottom: "1rem" }}
            onClick={() => fileRef.current?.click()}
          >
            {templateFile ? (
              <div style={{ color: "var(--success)" }}>✅ {templateFile.name}</div>
            ) : (
              <div style={{ color: "var(--text-muted)", fontSize: "0.875rem" }}>クリックして選択</div>
            )}
          </div>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
            onChange={(e) => setTemplateFile(e.target.files?.[0] ?? null)} />

          {error && <div style={{ color: "var(--danger)", fontSize: "0.85rem", marginBottom: "1rem" }}>⚠️ {error}</div>}

          <button
            className="btn btn-primary"
            style={{ width: "100%", justifyContent: "center" }}
            onClick={handleExport}
            disabled={!templateFile || isExporting}
          >
            {isExporting ? "⟳ 書き戻し中..." : "📤 Excelに書き戻す"}
          </button>

          {downloadLinks && (
            <div style={{ marginTop: "1.25rem", display: "grid", gap: "0.5rem" }}>
              <a href={downloadLinks.excel} className="btn btn-success" style={{ justifyContent: "center" }}>📊 output.xlsx をダウンロード</a>
              <a href={downloadLinks.json} className="btn btn-secondary" style={{ justifyContent: "center" }}>📋 normalized.json</a>
              <a href={downloadLinks.auditLog} className="btn btn-secondary" style={{ justifyContent: "center" }}>🔍 audit-log.json</a>
            </div>
          )}
        </div>
      </div>

      <div style={{ display: "flex", justifyContent: "space-between", marginTop: "2rem" }}>
        <button className="btn btn-secondary" onClick={() => setStep("review-regimen")}>← 投与方法へ</button>
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────
export default function Home() {
  const { step } = useAppStore();

  const renderPage = () => {
    switch (step) {
      case "upload": return <UploadPage />;
      case "extracting": return <ExtractingPage />;
      case "review-basic": return <BasicInfoReview />;
      case "review-regimen": return <RegimenReview />;
      case "export": return <ExportPage />;
      default: return <UploadPage />;
    }
  };

  return (
    <>
      <ProgressBar current={step} />
      {renderPage()}
    </>
  );
}
