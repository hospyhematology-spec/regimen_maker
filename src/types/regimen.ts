// Regimen Maker - 共通型定義
// JSON Schema から派生した TypeScript 型

export type FileType = "excel" | "pdf" | "word" | "web" | "text";

export type TreatmentPurpose =
  | "進行再発"
  | "術前補助化学療法"
  | "術後補助化学療法"
  | "その他";

export type Category = "入院" | "外来" | "入院・外来";

// -------------------------------------------------------
// 出典トレース
// -------------------------------------------------------
export interface SourceTrace {
  sourceType: FileType;
  fileName?: string | null;
  sheetName?: string | null;
  cellRange?: string | null;
  page?: number | null;
  paragraphIndex?: number | null;
  url?: string | null;
  quotedText: string;
  aiInterpretation: string;
  confidence: number; // 0.0 〜 1.0
}

// -------------------------------------------------------
// 箇条書き項目（出典付き）
// -------------------------------------------------------
export interface RichBulletItem {
  text: string;
  sourceTrace?: SourceTrace[];
}

// -------------------------------------------------------
// 入力ソースファイル
// -------------------------------------------------------
export interface SourceFile {
  name: string;
  type: FileType;
  url?: string | null;
}

// -------------------------------------------------------
// 基本情報
// -------------------------------------------------------
export interface BasicInfo {
  applicationDate?: string | null;
  applicantDoctor?: string | null;
  departmentChief?: string | null;
  cancerType?: string | null;
  regimenName?: string | null;
  courseLengthDays?: number | null;
  totalCourses?: string | null;
  treatmentPurpose: {
    selected: TreatmentPurpose | null;
    otherText?: string | null;
  };
  category: Category | null;
  contraindications: RichBulletItem[];
  eligibilityCriteria: RichBulletItem[];
  stopCriteria: RichBulletItem[];
  doseReductionCriteria: RichBulletItem[];
  precautions: RichBulletItem[];
  popupNotes: RichBulletItem[];
}

// -------------------------------------------------------
// 主薬 / 溶解液 / フラッシュ液
// -------------------------------------------------------
export interface DrugComponent {
  drugName: string;
  dose?: string | null;
  doseUnit?: string | null;
}

export interface DiluentItem {
  name: string;
  volume?: string | null;
  unit?: string | null;
}

export interface FlushItem {
  name: string;
  volume?: string | null;
  unit?: string | null;
}

// -------------------------------------------------------
// 投与ステップ
// -------------------------------------------------------
export interface AdministrationStep {
  stepNo: number;
  mainDrugs: DrugComponent[];
  diluents: DiluentItem[];
  flushes: FlushItem[];
  route?: string | null;
  infusionRate?: string | null;
  rateUnit?: string | null;
  administrationDays: string[];
  notes: RichBulletItem[];
  sourceTrace?: SourceTrace[];
}

// -------------------------------------------------------
// 経口投与 / 補足指示
// -------------------------------------------------------
export interface SupplementaryOrder {
  orderId: string;
  instructionText: string;
  notes?: RichBulletItem[];
  sourceTrace?: SourceTrace[];
}

// -------------------------------------------------------
// レジメンブロック
// -------------------------------------------------------
export interface RegimenBlock {
  regimenId: string;
  regimenLabel: string;
  applicableCourseRange?: string | null;
  sourceSheetName?: string | null;
  administrationSteps: AdministrationStep[];
  oralOrders?: SupplementaryOrder[];
}

// -------------------------------------------------------
// 代表論文
// -------------------------------------------------------
export interface ReferenceItem {
  title: string;
  firstAuthor?: string | null;
  journal?: string | null;
  year?: number | null;
  doiOrUrl?: string | null;
  representative: boolean;
  sourceTrace?: SourceTrace[];
}

// -------------------------------------------------------
// 監査情報
// -------------------------------------------------------
export interface AuditInfo {
  createdAt: string; // ISO8601
  updatedAt: string; // ISO8601
  operator?: string | null;
  changes: string[];
}

// -------------------------------------------------------
// トップレベル
// -------------------------------------------------------
export interface RegimenMaster {
  id: string;
  templateVersion: string;
  sourceFiles: SourceFile[];
  basicInfo: BasicInfo;
  regimenBlocks: RegimenBlock[];
  supplementaryOrders?: SupplementaryOrder[];
  references: ReferenceItem[];
  audit?: AuditInfo;
}

// -------------------------------------------------------
// UI 用候補値ラッパー（信頼度・出典付き）
// -------------------------------------------------------
export interface CandidateValue<T> {
  value: T;
  confidence: number;
  sourceTrace: SourceTrace[];
}

export interface ExtractionCandidate {
  basicInfo: Partial<{
    [K in keyof BasicInfo]: CandidateValue<BasicInfo[K]>;
  }>;
  regimenBlocks: CandidateValue<RegimenBlock>[];
}
