import { NextRequest, NextResponse } from "next/server";
import { getDb } from "@/lib/db";
import type { RegimenMaster } from "@/types/regimen";

export async function GET(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const db = getDb();
    
    // 主テーブルから抽出
    const masterRow = db.prepare(`SELECT * FROM regimen_master WHERE id = ?`).get(id) as any;
    if (!masterRow) return NextResponse.json({ error: "Not found" }, { status: 404 });

    const basicInfoRow = db.prepare(`SELECT * FROM basic_info WHERE regimen_id = ?`).get(id) as any;
    const blockRows = db.prepare(`SELECT * FROM regimen_block WHERE regimen_id = ?`).all(id) as any[];

    // 今回は簡易比較のためのデータ構造としてJSON化する（本来は完全な復元が必要）
    // UI側の差分表示用モックデータとして最低限の形を返す
    const regimenData: Partial<RegimenMaster> = {
      id: masterRow.id,
      basicInfo: {
        applicationDate: basicInfoRow?.application_date,
        applicantDoctor: basicInfoRow?.applicant_doctor,
        departmentChief: basicInfoRow?.department_chief,
        cancerType: basicInfoRow?.cancer_type,
        regimenName: basicInfoRow?.regimen_name,
        courseLengthDays: basicInfoRow?.course_length_days,
        totalCourses: basicInfoRow?.total_courses,
        treatmentPurpose: { selected: basicInfoRow?.treatment_purpose },
        category: basicInfoRow?.category,
        contraindications: [],
        eligibilityCriteria: [],
        stopCriteria: [],
        doseReductionCriteria: [],
        precautions: [],
        popupNotes: [],
      },
      regimenBlocks: blockRows.map(b => ({
        regimenId: b.block_id,
        regimenLabel: b.regimen_label,
        applicableCourseRange: b.applicable_course_range,
        administrationSteps: [],
      })),
    };

    return NextResponse.json({ success: true, data: regimenData });
  } catch (e) {
    console.error("Failed to fetch regimen detail:", e);
    return NextResponse.json({ error: String(e) }, { status: 500 });
  }
}
