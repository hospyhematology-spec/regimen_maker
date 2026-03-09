import { NextResponse } from "next/server";
import { getDb } from "@/lib/db";

export async function GET() {
  try {
    const db = getDb();
    // 簡易的にレジメンマスターの一覧を取得
    const rows = db.prepare(`
      SELECT r.id, b.regimen_name, b.cancer_type, r.created_at
      FROM regimen_master r
      LEFT JOIN basic_info b ON r.id = b.regimen_id
      ORDER BY r.created_at DESC
      LIMIT 20
    `).all();

    return NextResponse.json({ success: true, regimens: rows });
  } catch (e) {
    console.error("Failed to fetch regimens:", e);
    return NextResponse.json({ error: String(e) }, { status: 500 });
  }
}
