import { NextRequest, NextResponse } from "next/server";
import { readFile } from "fs/promises";
import path from "path";

const OUTPUT_DIR = path.join(process.cwd(), "data", "outputs");

export async function GET(req: NextRequest) {
  const { searchParams } = new URL(req.url);
  const file = searchParams.get("file");
  if (!file) {
    return NextResponse.json({ error: "fileパラメータが必要です" }, { status: 400 });
  }

  const filePath = path.join(OUTPUT_DIR, path.basename(file));
  try {
    const buffer = await readFile(filePath);
    const ext = path.extname(file).toLowerCase();
    const contentType =
      ext === ".xlsx"
        ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        : ext === ".json"
        ? "application/json"
        : "application/octet-stream";

    return new NextResponse(buffer, {
      headers: {
        "Content-Type": contentType,
        "Content-Disposition": `attachment; filename="${encodeURIComponent(path.basename(file))}"`,
      },
    });
  } catch {
    return NextResponse.json({ error: "ファイルが見つかりません" }, { status: 404 });
  }
}
