import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Regimen Maker | レジメン作成支援システム",
  description: "Excel・PDF・Wordから抗がん剤レジメン情報を自動抽出し、院内テンプレートに書き戻すシステム",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ja">
      <head>
        <link rel="preconnect" href="https://fonts.googleapis.com" />
        <link rel="preconnect" href="https://fonts.gstatic.com" crossOrigin="anonymous" />
      </head>
      <body>
        <div className="page-container">
          <header style={{
            background: "rgba(15,17,23,0.85)",
            backdropFilter: "blur(12px)",
            borderBottom: "1px solid var(--border-subtle)",
            padding: "0 1.5rem",
            position: "sticky", top: 0, zIndex: 50,
          }}>
            <div style={{ maxWidth: 1400, margin: "0 auto", height: 60, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div style={{ display: "flex", alignItems: "center", gap: "0.75rem" }}>
                <div style={{
                  width: 32, height: 32, borderRadius: 10,
                  background: "linear-gradient(135deg, var(--accent) 0%, var(--accent-secondary) 100%)",
                  display: "flex", alignItems: "center", justifyContent: "center",
                  fontSize: "1rem"
                }}>💊</div>
                <div>
                  <div style={{ fontSize: "1rem", fontWeight: 700, color: "var(--text-primary)" }}>Regimen Maker</div>
                  <div style={{ fontSize: "0.65rem", color: "var(--text-muted)", marginTop: -2 }}>レジメン作成支援システム</div>
                </div>
              </div>
              <div style={{ fontSize: "0.72rem", color: "var(--text-muted)" }}>MVP-1</div>
            </div>
          </header>
          <main className="main-content">
            {children}
          </main>
        </div>
      </body>
    </html>
  );
}
