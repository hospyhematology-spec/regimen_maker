import Database from "better-sqlite3";
import path from "path";
import fs from "fs";

const DB_DIR = path.join(process.cwd(), "data");
const DB_PATH = path.join(DB_DIR, "regimen.db");

let _db: Database.Database | null = null;

export function getDb(): Database.Database {
  if (_db) return _db;
  if (!fs.existsSync(DB_DIR)) fs.mkdirSync(DB_DIR, { recursive: true });
  _db = new Database(DB_PATH);
  _db.pragma("journal_mode = WAL");
  initSchema(_db);
  return _db;
}

function initSchema(db: Database.Database) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS regimen_master (
      id TEXT PRIMARY KEY,
      template_version TEXT NOT NULL,
      cancer_type TEXT,
      regimen_name TEXT,
      category TEXT,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS source_file (
      id TEXT PRIMARY KEY,
      regimen_master_id TEXT NOT NULL REFERENCES regimen_master(id) ON DELETE CASCADE,
      file_name TEXT NOT NULL,
      file_type TEXT NOT NULL,
      url TEXT
    );

    CREATE TABLE IF NOT EXISTS basic_info (
      regimen_master_id TEXT PRIMARY KEY REFERENCES regimen_master(id) ON DELETE CASCADE,
      application_date TEXT,
      applicant_doctor TEXT,
      department_chief TEXT,
      course_length_days INTEGER,
      total_courses TEXT,
      treatment_purpose TEXT,
      purpose_other TEXT,
      category TEXT
    );

    CREATE TABLE IF NOT EXISTS bullet_item (
      id TEXT PRIMARY KEY,
      regimen_master_id TEXT NOT NULL REFERENCES regimen_master(id) ON DELETE CASCADE,
      field_type TEXT NOT NULL,
      order_no INTEGER NOT NULL,
      text TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS regimen_block (
      id TEXT PRIMARY KEY,
      regimen_master_id TEXT NOT NULL REFERENCES regimen_master(id) ON DELETE CASCADE,
      label TEXT NOT NULL,
      course_range TEXT,
      source_sheet_name TEXT,
      sort_order INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS administration_step (
      id TEXT PRIMARY KEY,
      regimen_block_id TEXT NOT NULL REFERENCES regimen_block(id) ON DELETE CASCADE,
      step_no INTEGER NOT NULL,
      route TEXT,
      infusion_rate TEXT,
      rate_unit TEXT
    );

    CREATE TABLE IF NOT EXISTS drug_component (
      id TEXT PRIMARY KEY,
      administration_step_id TEXT NOT NULL REFERENCES administration_step(id) ON DELETE CASCADE,
      component_type TEXT NOT NULL,
      drug_name TEXT NOT NULL,
      dose TEXT,
      dose_unit TEXT,
      volume TEXT,
      volume_unit TEXT
    );

    CREATE TABLE IF NOT EXISTS step_day (
      id TEXT PRIMARY KEY,
      administration_step_id TEXT NOT NULL REFERENCES administration_step(id) ON DELETE CASCADE,
      day_label TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS source_trace (
      id TEXT PRIMARY KEY,
      owner_type TEXT NOT NULL,
      owner_id TEXT NOT NULL,
      source_type TEXT NOT NULL,
      file_name TEXT,
      sheet_name TEXT,
      cell_range TEXT,
      page INTEGER,
      paragraph_index INTEGER,
      url TEXT,
      quoted_text TEXT NOT NULL,
      ai_interpretation TEXT NOT NULL,
      confidence REAL NOT NULL
    );

    CREATE TABLE IF NOT EXISTS reference_item (
      id TEXT PRIMARY KEY,
      regimen_master_id TEXT NOT NULL REFERENCES regimen_master(id) ON DELETE CASCADE,
      title TEXT NOT NULL,
      first_author TEXT,
      journal TEXT,
      year INTEGER,
      doi_or_url TEXT,
      representative INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS audit_log (
      id TEXT PRIMARY KEY,
      regimen_master_id TEXT NOT NULL REFERENCES regimen_master(id) ON DELETE CASCADE,
      action TEXT NOT NULL,
      operator TEXT,
      created_at TEXT NOT NULL,
      before_json TEXT,
      after_json TEXT
    );
  `);
}
