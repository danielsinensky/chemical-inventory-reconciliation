-- Schema for the EH&S inventory database (SQLite)
CREATE TABLE labs (lab_id TEXT PRIMARY KEY, pi_name TEXT NOT NULL, building TEXT NOT NULL);
CREATE TABLE rooms (room_id TEXT PRIMARY KEY, lab_id TEXT NOT NULL REFERENCES labs(lab_id), building TEXT, room_number TEXT);
CREATE TABLE people (person_id TEXT PRIMARY KEY, lab_id TEXT NOT NULL REFERENCES labs(lab_id), name TEXT NOT NULL, email TEXT, role TEXT, last_confirmed TEXT);
CREATE TABLE chemicals (cas_number TEXT PRIMARY KEY, chemical_name TEXT NOT NULL, hazard_class TEXT);
CREATE TABLE inventory (container_id TEXT PRIMARY KEY, lab_id TEXT NOT NULL REFERENCES labs(lab_id), room_id TEXT REFERENCES rooms(room_id), cas_number TEXT REFERENCES chemicals(cas_number), quantity REAL, unit TEXT, last_verified TEXT);
CREATE TABLE inspection_findings (finding_id TEXT PRIMARY KEY, lab_id TEXT NOT NULL REFERENCES labs(lab_id), date_found TEXT, category TEXT, severity TEXT, status TEXT, date_closed TEXT);
CREATE TABLE app_users (user_id INTEGER PRIMARY KEY, person_id TEXT NOT NULL REFERENCES people(person_id), role TEXT NOT NULL CHECK (role IN ('Viewer', 'Lab Manager', 'EH&S Admin')), is_active INTEGER NOT NULL DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE audit_log (log_id INTEGER PRIMARY KEY, action TEXT NOT NULL, detail TEXT, done_by TEXT NOT NULL, done_at TEXT DEFAULT CURRENT_TIMESTAMP);
