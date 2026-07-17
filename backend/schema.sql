CREATE TABLE IF NOT EXISTS components (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  item_id TEXT,
  name TEXT NOT NULL UNIQUE,
  year INTEGER,
  month_index INTEGER,
  month_name TEXT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS sub_components (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  item_id TEXT,
  component_id INTEGER NOT NULL,
  name TEXT NOT NULL,
  amount_value REAL DEFAULT 0,
  amount_cell TEXT,
  year INTEGER,
  month_index INTEGER,
  month_name TEXT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (component_id) REFERENCES components(id)
);

CREATE TABLE IF NOT EXISTS budget_submissions (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  year INTEGER NOT NULL,
  month_index INTEGER NOT NULL,
  month_name TEXT NOT NULL,
  submitted_by TEXT NOT NULL,
  submitted_at TEXT DEFAULT CURRENT_TIMESTAMP,
  UNIQUE(year, month_index)
);

CREATE TABLE IF NOT EXISTS ceo_submission_views (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  submission_id INTEGER NOT NULL,
  ceo_username TEXT NOT NULL,
  viewed_at TEXT DEFAULT CURRENT_TIMESTAMP,
  UNIQUE(submission_id, ceo_username),
  FOREIGN KEY (submission_id) REFERENCES budget_submissions(id)
);

CREATE TABLE IF NOT EXISTS budget_submission_components (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  submission_id INTEGER NOT NULL,
  component_id INTEGER NOT NULL,
  component_name TEXT NOT NULL,
  UNIQUE(submission_id, component_id),
  FOREIGN KEY (submission_id) REFERENCES budget_submissions(id)
);

CREATE TABLE IF NOT EXISTS budget_submission_rows (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  submission_id INTEGER NOT NULL,
  sub_component_id INTEGER NOT NULL,
  component_id INTEGER NOT NULL,
  review_status TEXT NOT NULL DEFAULT 'Pending',
  reviewed_at TEXT,
  UNIQUE(submission_id, sub_component_id),
  FOREIGN KEY (submission_id) REFERENCES budget_submissions(id)
);
