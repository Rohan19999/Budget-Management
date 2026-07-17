
const express = require("express");
const cors = require("cors");
const sqlite3 = require("sqlite3").verbose();
const path = require("path");
const crypto = require("crypto");
const multer = require("multer");
const ExcelJS = require("exceljs");

const app = express();
const PORT = process.env.PORT || 5050;
const dbPath = process.env.DB_PATH || path.join(__dirname, "finance.db");
const upload = multer({ storage: multer.memoryStorage() });

app.use(cors());
app.use(express.json());

function hashPassword(password) {
  return crypto.createHash("sha256").update(String(password)).digest("hex");
}

const db = new sqlite3.Database(dbPath, (err) => {
  if (err) {
    console.error("Database connection failed:", err.message);
    process.exit(1);
  }
  console.log(`Connected to database at: ${dbPath}`);
});

db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS components (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      item_id TEXT,
      name TEXT NOT NULL UNIQUE,
      year INTEGER,
      month_index INTEGER,
      month_name TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS sub_components (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      item_id TEXT,
      component_id INTEGER NOT NULL,
      name TEXT NOT NULL,
      amount_value REAL DEFAULT 0,
      amount_cell TEXT,
      existing_liability REAL DEFAULT 0,
      estimated_liability REAL DEFAULT 0,
      total_estimated_liability REAL DEFAULT 0,
      proposed_budget REAL DEFAULT 0,
      balance_liability REAL DEFAULT 0,
      actual_payment REAL DEFAULT 0,
      last_modified_by TEXT,
      last_modified_at TEXT,
      year INTEGER,
      month_index INTEGER,
      month_name TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY (component_id) REFERENCES components(id)
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS actual_payment_history (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      sub_component_id INTEGER NOT NULL,
      previous_amount REAL DEFAULT 0,
      adjustment_amount REAL DEFAULT 0,
      new_amount REAL DEFAULT 0,
      action_type TEXT,
      changed_by TEXT,
      payment_date TEXT,
      changed_at TEXT DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY (sub_component_id) REFERENCES sub_components(id)
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS revision_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      component_id INTEGER,
      component_item_id TEXT,
      component_name TEXT,
      reason TEXT,
      revised_by TEXT,
      revised_at TEXT DEFAULT CURRENT_TIMESTAMP,
      year INTEGER,
      month_index INTEGER,
      month_name TEXT
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS budget_change_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      sub_component_id INTEGER,
      component_id INTEGER,
      component_item_id TEXT,
      component_name TEXT,
      item_id TEXT,
      sub_component_name TEXT,
      field_name TEXT,
      previous_amount REAL DEFAULT 0,
      new_amount REAL DEFAULT 0,
      change_amount REAL DEFAULT 0,
      exceeded_amount REAL DEFAULT 0,
      changed_by TEXT,
      changed_at TEXT DEFAULT CURRENT_TIMESTAMP,
      year INTEGER,
      month_index INTEGER,
      month_name TEXT
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS revision_change_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      revision_log_id INTEGER,
      component_id INTEGER,
      component_item_id TEXT,
      component_name TEXT,
      item_id TEXT,
      sub_component_name TEXT,
      change_type TEXT,
      field_label TEXT,
      previous_value TEXT,
      new_value TEXT,
      changed_at TEXT DEFAULT CURRENT_TIMESTAMP,
      year INTEGER,
      month_index INTEGER,
      month_name TEXT,
      FOREIGN KEY (revision_log_id) REFERENCES revision_logs(id)
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS budget_submissions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      year INTEGER NOT NULL,
      month_index INTEGER NOT NULL,
      month_name TEXT NOT NULL,
      submitted_by TEXT NOT NULL,
      submitted_at TEXT DEFAULT CURRENT_TIMESTAMP,
      UNIQUE(year, month_index)
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS ceo_submission_views (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      submission_id INTEGER NOT NULL,
      ceo_username TEXT NOT NULL,
      viewed_at TEXT DEFAULT CURRENT_TIMESTAMP,
      UNIQUE(submission_id, ceo_username),
      FOREIGN KEY (submission_id) REFERENCES budget_submissions(id)
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS budget_submission_components (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      submission_id INTEGER NOT NULL,
      component_id INTEGER NOT NULL,
      component_name TEXT NOT NULL,
      UNIQUE(submission_id, component_id),
      FOREIGN KEY (submission_id) REFERENCES budget_submissions(id)
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS budget_submission_rows (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      submission_id INTEGER NOT NULL,
      sub_component_id INTEGER NOT NULL,
      component_id INTEGER NOT NULL,
      review_status TEXT NOT NULL DEFAULT 'Pending',
      reviewed_at TEXT,
      UNIQUE(submission_id, sub_component_id),
      FOREIGN KEY (submission_id) REFERENCES budget_submissions(id)
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT NOT NULL UNIQUE,
      password_hash TEXT NOT NULL,
      role TEXT NOT NULL,
      full_name TEXT,
      employee_id TEXT,
      email TEXT,
      designation TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  db.run(
    `INSERT OR IGNORE INTO users (username, password_hash, role, full_name, employee_id, designation) VALUES (?, ?, ?, ?, ?, ?);`,
    ["ceo", hashPassword("CEO@1234"), "Admin(CEO)", "CEO", "CEO-001", "Chief Executive Officer"]
  );

  db.run(
    `INSERT OR IGNORE INTO users (username, password_hash, role, full_name, employee_id, designation) VALUES (?, ?, ?, ?, ?, ?);`,
    ["financeadmin", hashPassword("Finance@1234"), "Finance Admin", "Finance Admin", "FIN-001", "Finance Administrator"]
  );

  db.run(
    `INSERT OR IGNORE INTO users (username, password_hash, role, full_name, employee_id, designation) VALUES (?, ?, ?, ?, ?, ?);`,
    ["accountteam", hashPassword("Account@1234"), "Account Team", "Account Team", "ACC-001", "Account Team"]
  );

  ensureColumn("components", "year", "INTEGER");
  ensureColumn("components", "month_index", "INTEGER");
  ensureColumn("components", "month_name", "TEXT");
  ensureColumn("sub_components", "year", "INTEGER");
  ensureColumn("sub_components", "month_index", "INTEGER");
  ensureColumn("sub_components", "month_name", "TEXT");
  ensureColumn("sub_components", "amount_value", "REAL DEFAULT 0");
  ensureColumn("sub_components", "amount_cell", "TEXT");
  ensureColumn("sub_components", "existing_liability", "REAL DEFAULT 0");
  ensureColumn("sub_components", "estimated_liability", "REAL DEFAULT 0");
  ensureColumn("sub_components", "total_estimated_liability", "REAL DEFAULT 0");
  ensureColumn("sub_components", "proposed_budget", "REAL DEFAULT 0");
  ensureColumn("sub_components", "balance_liability", "REAL DEFAULT 0");
  ensureColumn("sub_components", "actual_payment", "REAL DEFAULT 0");
  ensureColumn("actual_payment_history", "payment_date", "TEXT");
  ensureColumn("sub_components", "approval_status", "TEXT DEFAULT 'Pending'");
ensureColumn("sub_components", "ceo_remarks", "TEXT");
ensureColumn("sub_components", "approved_by", "TEXT");
ensureColumn("sub_components", "approved_at", "TEXT");
ensureColumn("sub_components", "rejected_by", "TEXT");
ensureColumn("sub_components", "rejected_at", "TEXT");
ensureColumn("sub_components", "last_modified_by", "TEXT");
ensureColumn("sub_components", "last_modified_at", "TEXT");
});

async function migrateComponentsTable() {
  try {
    const existing = await all(`SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'components'`);
    if (!existing || existing.length === 0) return;

    const createSql = String(existing[0].sql || "").toUpperCase();
    if (!createSql.includes("UNIQUE") || !createSql.includes("NAME")) return;

    console.log("Migrating components table to support month-specific component rows...");
    await run("PRAGMA foreign_keys = OFF");
    await run("ALTER TABLE components RENAME TO components_old");
    await run(`
      CREATE TABLE IF NOT EXISTS components (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        item_id TEXT,
        name TEXT NOT NULL,
        year INTEGER,
        month_index INTEGER,
        month_name TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
      )
    `);
    await run("CREATE UNIQUE INDEX IF NOT EXISTS components_monthly_unique ON components(name, year, month_index)");
    await run(`
      INSERT INTO components (id, item_id, name, year, month_index, month_name, created_at)
      SELECT id, item_id, name, year, month_index, month_name, created_at FROM components_old
    `);
    await run("DROP TABLE components_old");
    await run("PRAGMA foreign_keys = ON");
  } catch (error) {
    console.error("Component table migration failed:", error.message);
  }
}

migrateComponentsTable().catch((err) => console.error(err));

function ensureColumn(tableName, columnName, definition) {
  return new Promise((resolve) => {
    db.all(`PRAGMA table_info(${tableName})`, (error, columns) => {
      if (error) {
        console.error(`Unable to inspect ${tableName}:`, error.message);
        resolve();
        return;
      }

      if (!columns.some((column) => column.name === columnName)) {
        db.run(`ALTER TABLE ${tableName} ADD COLUMN ${columnName} ${definition}`, (alterError) => {
          if (alterError) {
            console.error(`Unable to add column ${columnName} to ${tableName}:`, alterError.message);
          }
          resolve();
        });
      } else {
        resolve();
      }
    });
  });
}

function getMonthName(monthIndex) {
  const monthNames = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
  ];

  return monthNames[monthIndex] || "";
}

const financialQuarterMonths = {
  Q1: [3, 4, 5],
  Q2: [6, 7, 8],
  Q3: [9, 10, 11],
  Q4: [0, 1, 2]
};

function parseCsvQuery(value) {
  if (!value) return [];
  return String(value)
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function parseListQuery(value) {
  if (!value) return [];
  return String(value)
    .split("||")
    .map((item) => item.trim())
    .filter(Boolean);
}

function getFinancialYearStart(date = new Date()) {
  const monthIndex = date.getMonth();
  const year = date.getFullYear();
  return monthIndex >= 3 ? year : year - 1;
}

function buildDashboardFilters(query = {}) {
  const today = new Date();
  const hasFilters = ["financialYear", "quarters", "months", "components"].some((key) => query[key] !== undefined);
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth();
  const financialYear = Number(query.financialYear) || getFinancialYearStart(today);
  const quarters = parseCsvQuery(query.quarters).filter((quarter) => financialQuarterMonths[quarter]);
  const componentIds = parseListQuery(query.components)
    .map((componentId) => Number(componentId))
    .filter((componentId) => Number.isInteger(componentId) && componentId > 0);
  const requestedMonths = parseCsvQuery(query.months)
    .map((month) => Number(month))
    .filter((month) => Number.isInteger(month) && month >= 0 && month <= 11);

  let months = [];
  if (hasFilters) {
    const quarterMonths = quarters.length
      ? quarters.flatMap((quarter) => financialQuarterMonths[quarter])
      : [3, 4, 5, 6, 7, 8, 9, 10, 11, 0, 1, 2];

    months = requestedMonths.length
      ? requestedMonths.filter((month) => quarterMonths.includes(month))
      : quarterMonths;
  } else {
    months = [currentMonth];
  }

  const periods = months.map((monthIndex) => ({
    year: hasFilters && monthIndex <= 2 ? financialYear + 1 : hasFilters ? financialYear : currentYear,
    monthIndex,
    monthName: getMonthName(monthIndex)
  }));

  return {
    hasFilters,
    financialYear,
    quarters,
    months,
    componentIds,
    periods
  };
}

function appendDashboardWhere(filters, alias = "sc") {
  const params = [];
  const clauses = [];
  const yearExpr = `COALESCE(${alias}.year, CAST(strftime('%Y', ${alias}.created_at) AS INTEGER))`;
  const monthExpr = `COALESCE(${alias}.month_index, CAST(strftime('%m', ${alias}.created_at) AS INTEGER) - 1)`;

  if (filters.periods.length > 0) {
    clauses.push(`(${filters.periods.map(() => `(${yearExpr} = ? AND ${monthExpr} = ?)`).join(" OR ")})`);
    filters.periods.forEach((period) => {
      params.push(period.year, period.monthIndex);
    });
  } else {
    clauses.push("1 = 0");
  }

  if (filters.componentIds.length > 0) {
    clauses.push(`c.id IN (${filters.componentIds.map(() => "?").join(", ")})`);
    params.push(...filters.componentIds);
  }

  return {
    where: clauses.length ? `WHERE ${clauses.join(" AND ")}` : "",
    params
  };
}

function normalizePeriod(body) {
  const today = new Date();
  const year = Number(body.year) || today.getFullYear();
  const monthIndex =
    body.monthIndex === 0 || body.monthIndex
      ? Number(body.monthIndex)
      : today.getMonth();
  const monthName = String(body.monthName || getMonthName(monthIndex)).trim();

  return { year, monthIndex, monthName };
}

function isYellowFill(cell) {
  const fill = cell.fill;

  if (!fill || fill.type !== "pattern" || !fill.fgColor) {
    return false;
  }

  const argb = String(fill.fgColor.argb || "").toUpperCase();

  return argb === "FFFFFF00" || argb === "FFFF00";
}

function cleanText(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "object" && value.text) {
    return String(value.text).trim();
  }

  if (typeof value === "object" && Array.isArray(value.richText)) {
    return value.richText.map((part) => part.text || "").join("").trim();
  }

  return String(value).trim();
}

function readNumber(value) {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  if (typeof value === "object" && value.result !== undefined) {
    return Number(value.result) || 0;
  }

  return Number(value) || 0;
}

function isApprovedStatus(value) {
  return String(value || "").trim() === "Approved";
}

function getRowsBudgetApproved(rows) {
  return rows.length > 0 && rows.every((row) => isApprovedStatus(row.approvalStatus || row.approval_status));
}

function maskRowsForPendingBudget(rows, budgetApproved) {
  if (budgetApproved) {
    return rows;
  }

  return rows.map((row) => ({
    ...row,
    actualPayment: 0
  }));
}

async function getCurrentMonthBudgetStatus(year, monthIndex) {
  const rows = await all(
    `SELECT COUNT(1) AS totalRows,
            SUM(CASE WHEN approval_status = 'Approved' THEN 1 ELSE 0 END) AS approvedRows,
            COALESCE(SUM(COALESCE(proposed_budget, amount_value, 0)), 0) AS approvedBudget,
            COALESCE(SUM(COALESCE(actual_payment, 0)), 0) AS actualPayment
     FROM sub_components
     WHERE COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ?
       AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?`,
    [year, monthIndex]
  );

  const status = rows[0] || {};
  const totalRows = Number(status.totalRows) || 0;
  const approvedRows = Number(status.approvedRows) || 0;

  return {
    budgetApproved: totalRows > 0 && totalRows === approvedRows,
    approvedBudget: Number(status.approvedBudget) || 0,
    actualPayment: Number(status.actualPayment) || 0
  };
}

function validateActualPaymentLimit(actualPayment, approvedBudget) {
  const payment = Number(actualPayment) || 0;
  const budget = Number(approvedBudget) || 0;

  if (payment > budget) {
    throw new Error("Actual payment cannot exceed the approved budget.");
  }
}

async function parseExcelComponents(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];

  if (!worksheet) {
    throw new Error("The uploaded Excel file does not contain any worksheet.");
  }

  const components = [];
  let currentComponent = null;

  worksheet.eachRow({ includeEmpty: false }, (row) => {
    const itemId = cleanText(row.getCell(1).value);
    const name = cleanText(row.getCell(2).value);
    const amountCell = row.getCell(6);
    const amountValue = readNumber(amountCell.value);

    if (!name) {
      return;
    }

    if (isYellowFill(row.getCell(2))) {
      currentComponent = {
        itemId,
        name,
        subComponents: []
      };
      components.push(currentComponent);
      return;
    }

    if (currentComponent) {
      currentComponent.subComponents.push({
        itemId: itemId || `${currentComponent.itemId}.${currentComponent.subComponents.length + 1}`,
        name,
        amountValue,
        amountCell: amountCell.address
      });
    }
  });

  return components;
}

const subComponentSelect = `
  SELECT
    sc.id,
    sc.item_id AS itemId,
    sc.name,
    COALESCE(sc.amount_value, 0) AS amountValue,
    sc.amount_cell AS amountCell,
    COALESCE(sc.existing_liability, 0) AS existingLiability,
    COALESCE(sc.estimated_liability, 0) AS estimatedLiability,
    COALESCE(sc.total_estimated_liability, 0) AS totalEstimatedLiability,
    COALESCE(sc.proposed_budget, sc.amount_value, 0) AS proposedBudget,
    COALESCE(sc.balance_liability, 0) AS balanceLiability,
    COALESCE(sc.actual_payment, 0) AS actualPayment,
    sc.last_modified_by AS lastModifiedBy,
    sc.last_modified_at AS lastModifiedAt,
    sc.approval_status AS approvalStatus,
    sc.ceo_remarks AS ceoRemarks,
    sc.approved_by AS approvedBy,
    sc.approved_at AS approvedAt,
    sc.rejected_by AS rejectedBy,
    sc.rejected_at AS rejectedAt,
    sc.component_id AS componentId,
    c.item_id AS componentItemId,
    c.name AS componentName,
    COALESCE(sc.year, CAST(strftime('%Y', sc.created_at) AS INTEGER)) AS year,
    COALESCE(sc.month_index, CAST(strftime('%m', sc.created_at) AS INTEGER) - 1) AS monthIndex,
    COALESCE(
      sc.month_name,
      CASE CAST(strftime('%m', sc.created_at) AS INTEGER) - 1
        WHEN 0 THEN 'January'
        WHEN 1 THEN 'February'
        WHEN 2 THEN 'March'
        WHEN 3 THEN 'April'
        WHEN 4 THEN 'May'
        WHEN 5 THEN 'June'
        WHEN 6 THEN 'July'
        WHEN 7 THEN 'August'
        WHEN 8 THEN 'September'
        WHEN 9 THEN 'October'
        WHEN 10 THEN 'November'
        WHEN 11 THEN 'December'
      END
    ) AS monthName,
    sc.created_at AS createdAt
  FROM sub_components sc
  JOIN components c ON c.id = sc.component_id
`;

function all(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.all(sql, params, (error, rows) => {
      if (error) reject(error);
      else resolve(rows);
    });
  });
}

function run(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.run(sql, params, function onRun(error) {
      if (error) reject(error);
      else resolve(this);
    });
  });
}

function isHiddenComponentName(name) {
  const normalized = String(name || "").trim().toUpperCase();
  return normalized === "TOTAL PAYMENT";
}

async function ensureLogTables() {
  await run(`
    CREATE TABLE IF NOT EXISTS revision_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      component_id INTEGER,
      component_item_id TEXT,
      component_name TEXT,
      reason TEXT,
      revised_by TEXT,
      revised_at TEXT DEFAULT CURRENT_TIMESTAMP,
      year INTEGER,
      month_index INTEGER,
      month_name TEXT
    )
  `);
  await run(`
    CREATE TABLE IF NOT EXISTS budget_change_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      sub_component_id INTEGER,
      component_id INTEGER,
      component_item_id TEXT,
      component_name TEXT,
      item_id TEXT,
      sub_component_name TEXT,
      field_name TEXT,
      previous_amount REAL DEFAULT 0,
      new_amount REAL DEFAULT 0,
      change_amount REAL DEFAULT 0,
      exceeded_amount REAL DEFAULT 0,
      changed_by TEXT,
      changed_at TEXT DEFAULT CURRENT_TIMESTAMP,
      year INTEGER,
      month_index INTEGER,
      month_name TEXT
    )
  `);
  await run(`
    CREATE TABLE IF NOT EXISTS revision_change_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      revision_log_id INTEGER,
      component_id INTEGER,
      component_item_id TEXT,
      component_name TEXT,
      item_id TEXT,
      sub_component_name TEXT,
      change_type TEXT,
      field_label TEXT,
      previous_value TEXT,
      new_value TEXT,
      changed_at TEXT DEFAULT CURRENT_TIMESTAMP,
      year INTEGER,
      month_index INTEGER,
      month_name TEXT,
      FOREIGN KEY (revision_log_id) REFERENCES revision_logs(id)
    )
  `);
}

async function logBudgetChange({
  subComponentId,
  componentId,
  componentItemId,
  componentName,
  itemId,
  subComponentName,
  fieldName,
  previousAmount,
  newAmount,
  exceededAmount = 0,
  changedBy,
  year,
  monthIndex,
  monthName
}) {
  await ensureLogTables();
  const previous = Number(previousAmount) || 0;
  const next = Number(newAmount) || 0;
  if (Math.abs(previous - next) <= 0.004) return;

  await run(
    `INSERT INTO budget_change_logs
      (sub_component_id, component_id, component_item_id, component_name, item_id, sub_component_name,
       field_name, previous_amount, new_amount, change_amount, exceeded_amount, changed_by, year, month_index, month_name)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [
      subComponentId || null,
      componentId || null,
      componentItemId || "",
      componentName || "",
      itemId || "",
      subComponentName || "",
      fieldName || "",
      previous,
      next,
      next - previous,
      Number(exceededAmount) || 0,
      changedBy || "Unknown",
      year,
      monthIndex,
      monthName
    ]
  );
}

function groupRowsByComponent(rows, components = []) {
  const groups = [];
  const byId = new Map();

  const addGroup = (component) => {
    const componentId = String(component.componentId ?? component.id ?? "");

    if (!componentId || byId.has(componentId)) {
      return;
    }

    const group = {
      componentId,
      componentItemId: component.componentItemId || component.itemId || "",
      componentName: component.componentName || component.name || "",
      totalAmount: 0,
      rows: []
    };

    byId.set(componentId, group);
    groups.push(group);
  };

  if (Array.isArray(components)) {
    components
      .filter((component) => !isHiddenComponentName(component.componentName || component.name))
      .forEach(addGroup);
  }

  for (const row of rows) {
    const componentId = String(row.componentId ?? row.component_id ?? "");

    if (!byId.has(componentId)) {
      addGroup({
        componentId,
        componentItemId: row.componentItemId || row.component_item_id || "",
        componentName: row.componentName || row.component_name || ""
      });
    }

    const group = byId.get(componentId);
    group.rows.push(row);
    group.totalAmount += Number(row.amountValue) || 0;
  }

  return groups;
}

function formatAmount(value) {
  const amount = Number(value) || 0;
  return amount.toLocaleString("en-IN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

async function buildBudgetWorkbook({ title, rows, components = [] }) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("PAYMENT");
  const budgetApproved = getRowsBudgetApproved(rows);
  const headers = [
    "SL. NO",
    "PARTICULARS",
    "EXISTING LIABILITY",
    "ESTIMATED LIABILITY",
    "TOTAL ESTIMATED LIABILITY",
    budgetApproved ? "APPROVED BUDGET" : "PROPOSED BUDGET",
    "BALANCE LIABILITY"
  ];
  if (budgetApproved) {
    headers.push("ACTUAL PAYMENT", "NET BALANCE");
  }
  const columnCount = headers.length;

  worksheet.mergeCells(1, 1, 1, columnCount);
  worksheet.getCell("A1").value = title;
  worksheet.getCell("A1").font = { bold: true, size: 14 };
  worksheet.getCell("A1").alignment = { horizontal: "center" };

  worksheet.getRow(2).values = headers;
  worksheet.getRow(2).eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });

  let currentRow = 3;
  const groups = groupRowsByComponent(rows, components);

  for (const group of groups) {
    const componentRow = worksheet.getRow(currentRow);
    componentRow.getCell(1).value = group.componentItemId || "";
    componentRow.getCell(2).value = group.componentName;
    const totals = group.rows.reduce((sum, row) => {
      const existingLiability = Number(row.existingLiability) || 0;
      const estimatedLiability = Number(row.estimatedLiability) || 0;
      const proposedBudget = Number(row.proposedBudget || row.amountValue || 0);
      const actualPayment = Number(row.actualPayment) || 0;
      const totalEstimatedLiability = existingLiability + estimatedLiability;
      const balanceLiability = totalEstimatedLiability - proposedBudget;

      return {
        existingLiability: sum.existingLiability + existingLiability,
        estimatedLiability: sum.estimatedLiability + estimatedLiability,
        totalEstimatedLiability: sum.totalEstimatedLiability + totalEstimatedLiability,
        proposedBudget: sum.proposedBudget + proposedBudget,
        balanceLiability: sum.balanceLiability + balanceLiability,
        actualPayment: sum.actualPayment + actualPayment
      };
    }, { existingLiability: 0, estimatedLiability: 0, totalEstimatedLiability: 0, proposedBudget: 0, balanceLiability: 0, actualPayment: 0 });

    componentRow.getCell(3).value = totals.existingLiability;
    componentRow.getCell(4).value = totals.estimatedLiability;
    componentRow.getCell(5).value = totals.totalEstimatedLiability;
    componentRow.getCell(6).value = totals.proposedBudget;
    componentRow.getCell(7).value = totals.balanceLiability;
    if (budgetApproved) {
      componentRow.getCell(8).value = totals.actualPayment;
      componentRow.getCell(9).value = totals.balanceLiability - totals.actualPayment;
    }

    for (let colNumber = 1; colNumber <= columnCount; colNumber += 1) {
      const cell = componentRow.getCell(colNumber);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
      cell.font = { bold: true };
    }

    currentRow += 1;

    group.rows.forEach((row, index) => {
      const subRow = worksheet.getRow(currentRow);
      const existingLiability = Number(row.existingLiability) || 0;
      const estimatedLiability = Number(row.estimatedLiability) || 0;
      const proposedBudget = Number(row.proposedBudget || row.amountValue || 0);
      const actualPayment = Number(row.actualPayment) || 0;
      const totalEstimatedLiability = existingLiability + estimatedLiability;
      const balanceLiability = totalEstimatedLiability - proposedBudget;
      subRow.getCell(1).value = row.itemId || `${group.componentItemId}.${index + 1}`;
      subRow.getCell(2).value = row.name;
      subRow.getCell(3).value = existingLiability;
      subRow.getCell(4).value = estimatedLiability;
      subRow.getCell(5).value = totalEstimatedLiability;
      subRow.getCell(6).value = proposedBudget;
      subRow.getCell(7).value = balanceLiability;
      if (budgetApproved) {
        subRow.getCell(8).value = actualPayment;
        subRow.getCell(9).value = balanceLiability - actualPayment;
      }
      currentRow += 1;
    });
  }

  const worksheetColumns = [
    { width: 12 },
    { width: 42 },
    { width: 18 },
    { width: 18 },
    { width: 22 },
    { width: 18 },
    { width: 18 }
  ];
  if (budgetApproved) {
    worksheetColumns.push({ width: 18 }, { width: 18 });
  }
  worksheet.columns = worksheetColumns;

  worksheet.eachRow((row) => {
    for (let colNumber = 1; colNumber <= columnCount; colNumber += 1) {
      const cell = row.getCell(colNumber);
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
      cell.alignment = { vertical: "middle", wrapText: true };
    }
  });

  for (let colNumber = 3; colNumber <= columnCount; colNumber += 1) {
    worksheet.getColumn(colNumber).numFmt = "#,##0.00";
  }
  worksheet.views = [{ state: "frozen", ySplit: 2 }];

  return workbook;
}

async function sendBudgetWorkbook(res, { title, fileName, rows, components = [] }) {
  const workbook = await buildBudgetWorkbook({ title, rows, components });
  const buffer = await workbook.xlsx.writeBuffer();

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
  res.send(Buffer.from(buffer));
}

app.post("/api/login", async (req, res) => {
  const role = String(req.body?.role || "").trim();
  const username = String(req.body?.username || "").trim();
  const password = String(req.body?.password || "");

  if (!role || !username || !password) {
    return res.status(400).json({ error: "Role, username, and password are required." });
  }

  try {
    const users = await all(
      `
        SELECT id, username, role, full_name AS fullName, employee_id AS employeeId, email, designation
        FROM users
        WHERE username = ?
          AND role = ?
          AND password_hash = ?
        LIMIT 1
      `,
      [username.toLowerCase(), role, hashPassword(password)]
    );

    if (users.length === 0) {
      return res.status(401).json({ error: "Invalid username, password, or role." });
    }

    res.json({ success: true, message: "Login successful.", user: users[0] });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/users", async (req, res) => {
  const username = String(req.body?.username || "").trim().toLowerCase();
  const password = String(req.body?.password || "");
  const role = String(req.body?.role || "").trim();
  const fullName = String(req.body?.fullName || "").trim();
  const employeeId = String(req.body?.employeeId || "").trim();
  const email = String(req.body?.email || "").trim();
  const designation = String(req.body?.designation || "").trim();

  if (!username || !password || !role) {
    return res.status(400).json({ error: "Username, password, and role are required." });
  }

  try {
    const existing = await all("SELECT id FROM users WHERE username = ?", [username]);
    if (existing.length > 0) {
      return res.status(409).json({ error: "That username already exists." });
    }

    const result = await run(
      `
        INSERT INTO users (username, password_hash, role, full_name, employee_id, email, designation)
        VALUES (?, ?, ?, ?, ?, ?, ?)
      `,
      [username, hashPassword(password), role, fullName, employeeId, email, designation]
    );

    res.status(201).json({
      success: true,
      message: "User created successfully.",
      user: {
        id: result.lastID,
        username,
        role,
        fullName,
        employeeId,
        email,
        designation
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/health", (req, res) => {
  res.json({ status: "ok" });
});

app.get("/api/components", async (req, res) => {
  try {
    const year = req.query.year !== undefined ? Number(req.query.year) : null;
    const monthIndex = req.query.monthIndex !== undefined ? Number(req.query.monthIndex) : null;
    let sql = "SELECT id, item_id AS itemId, name, year, month_index AS monthIndex, month_name AS monthName FROM components";
    const params = [];

    if (!Number.isNaN(year) && !Number.isNaN(monthIndex) && year !== null && monthIndex !== null) {
      sql += " WHERE COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ? AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?";
      params.push(year, monthIndex);
    }

    sql += " ORDER BY id";
    const rows = await all(sql, params);
    res.json(rows);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/components", async (req, res) => {
  const period = normalizePeriod(req.body);
  const components = Array.isArray(req.body.components)
    ? req.body.components
    : [{ itemId: req.body.itemId, name: req.body.name || req.body.component_name }];

  const cleanComponents = components
    .map((component) => ({
      itemId: String(component.itemId || "").trim(),
      name: String(component.name || "").trim()
    }))
    .filter((component) => component.name);

  if (cleanComponents.length === 0) {
    return res.status(400).json({ error: "At least one component name is required." });
  }

  try {
    await run("BEGIN TRANSACTION");

    for (const component of cleanComponents) {
      await run(
        `
          INSERT INTO components (item_id, name, year, month_index, month_name)
          VALUES (?, ?, ?, ?, ?)
          ON CONFLICT(name, year, month_index) DO UPDATE SET
            item_id = excluded.item_id,
            month_name = excluded.month_name
        `,
        [component.itemId, component.name, period.year, period.monthIndex, period.monthName]
      );
    }

    await run("COMMIT");

    const rows = await all(
      "SELECT id, item_id AS itemId, name, year, month_index AS monthIndex, month_name AS monthName FROM components WHERE COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ? AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ? ORDER BY id",
      [period.year, period.monthIndex]
    );
    res.status(201).json(rows);
  } catch (error) {
    await run("ROLLBACK").catch(() => {});
    const statusCode = error.message === "Actual payment cannot exceed the approved budget." ? 400 : 500;
    res.status(statusCode).json({ error: error.message });
  }
});

app.get("/api/sub-components", async (req, res) => {
  try {
    const year = req.query.year !== undefined ? Number(req.query.year) : null;
    const monthIndex = req.query.monthIndex !== undefined ? Number(req.query.monthIndex) : null;
    let sql = `${subComponentSelect}`;
    const params = [];

    if (!Number.isNaN(year) && !Number.isNaN(monthIndex) && year !== null && monthIndex !== null) {
      sql += ` WHERE COALESCE(sc.year, CAST(strftime('%Y', sc.created_at) AS INTEGER)) = ? AND COALESCE(sc.month_index, CAST(strftime('%m', sc.created_at) AS INTEGER) - 1) = ?`;
      params.push(year, monthIndex);
    }

    sql += " ORDER BY sc.id";
    const rows = await all(sql, params);
    res.json(rows);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/sub-components", async (req, res) => {
  const period = normalizePeriod(req.body);
  const componentId = Number(req.body.componentId);
  const subComponents = Array.isArray(req.body.subComponents)
    ? req.body.subComponents
    : [
        {
          itemId: req.body.itemId,
          name: req.body.name || req.body.sub_component_name,
          amountValue: req.body.amountValue,
          amountCell: req.body.amountCell
        }
      ];

  const cleanSubComponents = subComponents
    .map((subComponent) => ({
      itemId: String(subComponent.itemId || "").trim(),
      name: String(subComponent.name || "").trim(),
      amountValue: Number(subComponent.amountValue) || 0,
      amountCell: String(subComponent.amountCell || "").trim()
    }))
    .filter((subComponent) => subComponent.name);

  if (!componentId) {
    return res.status(400).json({ error: "Please choose a component first." });
  }

  if (cleanSubComponents.length === 0) {
    return res.status(400).json({ error: "At least one sub-component name is required." });
  }

  try {
      const component = await all(
        "SELECT id FROM components WHERE id = ? AND COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ? AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?",
        [componentId, period.year, period.monthIndex]
      );

      if (component.length === 0) {
        return res.status(404).json({ error: "Selected component was not found for the chosen period." });
      }

      await run("BEGIN TRANSACTION");

      for (const subComponent of cleanSubComponents) {
        await run(
          `
            INSERT INTO sub_components
              (item_id, component_id, name, amount_value, amount_cell, year, month_index, month_name)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            subComponent.itemId,
            componentId,
            subComponent.name,
            subComponent.amountValue,
            subComponent.amountCell,
            period.year,
            period.monthIndex,
            period.monthName
          ]
        );
      }

      await run("COMMIT");

      const rows = await all(
        `${subComponentSelect} WHERE COALESCE(sc.year, CAST(strftime('%Y', sc.created_at) AS INTEGER)) = ? AND COALESCE(sc.month_index, CAST(strftime('%m', sc.created_at) AS INTEGER) - 1) = ? ORDER BY sc.id`,
        [period.year, period.monthIndex]
      );
      res.status(201).json(rows);
  } catch (error) {
    await run("ROLLBACK").catch(() => {});
    const statusCode = error.message === "Actual payment cannot exceed the approved budget." ? 400 : 500;
    res.status(statusCode).json({ error: error.message });
  }
});

app.delete("/api/components/:id", async (req, res) => {
  const componentId = Number(req.params.id);

  if (!componentId) {
    return res.status(400).json({ error: "Invalid component id." });
  }

  try {
    await run("BEGIN TRANSACTION");
    await run("DELETE FROM sub_components WHERE component_id = ?", [componentId]);
    await run("DELETE FROM components WHERE id = ?", [componentId]);
    await run("COMMIT");

    const rows = await all(
      "SELECT id, item_id AS itemId, name, year, month_index AS monthIndex, month_name AS monthName FROM components ORDER BY id"
    );

    res.json({
      success: true,
      message: "Component removed successfully.",
      components: rows
    });
  } catch (error) {
    await run("ROLLBACK").catch(() => {});
    res.status(500).json({ error: error.message });
  }
});

app.delete("/api/sub-components/:id", async (req, res) => {
  const subComponentId = Number(req.params.id);

  if (!subComponentId) {
    return res.status(400).json({ error: "Invalid sub-component id." });
  }

  try {
    await run("BEGIN TRANSACTION");
    await run("DELETE FROM sub_components WHERE id = ?", [subComponentId]);
    await run("COMMIT");

    const rows = await all(`${subComponentSelect} ORDER BY sc.id`);

    res.json({
      success: true,
      message: "Sub-component removed successfully.",
      subComponents: rows
    });
  } catch (error) {
    await run("ROLLBACK").catch(() => {});
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/import-excel", upload.single("excelFile"), async (req, res) => {
  const period = normalizePeriod(req.body);

  if (!req.file) {
    return res.status(400).json({ error: "Please upload an Excel file." });
  }

  try {
    const parsedComponents = await parseExcelComponents(req.file.buffer);

    if (parsedComponents.length === 0) {
      return res.status(400).json({
        error: "No yellow-highlighted components were found in the uploaded Excel file."
      });
    }

    let importedComponents = 0;
    let importedSubComponents = 0;

    await run("BEGIN TRANSACTION");

    for (const component of parsedComponents) {
      await run(
        `
          INSERT INTO components (item_id, name, year, month_index, month_name)
          VALUES (?, ?, ?, ?, ?)
          ON CONFLICT(name, year, month_index) DO UPDATE SET
            item_id = excluded.item_id,
            month_name = excluded.month_name
        `,
        [component.itemId, component.name, period.year, period.monthIndex, period.monthName]
      );

      importedComponents += 1;

      const savedComponent = await all(
        "SELECT id FROM components WHERE name = ? AND year = ? AND month_index = ?",
        [component.name, period.year, period.monthIndex]
      );

      if (savedComponent.length === 0) {
        continue;
      }

      const componentId = savedComponent[0].id;

      for (const subComponent of component.subComponents) {
        const existing = await all(
          `
            SELECT id
            FROM sub_components
            WHERE component_id = ?
              AND name = ?
              AND COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ?
              AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?
            LIMIT 1
          `,
          [componentId, subComponent.name, period.year, period.monthIndex]
        );

        if (existing.length > 0) {
          await run(
            `
              UPDATE sub_components
              SET item_id = ?,
                  amount_value = ?,
                  amount_cell = ?
              WHERE id = ?
            `,
            [
              subComponent.itemId,
              Number(subComponent.amountValue) || 0,
              subComponent.amountCell || "",
              existing[0].id
            ]
          );
          continue;
        }

        await run(
          `
            INSERT INTO sub_components
              (item_id, component_id, name, amount_value, amount_cell, year, month_index, month_name)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            subComponent.itemId,
            componentId,
            subComponent.name,
            Number(subComponent.amountValue) || 0,
            subComponent.amountCell || "",
            period.year,
            period.monthIndex,
            period.monthName
          ]
        );

        importedSubComponents += 1;
      }
    }

    await run("COMMIT");

    const rows = await all(
      "SELECT id, item_id AS itemId, name, year, month_index AS monthIndex, month_name AS monthName FROM components WHERE COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ? AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ? ORDER BY id",
      [period.year, period.monthIndex]
    );

    res.status(201).json({
      message: "Excel imported successfully.",
      importedComponents,
      importedSubComponents,
      components: rows
    });
  } catch (error) {
    await run("ROLLBACK").catch(() => {});
    res.status(500).json({ error: error.message });
  }
});

app.patch("/api/view-data/current-month", async (req, res) => {
  await ensureColumn("sub_components", "actual_payment", "REAL DEFAULT 0");
  await ensureColumn("sub_components", "last_modified_by", "TEXT");
  await ensureColumn("sub_components", "last_modified_at", "TEXT");

  const today = new Date();
  const year = today.getFullYear();
  const monthIndex = today.getMonth();
  const monthName = getMonthName(monthIndex);
  const role = String(req.body?.role || "").trim();
  const modifiedBy = String(req.body?.modifiedBy || role || "Unknown").trim();
  const submitForApproval = Boolean(req.body?.submitForApproval) && role === "Finance Admin";
  const submittedComponents = Array.isArray(req.body?.submittedComponents)
    ? req.body.submittedComponents
        .map((component) => ({
          id: Number(component?.id),
          name: String(component?.name || "").trim()
        }))
        .filter((component) => Number.isInteger(component.id) && component.id > 0 && component.name)
    : [];
  const submittedRows = Array.isArray(req.body?.submittedRows)
    ? req.body.submittedRows
        .map((row) => ({ id: Number(row?.id), componentId: Number(row?.componentId) }))
        .filter((row) => Number.isInteger(row.id) && row.id > 0 && Number.isInteger(row.componentId) && row.componentId > 0)
    : [];
  const updates = Array.isArray(req.body?.updates) ? req.body.updates : [];
  const deleteIds = Array.isArray(req.body?.deleteIds)
    ? req.body.deleteIds.map((value) => Number(value)).filter((id) => Number.isInteger(id) && id > 0)
    : [];

  if (updates.length === 0 && deleteIds.length === 0) {
    return res.status(400).json({ error: "No updates or deletes were provided." });
  }

  try {
    await run("BEGIN TRANSACTION");

    const budgetStatus = await getCurrentMonthBudgetStatus(year, monthIndex);

    let deletedCount = 0;
    if (deleteIds.length > 0) {
      const placeholders = deleteIds.map(() => "?").join(",");
      const deleteResult = await run(
        `DELETE FROM sub_components
         WHERE id IN (${placeholders})
           AND COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ?
           AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?`,
        [...deleteIds, year, monthIndex]
      );
      deletedCount = deleteResult.changes || 0;
    }

    for (const update of updates) {
      const id = Number(update.id);
      const componentId = Number(update.componentId);
      const name = String(update.name || "").trim();

      const existingRow = await all(
        `SELECT sc.item_id AS itemId,
                sc.name AS name,
                sc.component_id AS componentId,
                c.item_id AS componentItemId,
                c.name AS componentName,
                sc.existing_liability AS existingLiability,
                sc.estimated_liability AS estimatedLiability,
                sc.total_estimated_liability AS totalEstimatedLiability,
                sc.proposed_budget AS proposedBudget,
                sc.balance_liability AS balanceLiability,
                sc.actual_payment AS actualPayment,
                sc.approval_status AS approvalStatus
         FROM sub_components sc
         LEFT JOIN components c ON c.id = sc.component_id
         WHERE sc.id = ?`,
        [id]
      );
      const currentRow = existingRow[0] || {};
      const existingLiability = role === "Account Team"
        ? Number(currentRow.existingLiability) || 0
        : Number(update.existingLiability) || 0;
      const estimatedLiability = role === "Account Team"
        ? Number(currentRow.estimatedLiability) || 0
        : Number(update.estimatedLiability) || 0;
      const proposedBudget = role === "Account Team"
        ? Number(currentRow.proposedBudget) || 0
        : Number(update.proposedBudget) || 0;
      const currentActualPayment = Number(currentRow.actualPayment) || 0;
      const requestedActualPayment = Number(update.actualPayment) || 0;
      const actualPayment = budgetStatus.budgetApproved ? requestedActualPayment : currentActualPayment;
      const totalEstimatedLiability = existingLiability + estimatedLiability;
      const balanceLiability = totalEstimatedLiability - proposedBudget;

      if (!id) {
        if (!componentId || !name) {
          continue;
        }

        const siblingRows = await all(
          "SELECT COUNT(1) AS count FROM sub_components WHERE component_id = ? AND COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ? AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?",
          [componentId, year, monthIndex]
        );
        const siblingCount = Number(siblingRows[0]?.count || 0);
        const componentRows = await all("SELECT item_id AS itemId, name AS componentName FROM components WHERE id = ?", [componentId]);
        const componentItemId = componentRows[0]?.itemId || "";
        const itemId = String(update.itemId || `${componentItemId}.${siblingCount + 1}`).trim();
        if (budgetStatus.budgetApproved) {
          validateActualPaymentLimit(actualPayment, proposedBudget);
        }

        const insertResult = await run(
          `
            INSERT INTO sub_components
              (item_id, component_id, name, amount_value, existing_liability, estimated_liability, total_estimated_liability, proposed_budget, balance_liability, actual_payment, approval_status, last_modified_by, last_modified_at, year, month_index, month_name)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pending', ?, datetime('now'), ?, ?, ?)
          `,
          [
            itemId,
            componentId,
            name,
            proposedBudget,
            existingLiability,
            estimatedLiability,
            totalEstimatedLiability,
            proposedBudget,
            balanceLiability,
            budgetStatus.budgetApproved ? actualPayment : 0,
            modifiedBy,
            year,
            monthIndex,
            monthName
          ]
        );

        budgetStatus.actualPayment += budgetStatus.budgetApproved ? actualPayment : 0;
        budgetStatus.approvedBudget += proposedBudget;

        if (budgetStatus.budgetApproved && actualPayment !== 0) {
          await run(
            `INSERT INTO actual_payment_history
              (sub_component_id, previous_amount, adjustment_amount, new_amount, action_type, changed_by)
             VALUES (?, 0, ?, ?, 'initial', ?)`,
            [insertResult.lastID, actualPayment, actualPayment, modifiedBy]
          );
        }

        const componentName = componentRows[0]?.componentName || componentRows[0]?.name || "";
        const exceededAmount = Math.max(0, (budgetStatus.budgetApproved ? actualPayment : 0) - proposedBudget);
        await logBudgetChange({
          subComponentId: insertResult.lastID,
          componentId,
          componentItemId,
          componentName,
          itemId,
          subComponentName: name,
          fieldName: "Approved Budget",
          previousAmount: 0,
          newAmount: proposedBudget,
          exceededAmount,
          changedBy: modifiedBy,
          year,
          monthIndex,
          monthName
        });

        continue;
      }

      if (budgetStatus.budgetApproved) {
        validateActualPaymentLimit(actualPayment, proposedBudget);

        const monthActualAfterUpdate = budgetStatus.actualPayment - currentActualPayment + actualPayment;
        const monthBudgetAfterUpdate = budgetStatus.approvedBudget - (Number(currentRow.proposedBudget) || 0) + proposedBudget;
        validateActualPaymentLimit(monthActualAfterUpdate, monthBudgetAfterUpdate);
        budgetStatus.actualPayment = monthActualAfterUpdate;
        budgetStatus.approvedBudget = monthBudgetAfterUpdate;
      }

      const values = [
        name || update.name,
        existingLiability,
        estimatedLiability,
        totalEstimatedLiability,
        proposedBudget,
        balanceLiability,
        actualPayment,
        proposedBudget,
        modifiedBy,
        year,
        monthIndex,
        monthName,
        id,
        year,
        monthIndex
      ];

      await run(
        `
          UPDATE sub_components
          SET name = COALESCE(NULLIF(?, ''), name),
              existing_liability = ?,
              estimated_liability = ?,
              total_estimated_liability = ?,
              proposed_budget = ?,
              balance_liability = ?,
              actual_payment = ?,
              amount_value = ?,
              last_modified_by = ?,
              last_modified_at = datetime('now'),
              year = ?,
              month_index = ?,
              month_name = ?
          WHERE id = ?
            AND COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ?
            AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?
        `,
        values
      );

      const effectiveName = name || currentRow.name || update.name || "";
      const exceededAmount = Math.max(0, actualPayment - proposedBudget);
      const changeFields = [
        ["Existing Liability", currentRow.existingLiability, existingLiability],
        ["Estimated Liability", currentRow.estimatedLiability, estimatedLiability],
        [budgetStatus.budgetApproved ? "Approved Budget" : "Proposed Budget", currentRow.proposedBudget, proposedBudget],
        ["Actual Payment", currentActualPayment, actualPayment]
      ];

      for (const [fieldName, previousAmount, newAmount] of changeFields) {
        await logBudgetChange({
          subComponentId: id,
          componentId: currentRow.componentId || componentId,
          componentItemId: currentRow.componentItemId,
          componentName: currentRow.componentName,
          itemId: currentRow.itemId,
          subComponentName: effectiveName,
          fieldName,
          previousAmount,
          newAmount,
          exceededAmount,
          changedBy: modifiedBy,
          year,
          monthIndex,
          monthName
        });
      }

      const previousActualPayment = currentActualPayment;
      if (actualPayment !== previousActualPayment) {
        await run(
          `INSERT INTO actual_payment_history
            (sub_component_id, previous_amount, adjustment_amount, new_amount, action_type, changed_by)
           VALUES (?, ?, ?, ?, 'set', ?)`,
          [id, previousActualPayment, actualPayment - previousActualPayment, actualPayment, modifiedBy]
        );
      }
    }

    if (submitForApproval) {
      await run(
        `INSERT INTO budget_submissions (year, month_index, month_name, submitted_by, submitted_at)
         VALUES (?, ?, ?, ?, datetime('now'))
         ON CONFLICT(year, month_index) DO UPDATE SET
           month_name = excluded.month_name,
           submitted_by = excluded.submitted_by,
           submitted_at = excluded.submitted_at`,
        [year, monthIndex, monthName, modifiedBy]
      );
      const submissionRows = await all(
        "SELECT id FROM budget_submissions WHERE year = ? AND month_index = ?",
        [year, monthIndex]
      );
      const submissionId = submissionRows[0]?.id;
      if (submissionId) {
        await run("DELETE FROM budget_submission_components WHERE submission_id = ?", [submissionId]);
        await run("DELETE FROM budget_submission_rows WHERE submission_id = ?", [submissionId]);
        for (const component of submittedComponents) {
          await run(
            `INSERT OR IGNORE INTO budget_submission_components (submission_id, component_id, component_name)
             VALUES (?, ?, ?)`,
            [submissionId, component.id, component.name]
          );
        }
        for (const row of submittedRows) {
          await run(
            `INSERT OR IGNORE INTO budget_submission_rows (submission_id, sub_component_id, component_id)
             VALUES (?, ?, ?)`,
            [submissionId, row.id, row.componentId]
          );
        }
      }
      await run(
        `DELETE FROM ceo_submission_views
         WHERE submission_id = (
           SELECT id FROM budget_submissions WHERE year = ? AND month_index = ?
         )`,
        [year, monthIndex]
      );
    }

    await run("COMMIT");

    const updateCount = updates.filter((update) => Number(update.id)).length;
    let message = ``;
    if (updateCount > 0 && deletedCount > 0) {
      message = `Saved ${updateCount} update(s) and deleted ${deletedCount} row(s).`;
    } else if (updateCount > 0) {
      message = `Saved ${updateCount} update(s).`;
    } else if (deletedCount > 0) {
      message = `Deleted ${deletedCount} row(s).`;
    } else {
      message = "No changes saved.";
    }

    res.json({ message: submitForApproval ? "Budget submitted to the CEO for approval." : message });
  } catch (error) {
    await run("ROLLBACK").catch(() => {});
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/budget-submissions/pending", async (req, res) => {
  const ceoUsername = String(req.query?.username || "").trim();
  if (!ceoUsername) return res.status(400).json({ error: "CEO username is required." });

  const today = new Date();
  const year = today.getFullYear();
  const monthIndex = today.getMonth();

  try {
    const rows = await all(
      `SELECT bs.id, bs.month_name AS monthName, bs.submitted_by AS submittedBy, bs.submitted_at AS submittedAt
       FROM budget_submissions bs
       LEFT JOIN ceo_submission_views csv
         ON csv.submission_id = bs.id AND csv.ceo_username = ?
       WHERE bs.year = ? AND bs.month_index = ? AND csv.id IS NULL
       ORDER BY datetime(bs.submitted_at) DESC
       LIMIT 1`,
      [ceoUsername, year, monthIndex]
    );
    const submission = rows[0] || null;
    if (submission) {
      submission.components = await all(
        `SELECT component_id AS componentId, component_name AS componentName
         FROM budget_submission_components
         WHERE submission_id = ?
         ORDER BY component_name`,
        [submission.id]
      );
      submission.rows = await all(
        `SELECT sub_component_id AS subComponentId, component_id AS componentId, review_status AS reviewStatus
         FROM budget_submission_rows WHERE submission_id = ?`,
        [submission.id]
      );
    }
    res.json({ submission });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/budget-submissions/current/rows/:id/review", async (req, res) => {
  const subComponentId = Number(req.params.id);
  const action = String(req.body?.action || "").trim().toLowerCase();
  if (!Number.isInteger(subComponentId) || subComponentId <= 0 || !["approve", "reject"].includes(action)) {
    return res.status(400).json({ error: "A valid row review action is required." });
  }
  const today = new Date();
  try {
    const result = await run(
      `UPDATE budget_submission_rows
       SET review_status = ?, reviewed_at = datetime('now')
       WHERE sub_component_id = ?
         AND submission_id = (SELECT id FROM budget_submissions WHERE year = ? AND month_index = ?)`,
      [action === "approve" ? "Approved" : "Rejected", subComponentId, today.getFullYear(), today.getMonth()]
    );
    if (!result.changes) return res.status(404).json({ error: "Pending submission row not found." });
    res.json({ success: true, action });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/budget-submissions/current", async (_req, res) => {
  const today = new Date();
  try {
    const rows = await all(
      `SELECT id, month_name AS monthName, submitted_by AS submittedBy, submitted_at AS submittedAt
       FROM budget_submissions
       WHERE year = ? AND month_index = ?
       LIMIT 1`,
      [today.getFullYear(), today.getMonth()]
    );
    const submission = rows[0] || null;
    if (submission) {
      submission.components = await all(
        `SELECT component_id AS componentId, component_name AS componentName
         FROM budget_submission_components
         WHERE submission_id = ?
         ORDER BY component_name`,
        [submission.id]
      );
      submission.rows = await all(
        `SELECT sub_component_id AS subComponentId, component_id AS componentId, review_status AS reviewStatus
         FROM budget_submission_rows WHERE submission_id = ?`,
        [submission.id]
      );
    }
    res.json({ submission });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/budget-submissions/:id/seen", async (req, res) => {
  const submissionId = Number(req.params.id);
  const ceoUsername = String(req.body?.username || "").trim();
  if (!Number.isInteger(submissionId) || submissionId <= 0 || !ceoUsername) {
    return res.status(400).json({ error: "A valid submission and CEO username are required." });
  }

  try {
    await run(
      "INSERT OR IGNORE INTO ceo_submission_views (submission_id, ceo_username) VALUES (?, ?)",
      [submissionId, ceoUsername]
    );
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/view-data/:id/actual-payment-history", async (req, res) => {
  const id = Number(req.params.id);
  if (!id) {
    return res.status(400).json({ error: "Invalid sub-component id." });
  }

  try {
    const statusRows = await all(
      "SELECT approval_status AS approvalStatus FROM sub_components WHERE id = ?",
      [id]
    );

    if (!statusRows.length) {
      return res.status(404).json({ error: "Sub-component not found." });
    }

    if (!isApprovedStatus(statusRows[0].approvalStatus)) {
      return res.status(403).json({ error: "Actual payment history is available only after CEO approval." });
    }

    const history = await all(
      `SELECT id, previous_amount AS previousAmount, adjustment_amount AS adjustmentAmount,
              new_amount AS newAmount, action_type AS actionType, changed_by AS changedBy,
              payment_date AS paymentDate,
              changed_at AS changedAt
       FROM actual_payment_history
       WHERE sub_component_id = ?
       ORDER BY datetime(changed_at) DESC, id DESC`,
      [id]
    );
    res.json({ history });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/view-data/:id/actual-payment", async (req, res) => {
  await ensureColumn("sub_components", "actual_payment", "REAL DEFAULT 0");
  await ensureColumn("sub_components", "last_modified_by", "TEXT");
  await ensureColumn("sub_components", "last_modified_at", "TEXT");
  await ensureColumn("actual_payment_history", "payment_date", "TEXT");

  const id = Number(req.params.id);
  const actionType = String(req.body?.actionType || "add").trim().toLowerCase();
  const amount = Number(req.body?.amount);
  const changedBy = String(req.body?.changedBy || "Unknown").trim();
  const paymentDate = String(req.body?.paymentDate || "").trim();

  if (!id || !Number.isFinite(amount) || amount < 0) {
    return res.status(400).json({ error: "Valid amount is required." });
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(paymentDate) || Number.isNaN(new Date(`${paymentDate}T00:00:00`).getTime())) {
    return res.status(400).json({ error: "A valid payment date is required." });
  }

  try {
    await run("BEGIN TRANSACTION");

    const rows = await all(
      `SELECT actual_payment AS actualPayment,
              COALESCE(proposed_budget, amount_value, 0) AS approvedBudget,
              approval_status AS approvalStatus,
              sc.item_id AS itemId,
              sc.name AS subComponentName,
              sc.component_id AS componentId,
              c.item_id AS componentItemId,
              c.name AS componentName,
              COALESCE(sc.year, CAST(strftime('%Y', sc.created_at) AS INTEGER)) AS year,
              COALESCE(sc.month_index, CAST(strftime('%m', sc.created_at) AS INTEGER) - 1) AS monthIndex,
              COALESCE(sc.month_name,
                CASE CAST(strftime('%m', sc.created_at) AS INTEGER) - 1
                  WHEN 0 THEN 'January' WHEN 1 THEN 'February' WHEN 2 THEN 'March'
                  WHEN 3 THEN 'April' WHEN 4 THEN 'May' WHEN 5 THEN 'June'
                  WHEN 6 THEN 'July' WHEN 7 THEN 'August' WHEN 8 THEN 'September'
                  WHEN 9 THEN 'October' WHEN 10 THEN 'November' WHEN 11 THEN 'December'
                END
              ) AS monthName
       FROM sub_components sc
       LEFT JOIN components c ON c.id = sc.component_id
       WHERE sc.id = ?`,
      [id]
    );

    if (!rows.length) {
      await run("ROLLBACK").catch(() => {});
      return res.status(404).json({ error: "Sub-component not found." });
    }

    if (!isApprovedStatus(rows[0].approvalStatus)) {
      await run("ROLLBACK").catch(() => {});
      return res.status(403).json({ error: "Actual payment can be updated only after CEO approval." });
    }

    const previousAmount = Number(rows[0].actualPayment) || 0;
    let adjustmentAmount = amount;
    let newAmount = previousAmount + amount;

    if (actionType === "subtract") {
      adjustmentAmount = -amount;
      newAmount = previousAmount - amount;
    } else if (actionType === "set") {
      adjustmentAmount = amount - previousAmount;
      newAmount = amount;
    }

    if (newAmount < 0) {
      await run("ROLLBACK").catch(() => {});
      return res.status(400).json({ error: "Actual payment cannot be negative." });
    }

    validateActualPaymentLimit(newAmount, rows[0].approvedBudget);

    const monthStatus = await getCurrentMonthBudgetStatus(Number(rows[0].year), Number(rows[0].monthIndex));
    validateActualPaymentLimit(monthStatus.actualPayment - previousAmount + newAmount, monthStatus.approvedBudget);

    await run(
      `UPDATE sub_components
       SET actual_payment = ?,
           last_modified_by = ?,
           last_modified_at = datetime('now')
       WHERE id = ?`,
      [newAmount, changedBy, id]
    );

    await run(
      `INSERT INTO actual_payment_history
        (sub_component_id, previous_amount, adjustment_amount, new_amount, action_type, changed_by, payment_date)
       VALUES (?, ?, ?, ?, ?, ?, ?)`,
      [id, previousAmount, adjustmentAmount, newAmount, actionType, changedBy, paymentDate]
    );

    await logBudgetChange({
      subComponentId: id,
      componentId: rows[0].componentId,
      componentItemId: rows[0].componentItemId,
      componentName: rows[0].componentName,
      itemId: rows[0].itemId,
      subComponentName: rows[0].subComponentName,
      fieldName: "Actual Payment",
      previousAmount,
      newAmount,
      exceededAmount: Math.max(0, newAmount - (Number(rows[0].approvedBudget) || 0)),
      changedBy,
      year: Number(rows[0].year),
      monthIndex: Number(rows[0].monthIndex),
      monthName: rows[0].monthName || getMonthName(Number(rows[0].monthIndex))
    });

    await run("COMMIT");

    const history = await all(
      `SELECT id, previous_amount AS previousAmount, adjustment_amount AS adjustmentAmount,
              new_amount AS newAmount, action_type AS actionType, changed_by AS changedBy,
              payment_date AS paymentDate,
              changed_at AS changedAt
       FROM actual_payment_history
       WHERE sub_component_id = ?
       ORDER BY datetime(changed_at) DESC, id DESC`,
      [id]
    );

    res.json({ actualPayment: newAmount, history, message: "Actual payment updated." });
  } catch (error) {
    await run("ROLLBACK").catch(() => {});
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/view-data/current-month", async (req, res) => {
  await ensureColumn("sub_components", "actual_payment", "REAL DEFAULT 0");

  const today = new Date();
  const year = today.getFullYear();
  const monthIndex = today.getMonth();

  try {
    // If there are no rows for the current month, import previous month's components and sub_components
    const currentCountRows = await all(
      `SELECT COUNT(1) AS cnt FROM sub_components WHERE COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ? AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?`,
      [year, monthIndex]
    );

    const currentCount = Number(currentCountRows && currentCountRows[0] && currentCountRows[0].cnt) || 0;

    if (currentCount === 0) {
      const prevMonthIndex = monthIndex - 1 >= 0 ? monthIndex - 1 : 11;
      const prevYear = monthIndex - 1 >= 0 ? year : year - 1;

      const prevCountRows = await all(
        `SELECT COUNT(1) AS cnt FROM sub_components WHERE COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ? AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?`,
        [prevYear, prevMonthIndex]
      );

      const prevCount = Number(prevCountRows && prevCountRows[0] && prevCountRows[0].cnt) || 0;

      if (prevCount > 0) {
        await run("BEGIN TRANSACTION");

        // 1) Ensure components from previous month exist for current month (match by name)
        await run(
          `
            INSERT INTO components (item_id, name, year, month_index, month_name)
            SELECT item_id, name, ?, ?, ?
            FROM components
            WHERE COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ?
              AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?
            ON CONFLICT(name, year, month_index) DO UPDATE SET
              item_id = excluded.item_id,
              month_name = excluded.month_name
          `,
          [year, monthIndex, getMonthName(monthIndex), prevYear, prevMonthIndex]
        );

        // 2) Copy sub_components from previous month, mapping to the components for the current month by name
        await run(
          `
            INSERT INTO sub_components
              (item_id, component_id, name, amount_value, amount_cell, existing_liability, estimated_liability, total_estimated_liability, proposed_budget, balance_liability, actual_payment, approval_status, ceo_remarks, approved_by, approved_at, rejected_by, rejected_at, year, month_index, month_name, created_at)
            SELECT
              prev_sc.item_id,
              new_c.id,
              prev_sc.name,
              0,
              prev_sc.amount_cell,
              0,
              0,
              0,
              0,
              0,
              0,
              'Pending',
              NULL,
              NULL,
              NULL,
              NULL,
              NULL,
              ?, ?, ?, CURRENT_TIMESTAMP
            FROM sub_components prev_sc
            JOIN components prev_c ON prev_sc.component_id = prev_c.id
            JOIN components new_c ON new_c.name = prev_c.name
              AND COALESCE(new_c.year, CAST(strftime('%Y', new_c.created_at) AS INTEGER)) = ?
              AND COALESCE(new_c.month_index, CAST(strftime('%m', new_c.created_at) AS INTEGER) - 1) = ?
            WHERE COALESCE(prev_sc.year, CAST(strftime('%Y', prev_sc.created_at) AS INTEGER)) = ?
              AND COALESCE(prev_sc.month_index, CAST(strftime('%m', prev_sc.created_at) AS INTEGER) - 1) = ?
          `,
          [year, monthIndex, getMonthName(monthIndex), year, monthIndex, prevYear, prevMonthIndex]
        );

        await run("COMMIT");
      }
    }

    const rows = await all(
      `
        ${subComponentSelect}
        WHERE COALESCE(sc.year, CAST(strftime('%Y', sc.created_at) AS INTEGER)) = ?
          AND COALESCE(sc.month_index, CAST(strftime('%m', sc.created_at) AS INTEGER) - 1) = ?
        ORDER BY CAST(c.item_id AS INTEGER), c.name, sc.id
      `,
      [year, monthIndex]
    );

    const components = await all(
      `
        SELECT
          id AS componentId,
          item_id AS componentItemId,
          name AS componentName,
          year,
          month_index AS monthIndex,
          month_name AS monthName
        FROM components
        WHERE COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ?
          AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?
        ORDER BY CAST(item_id AS INTEGER), name
      `,
      [year, monthIndex]
    );

    const budgetApproved = getRowsBudgetApproved(rows);

    res.json({
      year,
      monthIndex,
      monthName: getMonthName(monthIndex),
      rows: maskRowsForPendingBudget(rows, budgetApproved),
      components,
      budgetApproved
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/view-data/current-month/download", async (req, res) => {
  const today = new Date();
  const year = today.getFullYear();
  const monthIndex = today.getMonth();
  const monthName = getMonthName(monthIndex);

  try {
    const rows = await all(
      `
        ${subComponentSelect}
        WHERE COALESCE(sc.year, CAST(strftime('%Y', sc.created_at) AS INTEGER)) = ?
          AND COALESCE(sc.month_index, CAST(strftime('%m', sc.created_at) AS INTEGER) - 1) = ?
        ORDER BY CAST(c.item_id AS INTEGER), c.name, sc.id
      `,
      [year, monthIndex]
    );

    await sendBudgetWorkbook(res, {
      title: `DETAILED PAYMENT BUDGET - ${monthName.toUpperCase()} ${year}`,
      fileName: `view-data-${monthName}-${year}.xlsx`,
      rows
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/reports/months", async (req, res) => {
  try {
    const rows = await all(`
      SELECT
        year,
        monthIndex,
        monthName,
        COUNT(*) AS totalSubComponents
      FROM (
        ${subComponentSelect}
      )
      GROUP BY year, monthIndex, monthName
      ORDER BY year DESC, monthIndex DESC
    `);

    res.json(rows);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/reports", async (req, res) => {
  const year = Number(req.query.year);
  const monthIndex = Number(req.query.monthIndex);

  if (!year || Number.isNaN(monthIndex)) {
    return res.status(400).json({ error: "Both year and monthIndex are required." });
  }

  try {
    const rows = await all(
      `
        ${subComponentSelect}
        WHERE COALESCE(sc.year, CAST(strftime('%Y', sc.created_at) AS INTEGER)) = ?
          AND COALESCE(sc.month_index, CAST(strftime('%m', sc.created_at) AS INTEGER) - 1) = ?
        ORDER BY CAST(c.item_id AS INTEGER), c.name, sc.id
      `,
      [year, monthIndex]
    );

    const components = await all(
      `
        SELECT
          id AS componentId,
          item_id AS componentItemId,
          name AS componentName,
          year,
          month_index AS monthIndex,
          month_name AS monthName
        FROM components
        WHERE COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ?
          AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?
        ORDER BY CAST(item_id AS INTEGER), name
      `,
      [year, monthIndex]
    );

    const budgetApproved = getRowsBudgetApproved(rows);

    res.json({
      year,
      monthIndex,
      monthName: rows[0]?.monthName || getMonthName(monthIndex),
      rows: maskRowsForPendingBudget(rows, budgetApproved),
      components,
      budgetApproved
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/view-data/data-logs", async (_req, res) => {
  const today = new Date();
  const year = today.getFullYear();
  const monthIndex = today.getMonth();

  try {
    await ensureLogTables();
    const revisionLogs = await all(
      `SELECT rl.id,
              rl.component_id AS componentId,
              rl.component_item_id AS componentItemId,
              rl.component_name AS componentName,
              rl.reason,
              rl.revised_by AS revisedBy,
              rl.revised_at AS revisedAt,
              (
                SELECT COUNT(1)
                FROM revision_logs prior
                WHERE prior.component_id = rl.component_id
                  AND prior.year = rl.year
                  AND prior.month_index = rl.month_index
                  AND datetime(prior.revised_at) <= datetime(rl.revised_at)
              ) AS revisionNumber
       FROM revision_logs rl
       WHERE rl.year = ?
         AND rl.month_index = ?
       ORDER BY datetime(rl.revised_at) DESC, rl.id DESC`,
      [year, monthIndex]
    );

    const remarkRevisionLogs = await all(
      `SELECT
              MIN(sc.id) AS id,
              sc.component_id AS componentId,
              c.item_id AS componentItemId,
              c.name AS componentName,
              sc.ceo_remarks AS reason,
              COALESCE(sc.rejected_by, 'CEO') AS revisedBy,
              MAX(sc.rejected_at) AS revisedAt
       FROM sub_components sc
       JOIN components c ON c.id = sc.component_id
       WHERE COALESCE(sc.year, CAST(strftime('%Y', sc.created_at) AS INTEGER)) = ?
         AND COALESCE(sc.month_index, CAST(strftime('%m', sc.created_at) AS INTEGER) - 1) = ?
         AND TRIM(COALESCE(sc.ceo_remarks, '')) <> ''
       GROUP BY sc.component_id, c.item_id, c.name, sc.ceo_remarks, sc.rejected_by
       ORDER BY datetime(MAX(sc.rejected_at)) DESC, MIN(sc.id) DESC`,
      [year, monthIndex]
    );

    const combinedRevisionLogs = [...revisionLogs];
    const revisionKeys = new Set(
      combinedRevisionLogs.map((log) => [
        log.componentId,
        String(log.reason || "").trim(),
        String(log.revisedAt || "").trim()
      ].join("|"))
    );

    remarkRevisionLogs.forEach((log) => {
      const key = [
        log.componentId,
        String(log.reason || "").trim(),
        String(log.revisedAt || "").trim()
      ].join("|");
      if (!revisionKeys.has(key)) {
        combinedRevisionLogs.push({
          ...log,
          revisionNumber: combinedRevisionLogs.filter((item) => String(item.componentId) === String(log.componentId)).length + 1
        });
        revisionKeys.add(key);
      }
    });

    const changeRequestLogs = await all(
      `SELECT id,
              sub_component_id AS subComponentId,
              component_id AS componentId,
              component_item_id AS componentItemId,
              component_name AS componentName,
              item_id AS itemId,
              sub_component_name AS subComponentName,
              field_name AS fieldName,
              previous_amount AS previousAmount,
              new_amount AS newAmount,
              change_amount AS changeAmount,
              exceeded_amount AS exceededAmount,
              changed_by AS changedBy,
              changed_at AS changedAt
       FROM budget_change_logs
       WHERE year = ?
         AND month_index = ?
       ORDER BY datetime(changed_at) DESC, id DESC`,
      [year, monthIndex]
    );

    const actualPaymentLogs = await all(
      `SELECT aph.id,
              aph.sub_component_id AS subComponentId,
              sc.component_id AS componentId,
              c.item_id AS componentItemId,
              c.name AS componentName,
              sc.item_id AS itemId,
              sc.name AS subComponentName,
              'Actual Payment' AS fieldName,
              aph.previous_amount AS previousAmount,
              aph.new_amount AS newAmount,
              aph.adjustment_amount AS changeAmount,
              MAX(0, aph.new_amount - COALESCE(sc.proposed_budget, sc.amount_value, 0)) AS exceededAmount,
              aph.changed_by AS changedBy,
              aph.changed_at AS changedAt
       FROM actual_payment_history aph
       JOIN sub_components sc ON sc.id = aph.sub_component_id
       JOIN components c ON c.id = sc.component_id
       WHERE COALESCE(sc.year, CAST(strftime('%Y', sc.created_at) AS INTEGER)) = ?
         AND COALESCE(sc.month_index, CAST(strftime('%m', sc.created_at) AS INTEGER) - 1) = ?
       ORDER BY datetime(aph.changed_at) DESC, aph.id DESC`,
      [year, monthIndex]
    );

    const mergedChangeLogs = [...changeRequestLogs];
    const changeKeys = new Set(
      mergedChangeLogs.map((log) => [
        log.subComponentId,
        log.fieldName,
        String(log.previousAmount),
        String(log.newAmount),
        String(log.changedAt || "")
      ].join("|"))
    );

    actualPaymentLogs.forEach((log) => {
      const key = [
        log.subComponentId,
        log.fieldName,
        String(log.previousAmount),
        String(log.newAmount),
        String(log.changedAt || "")
      ].join("|");
      if (!changeKeys.has(key)) {
        mergedChangeLogs.push(log);
        changeKeys.add(key);
      }
    });

    mergedChangeLogs.sort((a, b) => {
      const bTime = new Date(b.changedAt || 0).getTime();
      const aTime = new Date(a.changedAt || 0).getTime();
      return bTime - aTime;
    });

    combinedRevisionLogs.sort((a, b) => {
      const bTime = new Date(b.revisedAt || 0).getTime();
      const aTime = new Date(a.revisedAt || 0).getTime();
      return bTime - aTime;
    });

    const revisionIds = combinedRevisionLogs
      .map((log) => Number(log.id))
      .filter((id) => Number.isInteger(id) && id > 0);

    let revisionChangeDetails = [];
    if (revisionIds.length > 0) {
      const placeholders = revisionIds.map(() => "?").join(",");
      revisionChangeDetails = await all(
        `SELECT id,
                revision_log_id AS revisionLogId,
                component_id AS componentId,
                component_item_id AS componentItemId,
                component_name AS componentName,
                item_id AS itemId,
                sub_component_name AS subComponentName,
                change_type AS changeType,
                field_label AS fieldLabel,
                previous_value AS previousValue,
                new_value AS newValue,
                changed_at AS changedAt
         FROM revision_change_logs
         WHERE revision_log_id IN (${placeholders})
         ORDER BY id`,
        revisionIds
      );
    }

    combinedRevisionLogs.forEach((log) => {
      log.changeDetails = revisionChangeDetails.filter((detail) => Number(detail.revisionLogId) === Number(log.id));
    });

    res.json({ revisionLogs: combinedRevisionLogs, changeRequestLogs: mergedChangeLogs });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/reports/download", async (req, res) => {
  const year = Number(req.query.year);
  const monthIndex = Number(req.query.monthIndex);

  if (!year || Number.isNaN(monthIndex)) {
    return res.status(400).json({ error: "Both year and monthIndex are required." });
  }

  try {
    const rows = await all(
      `
        ${subComponentSelect}
        WHERE COALESCE(sc.year, CAST(strftime('%Y', sc.created_at) AS INTEGER)) = ?
          AND COALESCE(sc.month_index, CAST(strftime('%m', sc.created_at) AS INTEGER) - 1) = ?
        ORDER BY CAST(c.item_id AS INTEGER), c.name, sc.id
      `,
      [year, monthIndex]
    );
    const components = await all(
      `
        SELECT
          id AS componentId,
          item_id AS componentItemId,
          name AS componentName,
          year,
          month_index AS monthIndex,
          month_name AS monthName
        FROM components
        WHERE COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ?
          AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?
        ORDER BY CAST(item_id AS INTEGER), name
      `,
      [year, monthIndex]
    );
    const monthName = rows[0]?.monthName || getMonthName(monthIndex);

    await sendBudgetWorkbook(res, {
      title: `DETAILED PAYMENT BUDGET - ${monthName.toUpperCase()} ${year}`,
      fileName: `report-${monthName}-${year}.xlsx`,
      rows,
      components
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});
app.post("/api/approval/approve", async (req, res) => {

    try {
        const today = new Date();
        const year = today.getFullYear();
        const monthIndex = today.getMonth();

        await run(`
            UPDATE sub_components
            SET
                approval_status = 'Approved',
                approved_by = 'CEO',
                approved_at = datetime('now'),
                ceo_remarks = NULL,
                rejected_by = NULL,
                rejected_at = NULL,
                last_modified_by = 'CEO',
                last_modified_at = datetime('now')
            WHERE COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ?
              AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?
        `, [year, monthIndex]);

        await run(
            `DELETE FROM budget_submission_components
             WHERE submission_id = (SELECT id FROM budget_submissions WHERE year = ? AND month_index = ?)`,
            [year, monthIndex]
        );
        await run(
            `DELETE FROM budget_submission_rows
             WHERE submission_id = (SELECT id FROM budget_submissions WHERE year = ? AND month_index = ?)`,
            [year, monthIndex]
        );

        res.json({
            success: true,
            message: "Finance Data Approved Successfully."
        });

    } catch (error) {

        res.status(500).json({
            error: error.message
        });

    }

});
app.post("/api/approval/reject", async (req, res) => {

    const { remarks } = req.body;
    const revisionChanges = Array.isArray(req.body?.changes) ? req.body.changes : [];

    if (!remarks) {
        return res.status(400).json({
            error: "CEO Remarks are required."
        });
    }

    try {
        await ensureLogTables();
        const today = new Date();
        const year = today.getFullYear();
        const monthIndex = today.getMonth();
        const monthName = getMonthName(monthIndex);

        await run("BEGIN TRANSACTION");
        await run(
            `
            UPDATE sub_components
            SET
                approval_status = 'Rejected - Finance Admin Revision',
                ceo_remarks = ?,
                rejected_by = 'CEO',
                rejected_at = datetime('now'),
                last_modified_by = 'CEO',
                last_modified_at = datetime('now')
            WHERE COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) = ?
              AND COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) = ?
            `,
            [remarks, year, monthIndex]
        );

        const components = await all(
            `SELECT DISTINCT
                    sc.component_id AS componentId,
                    c.item_id AS componentItemId,
                    c.name AS componentName
             FROM sub_components sc
             JOIN components c ON c.id = sc.component_id
             WHERE COALESCE(sc.year, CAST(strftime('%Y', sc.created_at) AS INTEGER)) = ?
               AND COALESCE(sc.month_index, CAST(strftime('%m', sc.created_at) AS INTEGER) - 1) = ?
             ORDER BY CAST(c.item_id AS INTEGER), c.name`,
            [year, monthIndex]
        );

        for (const component of components) {
            const insertRevision = await run(
                `INSERT INTO revision_logs
                  (component_id, component_item_id, component_name, reason, revised_by, year, month_index, month_name)
                 VALUES (?, ?, ?, ?, 'CEO', ?, ?, ?)`,
                [
                    component.componentId,
                    component.componentItemId || "",
                    component.componentName || "",
                    remarks,
                    year,
                    monthIndex,
                    monthName
                ]
            );

            const componentItemId = String(component.componentItemId || "");
            const matchingChanges = revisionChanges.filter((change) => {
                const item = String(change?.item || "");
                return item === componentItemId || item.startsWith(`${componentItemId}.`);
            });

            for (const change of matchingChanges) {
                const fields = Array.isArray(change.fields) && change.fields.length > 0
                    ? change.fields
                    : [{ label: "Row", from: change.name || "", to: change.type || "Changed" }];

                for (const field of fields) {
                    await run(
                        `INSERT INTO revision_change_logs
                          (revision_log_id, component_id, component_item_id, component_name, item_id, sub_component_name,
                           change_type, field_label, previous_value, new_value, year, month_index, month_name)
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                        [
                            insertRevision.lastID,
                            component.componentId,
                            component.componentItemId || "",
                            component.componentName || "",
                            String(change.item || ""),
                            String(change.name || ""),
                            String(change.type || "Changed"),
                            String(field.label || ""),
                            String(field.from || ""),
                            String(field.to || ""),
                            year,
                            monthIndex,
                            monthName
                        ]
                    );
                }
            }
        }

        await run("COMMIT");

        res.json({
            success: true,
            message: "Finance Data Rejected and sent to Finance Admin for revision."
        });

    } catch (error) {
        await run("ROLLBACK").catch(() => {});

        res.status(500).json({
            error: error.message
        });

    }

});
app.get("/api/approval/status", async (req, res) => {

    try {

        const rows = await all(`
            SELECT
                id,
                approval_status,
                ceo_remarks,
                approved_by,
               approved_at,
                rejected_by,
                rejected_at
            FROM sub_components
        `);

        res.json(rows);

    } catch (err) {

        res.status(500).json({
            error: err.message
        });

    }

});

app.get("/api/dashboard/filter-options", async (_req, res) => {
  try {
    const periodRows = await all(`
      SELECT DISTINCT
        COALESCE(sc.year, CAST(strftime('%Y', sc.created_at) AS INTEGER)) AS year,
        COALESCE(sc.month_index, CAST(strftime('%m', sc.created_at) AS INTEGER) - 1) AS monthIndex
      FROM sub_components sc
      ORDER BY year DESC, monthIndex DESC
    `);

    const financialYearSet = new Set();
    periodRows.forEach((row) => {
      const year = Number(row.year);
      const monthIndex = Number(row.monthIndex);
      if (!Number.isNaN(year) && !Number.isNaN(monthIndex)) {
        financialYearSet.add(monthIndex >= 3 ? year : year - 1);
      }
    });

    if (financialYearSet.size === 0) {
      financialYearSet.add(getFinancialYearStart());
    }

    const componentRows = await all(`
      SELECT DISTINCT
        c.id AS value,
        c.name AS label,
        COALESCE(sc.year, CAST(strftime('%Y', sc.created_at) AS INTEGER)) AS year,
        COALESCE(sc.month_index, CAST(strftime('%m', sc.created_at) AS INTEGER) - 1) AS monthIndex
      FROM sub_components sc
      JOIN components c ON c.id = sc.component_id
      WHERE c.name IS NOT NULL AND TRIM(c.name) <> ''
      ORDER BY year DESC, monthIndex DESC, c.name
    `);

    const componentGroups = [];
    const groupsByPeriod = new Map();
    componentRows.forEach((row) => {
      const year = Number(row.year);
      const monthIndex = Number(row.monthIndex);
      const key = `${year}-${monthIndex}`;
      if (!groupsByPeriod.has(key)) {
        const group = { year, monthIndex, monthName: getMonthName(monthIndex), components: [] };
        groupsByPeriod.set(key, group);
        componentGroups.push(group);
      }
      groupsByPeriod.get(key).components.push({ value: row.value, label: row.label });
    });

    res.json({
      financialYears: [...financialYearSet].sort((a, b) => b - a).map((startYear) => ({
        value: startYear,
        label: `FY ${startYear}-${String(startYear + 1).slice(-2)}`
      })),
      quarters: [
        { value: "Q1", label: "Q1", months: [3, 4, 5] },
        { value: "Q2", label: "Q2", months: [6, 7, 8] },
        { value: "Q3", label: "Q3", months: [9, 10, 11] },
        { value: "Q4", label: "Q4", months: [0, 1, 2] }
      ],
      months: [3, 4, 5, 6, 7, 8, 9, 10, 11, 0, 1, 2].map((monthIndex) => ({
        value: monthIndex,
        label: getMonthName(monthIndex)
      })),
      componentGroups
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/dashboard/stats", async (req, res) => {
  const filters = buildDashboardFilters(req.query);
  const activePeriod = filters.hasFilters
    ? {
        year: filters.financialYear,
        monthIndex: null,
        monthName: `FY ${filters.financialYear}-${String(filters.financialYear + 1).slice(-2)}`
      }
    : {
        year: filters.periods[0]?.year,
        monthIndex: filters.periods[0]?.monthIndex,
        monthName: filters.periods[0]?.monthName
      };
  const summaryWhere = appendDashboardWhere(filters, "sc");

  try {
    // 1. Current month summary
    const summaryRows = await all(`
      SELECT
        COUNT(1) AS totalRows,
        SUM(CASE WHEN approval_status = 'Approved' THEN 1 ELSE 0 END) AS approvedRows,
        SUM(CASE WHEN approval_status LIKE 'Rejected%' THEN 1 ELSE 0 END) AS rejectedRows,
        COALESCE(SUM(COALESCE(proposed_budget, amount_value, 0)), 0) AS proposedBudget,
        COALESCE(SUM(COALESCE(existing_liability, 0)), 0) AS existingLiability,
        COALESCE(SUM(COALESCE(estimated_liability, 0)), 0) AS estimatedLiability,
        COALESCE(SUM(COALESCE(total_estimated_liability, 0)), 0) AS totalLiability,
        COALESCE(SUM(COALESCE(actual_payment, 0)), 0) AS actualPayment,
        COALESCE(SUM(COALESCE(balance_liability, 0)), 0) AS balanceLiability
      FROM sub_components
      JOIN components c ON c.id = sub_components.component_id
      ${summaryWhere.where.replaceAll("sc.", "sub_components.")}
    `, summaryWhere.params);

    const summary = summaryRows[0] || {};
    const totalRows = Number(summary.totalRows) || 0;
    const approvedRows = Number(summary.approvedRows) || 0;
    const rejectedRows = Number(summary.rejectedRows) || 0;

    let approvalStatus = 'No Data';
    if (totalRows > 0) {
      if (approvedRows === totalRows) {
        approvalStatus = 'Approved';
      } else if (rejectedRows > 0) {
        approvalStatus = 'Rejected - Finance Admin Revision';
      } else {
        approvalStatus = 'Pending';
      }
    }

    // Calculate Proposed vs Approved for current month
    let proposedBudgetSum = 0;
    let approvedBudgetSum = 0;

    if (totalRows > 0) {
      approvedBudgetSum = approvalStatus === 'Approved' ? Number(summary.proposedBudget) : 0;

      // Find initial proposed budget values
      const subRows = await all(`
        SELECT
          sc.id,
          sc.proposed_budget AS currentProposed,
          (SELECT new_amount FROM budget_change_logs WHERE sub_component_id = sc.id AND field_name = 'Approved Budget' ORDER BY id ASC LIMIT 1) AS initialInsertVal,
          (SELECT previous_amount FROM budget_change_logs WHERE sub_component_id = sc.id AND field_name = 'proposedBudget' ORDER BY id ASC LIMIT 1) AS initialUpdateVal
        FROM sub_components sc
        JOIN components c ON c.id = sc.component_id
        ${summaryWhere.where}
      `, summaryWhere.params);

      subRows.forEach(row => {
        const val = row.initialInsertVal !== null ? row.initialInsertVal : (row.initialUpdateVal !== null ? row.initialUpdateVal : row.currentProposed);
        proposedBudgetSum += Number(val) || 0;
      });
    }

    // Get CEO remarks if rejected
    let ceoRemarks = null;
    if (approvalStatus === 'Rejected - Finance Admin Revision') {
      const remarksRows = await all(`
        SELECT ceo_remarks
        FROM sub_components sc
        JOIN components c ON c.id = sc.component_id
        ${summaryWhere.where}
          AND ceo_remarks IS NOT NULL
        LIMIT 1
      `, summaryWhere.params);
      if (remarksRows.length > 0) {
        ceoRemarks = remarksRows[0].ceo_remarks;
      }
    }

    // 2. Recent Logs
    await ensureLogTables();
    const logs = await all(`
      SELECT
        id,
        sub_component_name AS subComponentName,
        component_name AS componentName,
        field_name AS fieldName,
        previous_amount AS previousAmount,
        new_amount AS newAmount,
        change_amount AS changeAmount,
        changed_by AS changedBy,
        changed_at AS changedAt
      FROM budget_change_logs
      ORDER BY id DESC
      LIMIT 5
    `);

    // 3. Chart Trend Data (Proposed vs Approved vs Actual)
    const trendMonths = filters.hasFilters ? filters.periods : await all(`
      SELECT DISTINCT
        COALESCE(year, CAST(strftime('%Y', created_at) AS INTEGER)) AS yr,
        COALESCE(month_index, CAST(strftime('%m', created_at) AS INTEGER) - 1) AS mi,
        COALESCE(month_name,
          CASE CAST(strftime('%m', created_at) AS INTEGER) - 1
            WHEN 0 THEN 'January'
            WHEN 1 THEN 'February'
            WHEN 2 THEN 'March'
            WHEN 3 THEN 'April'
            WHEN 4 THEN 'May'
            WHEN 5 THEN 'June'
            WHEN 6 THEN 'July'
            WHEN 7 THEN 'August'
            WHEN 8 THEN 'September'
            WHEN 9 THEN 'October'
            WHEN 10 THEN 'November'
            WHEN 11 THEN 'December'
          END
        ) AS mn
      FROM sub_components
      ORDER BY yr DESC, mi DESC
      LIMIT 6
    `);

    const charts = [];
    for (const m of trendMonths) {
      const yearVal = m.yr ?? m.year;
      const monthIdxVal = m.mi ?? m.monthIndex;
      const monthNameVal = m.mn ?? m.monthName;
      const monthFilters = {
        ...filters,
        periods: [{ year: Number(yearVal), monthIndex: Number(monthIdxVal) }]
      };
      const monthWhere = appendDashboardWhere(monthFilters, "sc");

      const summaryRows = await all(`
        SELECT
          COUNT(1) AS totalRows,
          SUM(CASE WHEN approval_status = 'Approved' THEN 1 ELSE 0 END) AS approvedRows,
          COALESCE(SUM(COALESCE(proposed_budget, amount_value, 0)), 0) AS currentBudget,
          COALESCE(SUM(COALESCE(actual_payment, 0)), 0) AS actualPayment
        FROM sub_components
        JOIN components c ON c.id = sub_components.component_id
        ${monthWhere.where.replaceAll("sc.", "sub_components.")}
      `, monthWhere.params);

      const s = summaryRows[0] || {};
      const totalRows = Number(s.totalRows) || 0;
      const approvedRows = Number(s.approvedRows) || 0;
      const isApproved = totalRows > 0 && approvedRows === totalRows;

      const approvedVal = isApproved ? Number(s.currentBudget) : 0;
      const actualVal = Number(s.actualPayment) || 0;

      // Find proposed budget sum (initial proposed budget)
      const subRows = await all(`
        SELECT
          sc.id,
          sc.proposed_budget AS currentProposed,
          (SELECT new_amount FROM budget_change_logs WHERE sub_component_id = sc.id AND field_name = 'Approved Budget' ORDER BY id ASC LIMIT 1) AS initialInsertVal,
          (SELECT previous_amount FROM budget_change_logs WHERE sub_component_id = sc.id AND field_name = 'proposedBudget' ORDER BY id ASC LIMIT 1) AS initialUpdateVal
        FROM sub_components sc
        JOIN components c ON c.id = sc.component_id
        ${monthWhere.where}
      `, monthWhere.params);

      let proposedVal = 0;
      subRows.forEach(row => {
        const val = row.initialInsertVal !== null ? row.initialInsertVal : (row.initialUpdateVal !== null ? row.initialUpdateVal : row.currentProposed);
        proposedVal += Number(val) || 0;
      });

      charts.push({
        year: yearVal,
        monthIndex: monthIdxVal,
        monthName: monthNameVal,
        proposedBudget: proposedVal,
        approvedBudget: approvedVal,
        actualPayment: actualVal
      });
    }

    if (!filters.hasFilters) charts.reverse();
    // 4. Component Budget Summary & Variance Table
    const componentSummary = await all(`
      SELECT
        c.name AS componentName,
        SUM(COALESCE(sc.proposed_budget, sc.amount_value, 0)) AS proposedSum,
        SUM(CASE WHEN sc.approval_status = 'Approved' THEN COALESCE(sc.proposed_budget, sc.amount_value, 0) ELSE 0 END) AS approvedSum,
        SUM(COALESCE(sc.actual_payment, 0)) AS actualSum
      FROM sub_components sc
      LEFT JOIN components c ON c.id = sc.component_id
      ${summaryWhere.where}
      GROUP BY c.name
      ORDER BY proposedSum DESC
    `, summaryWhere.params);

    // 5. User Summary
    const userSummaryRows = await all(`
      SELECT role, COUNT(1) AS count FROM users GROUP BY role
    `);

    res.json({
      activePeriod,
      appliedFilters: {
        financialYear: filters.financialYear,
        quarters: filters.quarters,
        months: filters.months,
        components: filters.componentIds
      },
      summary: {
        proposedBudget: proposedBudgetSum,
        approvedBudget: approvedBudgetSum,
        existingLiability: Number(summary.existingLiability) || 0,
        estimatedLiability: Number(summary.estimatedLiability) || 0,
        totalLiability: Number(summary.totalLiability) || 0,
        actualPayment: Number(summary.actualPayment) || 0,
        balanceLiability: Number(summary.balanceLiability) || 0,
        approvalStatus,
        ceoRemarks
      },
      recentLogs: logs,
      charts,
      componentSummary,
      userSummary: userSummaryRows
    });  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`Finance API running at http://localhost:${PORT}`);
  console.log(`SQL database file: ${dbPath}`);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
});

process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  process.exit(1);
});
