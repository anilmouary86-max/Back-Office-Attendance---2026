const ADMIN_PIN = "2408";
const VALID_SAVE_STATUSES = ["WFH", "CL", "SL"];
const VALID_REVIEW_STATUSES = ["WFH", "A", "CL", "SL"];

const EMPLOYEES = [
  ["4387", "MB003783", "Dr. Suresh Mathur", "CMO"],
  ["6291", "MB003784", "Vijay Shrivastava", "Manager MIS"],
  ["2748", "MB003724", "Vivek Jadoun", "MIS"],
  ["9156", "MB003725", "Anil Mourya", "MIS"],
  ["5832", "MB003730", "Dr. Pradeep Kumar", "Chief Investigation Officer"],
  ["7604", "", "Dr.Mahesh Kumar Meena", "Filed Coordinator"],
  ["2218", "MB003732", "Dr. Rohit Mobiya", "Desk Audit"],
  ["8462", "MB003716", "Dr. Shivani Joshi", "Desk Audit"],
  ["3059", "MB003781", "Dr. Swati Agarwal", "Desk Audit"],
  ["6715", "MB003728", "Dr. Priyanka Pratihar", "Desk Audit"],
  ["4206", "MB003782", "Akshay Bhan Singh", "Sr. Executve"],
  ["7394", "MOAD0033", "Ravi Arya", "Scanning"],
  ["2581", "", "Vijay Kumar Choudhary", "Back Office"],
  ["6047", "MB003842", "Neetu Kumawat", "Executive"],
  ["8173", "MOAD0045", "Jitendra Meena", "Call Center"],
  ["3960", "MOAD0046", "Santosh Bairwa", "Call Center"],
  ["9425", "MOAD0047", "Mantash Sharma", "Call Center"],
  ["1537", "MOAD0048", "Badal Singh Sehara", "Call Center"],
  ["5086", "", "Deepak Kumar Prajapat", "Call Center"],
  ["7812", "MB003799", "Dr. Komal Sharma", "Desk Audit"],
  ["2369", "MB003800", "Dr. Beena Mathur", "Desk Audit"],
  ["8940", "MB003801", "Dr. Sunita Kushwaha", "Desk Audit"],
  ["4671", "MB003818", "Dr. Shivangi Maurya", "Desk Audit"],
  ["3208", "MB003841", "Dr. Pushpendra Kumar Kauraiya", "Desk Audit"],
  ["6592", "", "Dr. Dharmendra", "Desk Audit"],
  ["1047", "", "Dr. Rekha Dhindwal", "Desk Audit"],
  ["5729", "MB003819", "Vanshika Goswami", "Executive - MIS"],
  ["9381", "MB003821", "Rakesh Kumar Jat", "Executive"],
  ["2156", "MOAD0069", "Hemraj Meena", "Office Boy"],
  ["6843", "-", "Brijesh Kumar Prajapat", "Reatil"],
  ["7490", "-", "Vijay Hathinya", "Reatil"],
];

function doGet(e) {
  setupSheets_();
  const params = e.parameter || {};
  const callback = params.callback || "callback";
  const action = params.action || "";

  let result;
  try {
    if (action === "employee") result = employee_(params.pin);
    else if (action === "status") result = status_(params.pin, params.date);
    else if (action === "save") result = save_(params.pin, params.date, params.status);
    else if (action === "admin") result = admin_(params.pin);
    else if (action === "approve") result = review_(params.adminPin, params.targetPin, params.date, "Approved");
    else if (action === "cancel") result = review_(params.adminPin, params.targetPin, params.date, "Cancelled");
    else if (action === "clearAll") result = clearAll_(params.adminPin, params.confirm);
    else result = { ok: false, error: "Action valid nahi hai." };
  } catch (error) {
    result = { ok: false, error: String(error && error.message ? error.message : error) };
  }

  return ContentService
    .createTextOutput(`${callback}(${JSON.stringify(result)})`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function setupSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let employeesSheet = ss.getSheetByName("Employees");
  if (!employeesSheet) employeesSheet = ss.insertSheet("Employees");
  employeesSheet.clear();
  employeesSheet.getRange(1, 1, 1, 4).setValues([["PIN", "Employee Code", "Employee Name", "Designation"]]);
  employeesSheet.getRange(2, 1, EMPLOYEES.length, 4).setValues(EMPLOYEES);

  let attendanceSheet = ss.getSheetByName("Attendance");
  if (!attendanceSheet) attendanceSheet = ss.insertSheet("Attendance");
  const headers = ["PIN", "Employee Code", "Employee Name", "Designation", "Date", "Status", "Approval", "Saved At", "Reviewed At", "Reviewed By"];
  if (attendanceSheet.getLastRow() === 0) {
    attendanceSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const currentHeaders = attendanceSheet.getRange(1, 1, 1, Math.max(attendanceSheet.getLastColumn(), headers.length)).getValues()[0];
    headers.forEach((header, index) => {
      if (currentHeaders[index] !== header) attendanceSheet.getRange(1, index + 1).setValue(header);
    });
  }
}

function findEmployee_(pin) {
  const cleanPin = String(pin || "").trim();
  for (const row of EMPLOYEES) {
    if (row[0] === cleanPin) {
      return { pin: row[0], code: row[1], name: row[2], department: row[3] };
    }
  }
  return null;
}

function employee_(pin) {
  const employee = findEmployee_(pin);
  if (!employee) return { ok: false, error: "PIN galat hai." };
  return { ok: true, employee: publicEmployee_(employee) };
}

function status_(pin, date) {
  const employee = findEmployee_(pin);
  if (!employee) return { ok: false, error: "PIN galat hai." };
  if (!/^2026-08-(0[1-9]|[12][0-9]|30)$/.test(String(date || ""))) {
    return { ok: false, error: "August ki valid date select karo." };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  const values = sheet.getDataRange().getValues();
  let entry = null;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === employee.pin && formatDate_(values[i][4]) === date) {
      entry = {
        status: String(values[i][5] || ""),
        approval: String(values[i][6] || "Pending"),
        savedAt: formatDateTime_(values[i][7]),
        reviewedAt: formatDateTime_(values[i][8]),
        reviewedBy: String(values[i][9] || ""),
      };
    }
  }

  return { ok: true, employee: publicEmployee_(employee), date, entry };
}

function save_(pin, date, status) {
  const employee = findEmployee_(pin);
  if (!employee) return { ok: false, error: "PIN galat hai." };
  if (!/^2026-08-(0[1-9]|[12][0-9]|30)$/.test(String(date || ""))) {
    return { ok: false, error: "August ki valid date select karo." };
  }
  if (!VALID_SAVE_STATUSES.includes(status)) {
    return { ok: false, error: "Status valid nahi hai." };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  const values = sheet.getDataRange().getValues();
  let updateRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === employee.pin && formatDate_(values[i][4]) === date) {
      updateRow = i + 1;
      break;
    }
  }

  const approval = "Pending";
  const row = [employee.pin, employee.code, employee.name, employee.department, date, status, approval, new Date(), "", ""];
  if (updateRow > 0) {
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === employee.pin && formatDate_(values[i][4]) === date) {
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      }
    }
  } else {
    sheet.appendRow(row);
  }

  return { ok: true, employee: publicEmployee_(employee), date, status, approval, message: "Pending for CMO approval" };
}

function review_(adminPin, targetPin, date, decision) {
  if (String(adminPin || "").trim() !== ADMIN_PIN) return { ok: false, error: "Admin PIN galat hai." };
  const employee = findEmployee_(targetPin);
  if (!employee) return { ok: false, error: "Employee PIN valid nahi hai." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  const values = sheet.getDataRange().getValues();
  let found = false;
  let currentStatus = "";
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === employee.pin && formatDate_(values[i][4]) === date) {
      const status = String(values[i][5] || "");
      if (!VALID_REVIEW_STATUSES.includes(status)) return { ok: false, error: "Status valid nahi hai." };
      currentStatus = status;
      sheet.getRange(i + 1, 7).setValue(decision);
      sheet.getRange(i + 1, 9).setValue(new Date());
      sheet.getRange(i + 1, 10).setValue("CMO/Admin");
      found = true;
    }
  }
  if (found) return { ok: true, message: `${employee.name} ${date} ${currentStatus} ${decision}.` };
  return { ok: false, error: "Attendance record nahi mila." };
}

function clearAll_(adminPin, confirmText) {
  if (String(adminPin || "").trim() !== ADMIN_PIN) return { ok: false, error: "Admin PIN galat hai." };
  if (String(confirmText || "").trim().toUpperCase() !== "CLEAR") {
    return { ok: false, error: "Clear karne ke liye CLEAR type karo." };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  if (!sheet) return { ok: true, message: "Attendance already clear hai.", cleared: 0 };

  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), 10);
  if (lastRow <= 1) return { ok: true, message: "Attendance already clear hai.", cleared: 0 };

  const cleared = lastRow - 1;
  sheet.getRange(2, 1, cleared, lastColumn).clearContent();
  return { ok: true, message: `${cleared} attendance records clear ho gaye.`, cleared };
}

function admin_(pin) {
  if (String(pin || "").trim() !== ADMIN_PIN) return { ok: false, error: "Admin PIN galat hai." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  const data = sheet.getDataRange().getValues().slice(1);
  const grouped = {};
  EMPLOYEES.forEach((row) => {
    grouped[row[0]] = {
      pin: row[0],
      code: row[1],
      name: row[2],
      department: row[3],
      days: {},
      counts: { WFH: 0, CL: 0, SL: 0 },
      pending: 0,
      cancelled: 0,
    };
  });

  data.forEach((row) => {
    const pinValue = String(row[0]);
    const date = formatDate_(row[4]);
    const status = String(row[5] || "");
    const approval = String(row[6] || "Pending");
    const savedAt = formatDateTime_(row[7]);
    const reviewedAt = formatDateTime_(row[8]);
    const reviewedBy = String(row[9] || "");
    if (grouped[pinValue] && date) {
      grouped[pinValue].days[date] = { status, approval, savedAt, reviewedAt, reviewedBy };
    }
  });

  Object.values(grouped).forEach((row) => {
    Object.values(row.days).forEach((entry) => {
      if (entry.approval === "Pending") row.pending++;
      if (entry.approval === "Cancelled") row.cancelled++;
      if (entry.approval === "Approved" && row.counts[entry.status] !== undefined) row.counts[entry.status]++;
    });
  });

  return { ok: true, rows: Object.values(grouped) };
}

function publicEmployee_(employee) {
  return { code: employee.code, name: employee.name, department: employee.department };
}

function formatDate_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value).slice(0, 10);
}

function formatDateTime_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");
  }
  return String(value);
}
