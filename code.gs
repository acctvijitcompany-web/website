const SHEET_ID  = '1RN17Isv6boXqbvUhhoGkNDgLOTZrycft71PL1PDz4Z0';
const SHEET_NAME = 'Appointments';

// ใช้/สร้างชีตเก็บข้อมูลนัดหมาย
function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.clear();
    sheet.appendRow([
      'Timestamp',        // 0
      'CustomerName',     // 1
      'Phone',            // 2
      'ServiceType',      // 3
      'AppointmentDate',  // 4
      'AppointmentTime',  // 5
      'Notes',            // 6
      'Status',           // 7
      'AppointmentId'     // 8
    ]);
  }

  return sheet;
}

// แสดงหน้าเว็บ
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('บริษัท วิจิตร เอ แอนด์ ที จำกัด - บริการครบวงจร')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ดึงรายการนัดหมายทั้งหมด
function getAppointments() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();

  const dataRows = values.slice(1); // ข้ามหัวตาราง

  const appointments = dataRows.map(row => {
    const createdAt = row[0] instanceof Date ? row[0].toISOString() : '';

    return {
      id: row[8] || '',
      customer_name: row[1] || '',
      phone: row[2] || '',
      service_type: row[3] || '',
      appointment_date: row[4] || '',
      appointment_time: row[5] || '',
      notes: row[6] || '',
      status: row[7] || '',
      created_at: createdAt
    };
  });

  return appointments;
}

// บันทึกนัดหมายใหม่
function saveAppointment(appointment) {
  const sheet = getSheet();
  const id = Utilities.getUuid();
  const now = new Date();

  const status = appointment.status || 'รอยืนยัน';

  sheet.appendRow([
    now,                                 // Timestamp
    appointment.customer_name || '',     // CustomerName
    appointment.phone || '',             // Phone
    appointment.service_type || '',      // ServiceType
    appointment.appointment_date || '',  // AppointmentDate
    appointment.appointment_time || '',  // AppointmentTime
    appointment.notes || '',             // Notes
    status,                              // Status
    id                                   // AppointmentId
  ]);

  return {
    isOk: true,
    id: id
  };
}

// ลบนัดหมายจาก id
function deleteAppointment(id) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { isOk: false, message: 'ไม่มีข้อมูล' };
  }

  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = range.getValues();

  let rowIndexToDelete = -1;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const appointmentId = row[8]; // คอลัมน์ AppointmentId
    if (appointmentId === id) {
      rowIndexToDelete = i + 2; // +2 เพราะเริ่มที่แถว 2
      break;
    }
  }

  if (rowIndexToDelete > -1) {
    sheet.deleteRow(rowIndexToDelete);
    return { isOk: true };
  } else {
    return { isOk: false, message: 'ไม่พบข้อมูลนัดหมาย' };
  }
}
