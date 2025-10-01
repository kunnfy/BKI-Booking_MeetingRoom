const SHEET_NAME = "Reservations";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ระบบจองห้องประชุม BKI-Meeting Room');
}

// ดึงรายการเวลาที่ว่างสำหรับห้องและวันที่ที่เลือก
function getAvailableTimes(roomName, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  // สร้าง slot 30 นาทีตั้งแต่ 08:00-17:00
  const allSlots = [];
  for (let h = 8; h < 17; h++) {
    allSlots.push(`${h.toString().padStart(2,'0')}:00`);
    allSlots.push(`${h.toString().padStart(2,'0')}:30`);
  }
  allSlots.push('17:00');

  // ดึง slot ที่ถูกจองของห้องและวันที่นั้น
  const reservedSlots = data
    .filter(row => row[0] === date && row[1] === roomName)
    .map(row => getTimeSlotsBetween(row[2], row[3]))
    .flat();

  // ลบ slot ที่ถูกจอง
  const availableSlots = allSlots.filter(slot => !reservedSlots.includes(slot));
  return availableSlots;
}

// helper function: สร้าง slot ของเวลา
function getTimeSlotsBetween(startTime, endTime) {
  const slots = [];
  let [h, m] = startTime.split(':').map(Number);
  const [endH, endM] = endTime.split(':').map(Number);
  
  while(h < endH || (h === endH && m < endM)) {
    slots.push(`${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}`);
    m += 30;
    if (m >= 60) { h += 1; m = 0; }
  }
  return slots;
}

// บันทึกการจองพร้อมตรวจสอบเวลาซ้ำ
function saveReservation(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const existingData = sheet.getDataRange().getValues();

  function timeToMinutes(t) {
    const [h, m] = t.split(':').map(Number);
    return h * 60 + m;
  }

  const newStart = timeToMinutes(data.startTime);
  const newEnd = timeToMinutes(data.endTime);

  // ตรวจสอบเวลาซ้ำ
  const conflict = existingData.some(row => {
    if (row[0] !== data.date || row[1] !== data.room) return false;
    const existingStart = timeToMinutes(row[2]);
    const existingEnd = timeToMinutes(row[3]);
    return (newStart < existingEnd && newEnd > existingStart); // ทับกัน
  });

  if (conflict) {
    return {status: 'error', message: 'เวลานี้ไม่ว่าง กรุณาเลือกช่วงเวลาอื่น'};
  }

  // บันทึกการจอง
  sheet.appendRow([
    data.date,
    data.room,
    data.startTime,
    data.endTime,
    data.name,
    data.phone
  ]);

  return {status: 'success'};
}

// ดึงการจองทั้งหมดในวันที่เลือก (สำหรับ Calendar View)
function getReservationsByDate(date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  return data
    .filter(row => row[0] === date)
    .map(row => ({
      room: row[1],
      start: row[2],
      end: row[3],
      name: row[4]
    }));
}
