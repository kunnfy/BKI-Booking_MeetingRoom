// ==================== code.gs ====================

// ใส่ Spreadsheet ID ของคุณตรงนี้
const SPREADSHEET_ID = '1Sjjcjdu7L0K89q3_sUcs88LAdNwgWpIA_mhEuNauL5k';

// ฟังก์ชันช่วยเปิดชีท
function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Reservations');
  if (!sheet) {
    sheet = ss.insertSheet('Reservations');
    // เพิ่มหัวคอลัมน์
    sheet.getRange(1, 1, 1, 6).setValues([['วันที่', 'ชื่อห้องที่ต้องการจอง', 'เวลาเริ่มจอง', 'เวลาจบ', 'ชื่อผู้จอง', 'เบอร์ติดต่อกลับ']]);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    // ตั้งค่า format สำหรับคอลัมน์เวลา
    sheet.getRange(2, 3, sheet.getMaxRows() - 1, 2).setNumberFormat('HH:mm');
  }
  return sheet;
}

// ฟังก์ชันสำหรับแสดงหน้าเว็บ
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบจองห้องประชุม')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันสำหรับบันทึกการจอง
function saveReservation(data) {
  try {
    const sheet = getSheet();
    const nextRow = sheet.getLastRow() + 1;

    sheet.getRange(nextRow, 1, 1, 6).setValues([[
      data.date,
      data.room,
      data.startTime,
      data.endTime,
      data.name,
      data.phone
    ]]);

    return {
      success: true,
      message: 'บันทึกการจองเรียบร้อยแล้ว'
    };

  } catch (error) {
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + error.toString()
    };
  }
}

// ฟังก์ชันดึงข้อมูลการจองตามวันที่และห้อง
function getReservationsByDateAndRoom(date, room) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const reservations = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[1] || row[2] === '' || row[3] === '') continue;

      const rowDate = formatDate(new Date(row[0]));
      const roomName = String(row[1]).trim();
      if (rowDate === date && roomName === room) {
        let startTime = row[2];
        let endTime = row[3];

        if (startTime instanceof Date) startTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'HH:mm');
        if (endTime instanceof Date) endTime = Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'HH:mm');

        reservations.push({
          date: rowDate,
          room: roomName,
          startTime,
          endTime,
          name: row[4],
          phone: row[5]
        });
      }
    }

    return reservations;

  } catch (error) {
    console.error('Error getting reservations:', error);
    return [];
  }
}

// ฟังก์ชันดึง time slots ที่ว่าง
function getAvailableTimeSlots(date, room) {
  try {
    const allSlots = [];
    for (let hour = 8; hour <= 16; hour++) {
      for (let minute = 0; minute < 60; minute += 30) {
        const time = `${hour.toString().padStart(2,'0')}:${minute.toString().padStart(2,'0')}`;
        allSlots.push(time);
      }
    }

    const reservations = getReservationsByDateAndRoom(date, room);
    const bookedSlots = new Set();

    reservations.forEach(reservation => {
      const startMinutes = timeToMinutes(reservation.startTime);
      const endMinutes = timeToMinutes(reservation.endTime);

      allSlots.forEach(slot => {
        const slotMinutes = timeToMinutes(slot);
        if (slotMinutes >= startMinutes && slotMinutes < endMinutes) bookedSlots.add(slot);
      });
    });

    return allSlots.filter(slot => !bookedSlots.has(slot));

  } catch (error) {
    console.error('Error in getAvailableTimeSlots:', error);
    return [];
  }
}

// ฟังก์ชันดึงเวลาสิ้นสุดที่เป็นไปได้
function getAvailableEndTimes(date, room, startTime) {
  try {
    if (!startTime) return [];

    const allSlots = [];
    for (let hour = 8; hour <= 17; hour++) {
      for (let minute = 0; minute < 60; minute += 30) {
        if (hour === 17 && minute > 0) break;
        const time = `${hour.toString().padStart(2,'0')}:${minute.toString().padStart(2,'0')}`;
        allSlots.push(time);
      }
    }

    const startMinutes = timeToMinutes(startTime);
    const reservations = getReservationsByDateAndRoom(date, room);

    let nextBookingStart = null;
    reservations.forEach(reservation => {
      const resStartMinutes = timeToMinutes(reservation.startTime);
      if (resStartMinutes > startMinutes) {
        if (!nextBookingStart || resStartMinutes < timeToMinutes(nextBookingStart)) {
          nextBookingStart = reservation.startTime;
        }
      }
    });

    return allSlots.filter(slot => {
      const slotMinutes = timeToMinutes(slot);
      if (slotMinutes <= startMinutes) return false;
      if (nextBookingStart && slotMinutes > timeToMinutes(nextBookingStart)) return false;
      return true;
    });

  } catch (error) {
    console.error('Error in getAvailableEndTimes:', error);
    return [];
  }
}

// ฟังก์ชันแปลงเวลาเป็นนาที
function timeToMinutes(time) {
  if (time instanceof Date) return time.getHours()*60 + time.getMinutes();
  if (typeof time !== 'string') return 0;
  const parts = time.includes(':') ? time.split(':') : time.split('.');
  return parseInt(parts[0],10)*60 + parseInt(parts[1],10);
}

// ฟังก์ชันแปลงวันที่เป็น string YYYY-MM-DD
function formatDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth()+1).padStart(2,'0');
  const d = String(date.getDate()).padStart(2,'0');
  return `${y}-${m}-${d}`;
}

// ตรวจสอบการจองซ้ำ
function checkDuplicateBooking(date, room, startTime, endTime) {
  const reservations = getReservationsByDateAndRoom(date, room);
  const newStart = timeToMinutes(startTime);
  const newEnd = timeToMinutes(endTime);

  for (const res of reservations) {
    const existingStart = timeToMinutes(res.startTime);
    const existingEnd = timeToMinutes(res.endTime);

    if ((newStart >= existingStart && newStart < existingEnd) ||
        (newEnd > existingStart && newEnd <= existingEnd) ||
        (newStart <= existingStart && newEnd >= existingEnd)) {
      return true;
    }
  }
  return false;
}


