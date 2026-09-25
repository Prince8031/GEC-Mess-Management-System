/**
 * mess — Smart Mess Management System
 * Developed by Prince kumar 
 * 
 * Production Backend Implementation
 */

const SHEETS = {
  USERS: 'Users',
  STUDENTS: 'Students',
  MANAGERS: 'Managers',
  HOSTELS: 'Hostels',
  ROOMS: 'Rooms',
  MEAL_RATES: 'MealRates',
  MENU: 'Menu',
  ATTENDANCE: 'Attendance',
  BILLS: 'Bills',
  PAYMENTS: 'Payments',
  INVENTORY: 'Inventory',
  COMPLAINTS: 'Complaints',
  NOTICES: 'Notices',
  AUDIT_LOGS: 'AuditLogs',
  SETTINGS: 'Settings'
};

const SCHEMAS = {
  [SHEETS.USERS]: ['UserId', 'Email', 'PasswordHash', 'Role', 'Status', 'CreatedAt', 'LastLogin'],
  [SHEETS.STUDENTS]: ['StudentId', 'Name', 'RollNumber', 'Email', 'Mobile', 'Course', 'Branch', 'Semester', 'Hostel', 'RoomNumber', 'JoiningDate', 'Status'],
  [SHEETS.MANAGERS]: ['ManagerId', 'Name', 'Email', 'Mobile', 'Role', 'Hostel', 'JoiningDate', 'Status'],
  [SHEETS.HOSTELS]: ['HostelId', 'HostelName', 'HostelType', 'Address', 'WardenName', 'Contact', 'Status'],
  [SHEETS.ROOMS]: ['RoomId', 'Hostel', 'RoomNumber', 'Capacity', 'Occupied', 'Available', 'Status'],
  [SHEETS.MEAL_RATES]: ['MealType', 'Rate', 'EffectiveDate', 'Status'],
  [SHEETS.MENU]: ['MenuId', 'Date', 'MealType', 'MenuItems', 'SpecialNote', 'Status'],
  [SHEETS.ATTENDANCE]: ['AttendanceId', 'Date', 'StudentId', 'StudentName', 'MealType', 'Status', 'MarkedBy', 'Timestamp'],
  [SHEETS.BILLS]: ['BillId', 'StudentId', 'StudentName', 'Month', 'Year', 'BreakfastCount', 'LunchCount', 'SnacksCount', 'DinnerCount', 'BreakfastRate', 'LunchRate', 'SnacksRate', 'DinnerRate', 'GrossAmount', 'Adjustment', 'FinalAmount', 'PaidAmount', 'DueAmount', 'Status', 'GeneratedDate'],
  [SHEETS.PAYMENTS]: ['PaymentId', 'BillId', 'StudentId', 'Date', 'Amount', 'PaymentMode', 'ReferenceNumber', 'CollectedBy', 'Remarks'],
  [SHEETS.INVENTORY]: ['ItemId', 'ItemName', 'Category', 'Unit', 'OpeningStock', 'PurchasedQuantity', 'UsedQuantity', 'CurrentStock', 'MinimumStock', 'Supplier', 'LastUpdated', 'Status'],
  [SHEETS.COMPLAINTS]: ['ComplaintId', 'StudentId', 'StudentName', 'Category', 'Subject', 'Description', 'Date', 'Status', 'Response', 'ResolvedBy', 'ResolvedDate'],
  [SHEETS.NOTICES]: ['NoticeId', 'Title', 'Message', 'Audience', 'Priority', 'PublishedDate', 'PublishedBy', 'Status'],
  [SHEETS.AUDIT_LOGS]: ['LogId', 'Timestamp', 'UserId', 'UserName', 'Role', 'Action', 'Module', 'RecordId', 'Description'],
  [SHEETS.SETTINGS]: ['Key', 'Value']
};

function doGet(e) {
  initializeDatabase();
  const output = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('mess — Smart Mess Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return output;
}

function getDb() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Unable to bind active spreadsheet. Please bind script to Google Sheet.');
  }
  return ss;
}

function initializeDatabase() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = getDb();
    
    Object.keys(SHEETS).forEach(sheetKey => {
      const sheetName = SHEETS[sheetKey];
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.appendRow(SCHEMAS[sheetName]);
        sheet.setFrozenRows(1);
      } else if (sheet.getLastRow() === 0) {
        sheet.appendRow(SCHEMAS[sheetName]);
        sheet.setFrozenRows(1);
      }
    });

    const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
    const existingSettings = readRows(settingsSheet);
    if (existingSettings.length === 0) {
      const defaults = [
        ['MESS_NAME', 'Smart Mess'],
        ['INSTITUTION_NAME', 'Central University Campus'],
        ['ADDRESS', 'Campus Dining Block, North Wing'],
        ['CONTACT', '+91 9876543210'],
        ['EMAIL', 'support@mess.local'],
        ['CURRENCY', '₹'],
        ['LOW_STOCK_THRESHOLD', '10'],
        ['APPLICATION_STATUS', 'ACTIVE']
      ];
      settingsSheet.getRange(2, 1, defaults.length, 2).setValues(defaults);
    }

    const mealRateSheet = ss.getSheetByName(SHEETS.MEAL_RATES);
    const existingRates = readRows(mealRateSheet);
    if (existingRates.length === 0) {
      const defaultRates = [
        ['BREAKFAST', 40, formatDate(new Date()), 'ACTIVE'],
        ['LUNCH', 80, formatDate(new Date()), 'ACTIVE'],
        ['SNACKS', 30, formatDate(new Date()), 'ACTIVE'],
        ['DINNER', 80, formatDate(new Date()), 'ACTIVE']
      ];
      mealRateSheet.getRange(2, 1, defaultRates.length, 4).setValues(defaultRates);
    }

    const usersSheet = ss.getSheetByName(SHEETS.USERS);
    const users = readRows(usersSheet);
    const hasAdmin = users.some(u => u.Role === 'ADMIN');
    if (!hasAdmin) {
      const adminEmail = 'admin@mess.local';
      const adminPasswordHash = hashPassword('ChangeMe@123');
      usersSheet.appendRow([
        'USR-' + Utilities.getUuid().substring(0, 8),
        adminEmail,
        adminPasswordHash,
        'ADMIN',
        'ACTIVE',
        formatTimestamp(new Date()),
        ''
      ]);
    }
    
    return { success: true, message: 'Database initialized successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ======================== AUTHENTICATION & SESSION ========================

function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return digest.map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0')).join('');
}

function loginUser(email, password) {
  try {
    if (!email || !password) return { success: false, message: 'Email and password are required.' };
    const ss = getDb();
    const sheet = ss.getSheetByName(SHEETS.USERS);
    const users = readRows(sheet);
    const user = users.find(u => u.Email.toLowerCase() === email.trim().toLowerCase());

    if (!user) return { success: false, message: 'Invalid credentials.' };
    if (user.Status !== 'ACTIVE') return { success: false, message: 'Account is inactive. Contact Administrator.' };

    const hashedInput = hashPassword(password);
    if (user.PasswordHash !== hashedInput) {
      return { success: false, message: 'Invalid credentials.' };
    }

    const token = Utilities.getUuid();
    const sessionData = {
      userId: user.UserId,
      email: user.Email,
      role: user.Role,
      token: token,
      expiresAt: new Date().getTime() + (12 * 60 * 60 * 1000)
    };

    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty('SESSION_DATA', JSON.stringify(sessionData));

    const rowIndex = users.findIndex(u => u.UserId === user.UserId) + 2;
    sheet.getRange(rowIndex, 7).setValue(formatTimestamp(new Date()));

    logAuditAction(user.UserId, user.Email, user.Role, 'LOGIN', 'USERS', user.UserId, 'User logged in successfully');

    let studentProfile = null;
    if (user.Role === 'STUDENT') {
      const studentSheet = ss.getSheetByName(SHEETS.STUDENTS);
      const students = readRows(studentSheet);
      studentProfile = students.find(s => s.Email.toLowerCase() === user.Email.toLowerCase()) || null;
    }

    return {
      success: true,
      message: 'Login successful',
      session: {
        token: token,
        userId: user.UserId,
        email: user.Email,
        role: user.Role,
        studentProfile: studentProfile
      }
    };
  } catch (err) {
    return { success: false, message: 'Authentication error: ' + err.toString() };
  }
}

function logoutUser(token) {
  try {
    const session = getValidatedSession(token);
    if (session) {
      logAuditAction(session.userId, session.email, session.role, 'LOGOUT', 'USERS', session.userId, 'User logged out');
    }
    PropertiesService.getUserProperties().deleteProperty('SESSION_DATA');
    return { success: true, message: 'Logged out successfully.' };
  } catch (err) {
    return { success: true };
  }
}

function validateSession(token) {
  const session = getValidatedSession(token);
  if (!session) return { success: false, message: 'Invalid or expired session.' };
  return { success: true, session: session };
}

function getValidatedSession(token) {
  if (!token) return null;
  const raw = PropertiesService.getUserProperties().getProperty('SESSION_DATA');
  if (!raw) return null;
  try {
    const session = JSON.parse(raw);
    if (session.token !== token) return null;
    if (new Date().getTime() > session.expiresAt) return null;
    return session;
  } catch (e) {
    return null;
  }
}

function requireAuth(token, allowedRoles) {
  const session = getValidatedSession(token);
  if (!session) throw new Error('Unauthorized: Session invalid or expired.');
  if (allowedRoles && allowedRoles.length > 0 && !allowedRoles.includes(session.role)) {
    throw new Error('Forbidden: Insufficient privileges.');
  }
  return session;
}

// ======================== DASHBOARD APIs ========================

function getDashboardData(token) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER', 'STUDENT']);
    const ss = getDb();

    if (session.role === 'STUDENT') {
      const studentSheet = ss.getSheetByName(SHEETS.STUDENTS);
      const student = readRows(studentSheet).find(s => s.Email.toLowerCase() === session.email.toLowerCase());
      if (!student) {
        return { success: false, message: 'Student record profile missing.' };
      }

      const todayStr = formatDate(new Date());
      const menu = readRows(ss.getSheetByName(SHEETS.MENU)).filter(m => m.Date === todayStr && m.Status === 'ACTIVE');
      const attendance = readRows(ss.getSheetByName(SHEETS.ATTENDANCE)).filter(a => a.StudentId === student.StudentId);
      const bills = readRows(ss.getSheetByName(SHEETS.BILLS)).filter(b => b.StudentId === student.StudentId);
      const complaints = readRows(ss.getSheetByName(SHEETS.COMPLAINTS)).filter(c => c.StudentId === student.StudentId);
      const notices = readRows(ss.getSheetByName(SHEETS.NOTICES)).filter(n => n.Status === 'ACTIVE' && (n.Audience === 'ALL' || n.Audience === 'STUDENTS'));

      const totalAttendance = attendance.length;
      const presentCount = attendance.filter(a => a.Status === 'PRESENT').length;
      const attendanceRate = totalAttendance > 0 ? Math.round((presentCount / totalAttendance) * 100) : 0;

      const currentMonth = new Date().toLocaleString('default', { month: 'long' });
      const currentYear = new Date().getFullYear().toString();
      const currentBill = bills.find(b => b.Month === currentMonth && b.Year === currentYear) || null;

      let totalDue = 0;
      let totalPaid = 0;
      bills.forEach(b => {
        totalDue += Number(b.DueAmount || 0);
        totalPaid += Number(b.PaidAmount || 0);
      });

      return {
        success: true,
        data: {
          role: 'STUDENT',
          student: student,
          todayMenu: menu,
          attendanceRate: attendanceRate,
          totalPresent: presentCount,
          totalMarkedMeals: totalAttendance,
          currentBill: currentBill,
          totalDue: totalDue,
          totalPaid: totalPaid,
          recentNotices: notices.slice(-5).reverse(),
          recentComplaints: complaints.slice(-5).reverse()
        }
      };
    }

    const students = readRows(ss.getSheetByName(SHEETS.STUDENTS));
    const managers = readRows(ss.getSheetByName(SHEETS.MANAGERS));
    const attendance = readRows(ss.getSheetByName(SHEETS.ATTENDANCE));
    const bills = readRows(ss.getSheetByName(SHEETS.BILLS));
    const payments = readRows(ss.getSheetByName(SHEETS.PAYMENTS));
    const inventory = readRows(ss.getSheetByName(SHEETS.INVENTORY));
    const complaints = readRows(ss.getSheetByName(SHEETS.COMPLAINTS));
    const notices = readRows(ss.getSheetByName(SHEETS.NOTICES));

    const todayStr = formatDate(new Date());
    const todayAttendance = attendance.filter(a => a.Date === todayStr);
    const todayPresent = todayAttendance.filter(a => a.Status === 'PRESENT').length;

    let pendingBillsCount = 0;
    let totalBilled = 0;
    let totalCollected = 0;
    let totalOutstanding = 0;

    bills.forEach(b => {
      const finalAmt = Number(b.FinalAmount || 0);
      const paidAmt = Number(b.PaidAmount || 0);
      const dueAmt = Number(b.DueAmount || 0);
      totalBilled += finalAmt;
      totalCollected += paidAmt;
      totalOutstanding += dueAmt;
      if (b.Status === 'PENDING' || b.Status === 'PARTIAL') pendingBillsCount++;
    });

    const lowStockItems = inventory.filter(i => Number(i.CurrentStock) <= Number(i.MinimumStock) && Number(i.CurrentStock) > 0);
    const outOfStockItems = inventory.filter(i => Number(i.CurrentStock) <= 0);
    const openComplaints = complaints.filter(c => c.Status === 'OPEN' || c.Status === 'IN PROGRESS');

    const mealCount = { BREAKFAST: 0, LUNCH: 0, SNACKS: 0, DINNER: 0 };
    todayAttendance.forEach(a => {
      if (a.Status === 'PRESENT' && mealCount[a.MealType] !== undefined) {
        mealCount[a.MealType]++;
      }
    });

    return {
      success: true,
      data: {
        role: session.role,
        stats: {
          totalStudents: students.length,
          totalManagers: managers.length,
          todayAttendance: todayPresent,
          todayTotalAttendanceRecords: todayAttendance.length,
          pendingBills: pendingBillsCount,
          totalBilled: totalBilled,
          totalCollected: totalCollected,
          totalOutstanding: totalOutstanding,
          lowStock: lowStockItems.length,
          outOfStock: outOfStockItems.length,
          openComplaints: openComplaints.length,
          todayMeals: mealCount
        },
        inventoryStatus: {
          inStock: inventory.filter(i => Number(i.CurrentStock) > Number(i.MinimumStock)).length,
          lowStock: lowStockItems.length,
          outOfStock: outOfStockItems.length
        },
        billingStatus: {
          paid: bills.filter(b => b.Status === 'PAID').length,
          pending: bills.filter(b => b.Status === 'PENDING').length,
          partial: bills.filter(b => b.Status === 'PARTIAL').length
        },
        recentNotices: notices.slice(-5).reverse(),
        recentComplaints: complaints.slice(-5).reverse()
      }
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ======================== STUDENT MANAGEMENT ========================

function getStudents(token) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER']);
    const sheet = getDb().getSheetByName(SHEETS.STUDENTS);
    return { success: true, data: readRows(sheet) };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function addStudent(token, studentData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const session = requireAuth(token, ['ADMIN']);
    validateRequired(studentData, ['Name', 'Email', 'RollNumber', 'Mobile', 'Hostel', 'RoomNumber']);

    const ss = getDb();
    const sheet = ss.getSheetByName(SHEETS.STUDENTS);
    const students = readRows(sheet);

    if (students.some(s => s.Email.toLowerCase() === studentData.Email.trim().toLowerCase())) {
      return { success: false, message: 'Student with this Email already exists.' };
    }
    if (students.some(s => s.RollNumber.toLowerCase() === studentData.RollNumber.trim().toLowerCase())) {
      return { success: false, message: 'Student with this Roll Number already exists.' };
    }

    const roomSheet = ss.getSheetByName(SHEETS.ROOMS);
    const rooms = readRows(roomSheet);
    const targetRoomIndex = rooms.findIndex(r => r.Hostel === studentData.Hostel && r.RoomNumber === studentData.RoomNumber);
    if (targetRoomIndex !== -1) {
      const room = rooms[targetRoomIndex];
      if (Number(room.Occupied) >= Number(room.Capacity)) {
        return { success: false, message: 'Selected room is already at full capacity.' };
      }
      roomSheet.getRange(targetRoomIndex + 2, 5).setValue(Number(room.Occupied) + 1);
      roomSheet.getRange(targetRoomIndex + 2, 6).setValue(Number(room.Capacity) - (Number(room.Occupied) + 1));
    }

    const studentId = 'STU-' + Utilities.getUuid().substring(0, 6).toUpperCase();
    const newStudent = [
      studentId,
      studentData.Name.trim(),
      studentData.RollNumber.trim(),
      studentData.Email.trim().toLowerCase(),
      studentData.Mobile.trim(),
      studentData.Course || '',
      studentData.Branch || '',
      studentData.Semester || '',
      studentData.Hostel,
      studentData.RoomNumber,
      studentData.JoiningDate || formatDate(new Date()),
      studentData.Status || 'ACTIVE'
    ];
    sheet.appendRow(newStudent);

    const usersSheet = ss.getSheetByName(SHEETS.USERS);
    const users = readRows(usersSheet);
    if (!users.some(u => u.Email.toLowerCase() === studentData.Email.trim().toLowerCase())) {
      usersSheet.appendRow([
        'USR-' + Utilities.getUuid().substring(0, 8),
        studentData.Email.trim().toLowerCase(),
        hashPassword(studentData.Password || 'Student@123'),
        'STUDENT',
        'ACTIVE',
        formatTimestamp(new Date()),
        ''
      ]);
    }

    logAuditAction(session.userId, session.email, session.role, 'CREATE', 'STUDENTS', studentId, `Added student ${studentData.Name}`);
    return { success: true, message: 'Student added successfully.', studentId: studentId };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateStudent(token, studentId, studentData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const session = requireAuth(token, ['ADMIN']);
    const ss = getDb();
    const sheet = ss.getSheetByName(SHEETS.STUDENTS);
    const students = readRows(sheet);
    const index = students.findIndex(s => s.StudentId === studentId);
    if (index === -1) return { success: false, message: 'Student not found.' };

    const oldStudent = students[index];
    if (studentData.Hostel !== oldStudent.Hostel || studentData.RoomNumber !== oldStudent.RoomNumber) {
      updateRoomOccupancy(ss, oldStudent.Hostel, oldStudent.RoomNumber, -1);
      updateRoomOccupancy(ss, studentData.Hostel, studentData.RoomNumber, 1);
    }

    const row = index + 2;
    sheet.getRange(row, 2).setValue(studentData.Name);
    sheet.getRange(row, 3).setValue(studentData.RollNumber);
    sheet.getRange(row, 4).setValue(studentData.Email.toLowerCase());
    sheet.getRange(row, 5).setValue(studentData.Mobile);
    sheet.getRange(row, 6).setValue(studentData.Course);
    sheet.getRange(row, 7).setValue(studentData.Branch);
    sheet.getRange(row, 8).setValue(studentData.Semester);
    sheet.getRange(row, 9).setValue(studentData.Hostel);
    sheet.getRange(row, 10).setValue(studentData.RoomNumber);
    sheet.getRange(row, 12).setValue(studentData.Status);

    logAuditAction(session.userId, session.email, session.role, 'UPDATE', 'STUDENTS', studentId, `Updated student ${studentData.Name}`);
    return { success: true, message: 'Student updated successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteStudent(token, studentId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const session = requireAuth(token, ['ADMIN']);
    const ss = getDb();
    const sheet = ss.getSheetByName(SHEETS.STUDENTS);
    const students = readRows(sheet);
    const index = students.findIndex(s => s.StudentId === studentId);
    if (index === -1) return { success: false, message: 'Student not found.' };

    const student = students[index];
    updateRoomOccupancy(ss, student.Hostel, student.RoomNumber, -1);
    sheet.deleteRow(index + 2);

    const usersSheet = ss.getSheetByName(SHEETS.USERS);
    const users = readRows(usersSheet);
    const uIndex = users.findIndex(u => u.Email.toLowerCase() === student.Email.toLowerCase());
    if (uIndex !== -1) {
      usersSheet.deleteRow(uIndex + 2);
    }

    logAuditAction(session.userId, session.email, session.role, 'DELETE', 'STUDENTS', studentId, `Deleted student ${student.Name}`);
    return { success: true, message: 'Student deleted successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ======================== MANAGER MANAGEMENT ========================

function getManagers(token) {
  try {
    requireAuth(token, ['ADMIN']);
    return { success: true, data: readRows(getDb().getSheetByName(SHEETS.MANAGERS)) };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function addManager(token, data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const session = requireAuth(token, ['ADMIN']);
    validateRequired(data, ['Name', 'Email', 'Mobile']);

    const ss = getDb();
    const sheet = ss.getSheetByName(SHEETS.MANAGERS);
    const managers = readRows(sheet);
    if (managers.some(m => m.Email.toLowerCase() === data.Email.trim().toLowerCase())) {
      return { success: false, message: 'Manager with this Email already exists.' };
    }

    const managerId = 'MGR-' + Utilities.getUuid().substring(0, 6).toUpperCase();
    sheet.appendRow([
      managerId,
      data.Name.trim(),
      data.Email.trim().toLowerCase(),
      data.Mobile.trim(),
      data.Role || 'Mess Manager',
      data.Hostel || 'All Hostels',
      data.JoiningDate || formatDate(new Date()),
      data.Status || 'ACTIVE'
    ]);

    const usersSheet = ss.getSheetByName(SHEETS.USERS);
    usersSheet.appendRow([
      'USR-' + Utilities.getUuid().substring(0, 8),
      data.Email.trim().toLowerCase(),
      hashPassword(data.Password || 'Manager@123'),
      'MANAGER',
      'ACTIVE',
      formatTimestamp(new Date()),
      ''
    ]);

    logAuditAction(session.userId, session.email, session.role, 'CREATE', 'MANAGERS', managerId, `Created manager ${data.Name}`);
    return { success: true, message: 'Manager added successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateManager(token, managerId, data) {
  try {
    const session = requireAuth(token, ['ADMIN']);
    const sheet = getDb().getSheetByName(SHEETS.MANAGERS);
    const managers = readRows(sheet);
    const idx = managers.findIndex(m => m.ManagerId === managerId);
    if (idx === -1) return { success: false, message: 'Manager not found.' };

    const row = idx + 2;
    sheet.getRange(row, 2).setValue(data.Name);
    sheet.getRange(row, 4).setValue(data.Mobile);
    sheet.getRange(row, 5).setValue(data.Role);
    sheet.getRange(row, 6).setValue(data.Hostel);
    sheet.getRange(row, 8).setValue(data.Status);

    logAuditAction(session.userId, session.email, session.role, 'UPDATE', 'MANAGERS', managerId, `Updated manager ${data.Name}`);
    return { success: true, message: 'Manager updated successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteManager(token, managerId) {
  try {
    const session = requireAuth(token, ['ADMIN']);
    const ss = getDb();
    const sheet = ss.getSheetByName(SHEETS.MANAGERS);
    const managers = readRows(sheet);
    const idx = managers.findIndex(m => m.ManagerId === managerId);
    if (idx === -1) return { success: false, message: 'Manager not found.' };

    const mgr = managers[idx];
    sheet.deleteRow(idx + 2);

    const usersSheet = ss.getSheetByName(SHEETS.USERS);
    const users = readRows(usersSheet);
    const uIdx = users.findIndex(u => u.Email.toLowerCase() === mgr.Email.toLowerCase());
    if (uIdx !== -1) usersSheet.deleteRow(uIdx + 2);

    logAuditAction(session.userId, session.email, session.role, 'DELETE', 'MANAGERS', managerId, `Deleted manager ${mgr.Name}`);
    return { success: true, message: 'Manager deleted successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ======================== HOSTEL & ROOM MANAGEMENT ========================

function getHostels(token) {
  try {
    requireAuth(token, ['ADMIN', 'MANAGER', 'STUDENT']);
    return { success: true, data: readRows(getDb().getSheetByName(SHEETS.HOSTELS)) };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function addHostel(token, data) {
  try {
    const session = requireAuth(token, ['ADMIN']);
    validateRequired(data, ['HostelName', 'HostelType']);
    const sheet = getDb().getSheetByName(SHEETS.HOSTELS);
    const id = 'HST-' + Utilities.getUuid().substring(0, 6).toUpperCase();
    sheet.appendRow([
      id,
      data.HostelName.trim(),
      data.HostelType,
      data.Address || '',
      data.WardenName || '',
      data.Contact || '',
      data.Status || 'ACTIVE'
    ]);
    logAuditAction(session.userId, session.email, session.role, 'CREATE', 'HOSTELS', id, `Added hostel ${data.HostelName}`);
    return { success: true, message: 'Hostel added successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteHostel(token, hostelId) {
  try {
    const session = requireAuth(token, ['ADMIN']);
    const sheet = getDb().getSheetByName(SHEETS.HOSTELS);
    const hostels = readRows(sheet);
    const idx = hostels.findIndex(h => h.HostelId === hostelId);
    if (idx === -1) return { success: false, message: 'Hostel not found.' };
    sheet.deleteRow(idx + 2);
    logAuditAction(session.userId, session.email, session.role, 'DELETE', 'HOSTELS', hostelId, 'Deleted hostel');
    return { success: true, message: 'Hostel deleted successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getRooms(token) {
  try {
    requireAuth(token, ['ADMIN', 'MANAGER', 'STUDENT']);
    return { success: true, data: readRows(getDb().getSheetByName(SHEETS.ROOMS)) };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function addRoom(token, data) {
  try {
    const session = requireAuth(token, ['ADMIN']);
    validateRequired(data, ['Hostel', 'RoomNumber', 'Capacity']);
    const sheet = getDb().getSheetByName(SHEETS.ROOMS);
    const rooms = readRows(sheet);
    if (rooms.some(r => r.Hostel === data.Hostel && r.RoomNumber === data.RoomNumber)) {
      return { success: false, message: 'Room number already exists in this hostel.' };
    }
    const id = 'RM-' + Utilities.getUuid().substring(0, 6).toUpperCase();
    const capacity = Number(data.Capacity);
    sheet.appendRow([
      id,
      data.Hostel,
      data.RoomNumber,
      capacity,
      0,
      capacity,
      'AVAILABLE'
    ]);
    logAuditAction(session.userId, session.email, session.role, 'CREATE', 'ROOMS', id, `Added room ${data.RoomNumber} in ${data.Hostel}`);
    return { success: true, message: 'Room created successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteRoom(token, roomId) {
  try {
    const session = requireAuth(token, ['ADMIN']);
    const sheet = getDb().getSheetByName(SHEETS.ROOMS);
    const rooms = readRows(sheet);
    const idx = rooms.findIndex(r => r.RoomId === roomId);
    if (idx === -1) return { success: false, message: 'Room not found.' };
    if (Number(rooms[idx].Occupied) > 0) return { success: false, message: 'Cannot delete occupied room.' };
    sheet.deleteRow(idx + 2);
    logAuditAction(session.userId, session.email, session.role, 'DELETE', 'ROOMS', roomId, 'Deleted room');
    return { success: true, message: 'Room deleted successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateRoomOccupancy(ss, hostelName, roomNumber, delta) {
  const roomSheet = ss.getSheetByName(SHEETS.ROOMS);
  const rooms = readRows(roomSheet);
  const idx = rooms.findIndex(r => r.Hostel === hostelName && r.RoomNumber === roomNumber);
  if (idx !== -1) {
    const row = idx + 2;
    const capacity = Number(rooms[idx].Capacity || 0);
    const currentOcc = Number(rooms[idx].Occupied || 0);
    const newOcc = Math.max(0, Math.min(capacity, currentOcc + delta));
    const newAvail = capacity - newOcc;
    roomSheet.getRange(row, 5).setValue(newOcc);
    roomSheet.getRange(row, 6).setValue(newAvail);
    roomSheet.getRange(row, 7).setValue(newAvail === 0 ? 'FULL' : 'AVAILABLE');
  }
}

// ======================== MEAL RATES & MENU ========================

function getMealRates(token) {
  try {
    requireAuth(token, ['ADMIN', 'MANAGER']);
    return { success: true, data: readRows(getDb().getSheetByName(SHEETS.MEAL_RATES)) };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function saveMealRate(token, mealType, rate) {
  try {
    const session = requireAuth(token, ['ADMIN']);
    const sheet = getDb().getSheetByName(SHEETS.MEAL_RATES);
    const rates = readRows(sheet);
    const idx = rates.findIndex(r => r.MealType.toUpperCase() === mealType.toUpperCase());
    const numericRate = Number(rate);
    if (isNaN(numericRate) || numericRate < 0) return { success: false, message: 'Invalid rate.' };

    if (idx !== -1) {
      sheet.getRange(idx + 2, 2).setValue(numericRate);
      sheet.getRange(idx + 2, 3).setValue(formatDate(new Date()));
    } else {
      sheet.appendRow([mealType.toUpperCase(), numericRate, formatDate(new Date()), 'ACTIVE']);
    }

    logAuditAction(session.userId, session.email, session.role, 'UPDATE', 'MEAL_RATES', mealType, `Updated rate for ${mealType} to ${numericRate}`);
    return { success: true, message: 'Meal rate saved successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getMenu(token) {
  try {
    requireAuth(token, ['ADMIN', 'MANAGER', 'STUDENT']);
    return { success: true, data: readRows(getDb().getSheetByName(SHEETS.MENU)) };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function addMenu(token, data) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER']);
    validateRequired(data, ['Date', 'MealType', 'MenuItems']);
    const sheet = getDb().getSheetByName(SHEETS.MENU);
    const id = 'MNU-' + Utilities.getUuid().substring(0, 6).toUpperCase();
    sheet.appendRow([
      id,
      data.Date,
      data.MealType.toUpperCase(),
      data.MenuItems,
      data.SpecialNote || '',
      data.Status || 'ACTIVE'
    ]);
    logAuditAction(session.userId, session.email, session.role, 'CREATE', 'MENU', id, `Created menu for ${data.Date} - ${data.MealType}`);
    return { success: true, message: 'Menu added successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteMenu(token, menuId) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER']);
    const sheet = getDb().getSheetByName(SHEETS.MENU);
    const rows = readRows(sheet);
    const idx = rows.findIndex(m => m.MenuId === menuId);
    if (idx === -1) return { success: false, message: 'Menu not found.' };
    sheet.deleteRow(idx + 2);
    logAuditAction(session.userId, session.email, session.role, 'DELETE', 'MENU', menuId, 'Deleted menu record');
    return { success: true, message: 'Menu deleted.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ======================== ATTENDANCE MANAGEMENT ========================

function getAttendance(token, filterDate, filterMealType) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER', 'STUDENT']);
    const records = readRows(getDb().getSheetByName(SHEETS.ATTENDANCE));

    if (session.role === 'STUDENT') {
      const studentSheet = getDb().getSheetByName(SHEETS.STUDENTS);
      const student = readRows(studentSheet).find(s => s.Email.toLowerCase() === session.email.toLowerCase());
      if (!student) return { success: true, data: [] };
      return { success: true, data: records.filter(r => r.StudentId === student.StudentId) };
    }

    let filtered = records;
    if (filterDate) filtered = filtered.filter(r => r.Date === filterDate);
    if (filterMealType) filtered = filtered.filter(r => r.MealType === filterMealType);
    return { success: true, data: filtered };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function markAttendance(token, records) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const session = requireAuth(token, ['ADMIN', 'MANAGER']);
    if (!Array.isArray(records) || records.length === 0) {
      return { success: false, message: 'No attendance records provided.' };
    }

    const sheet = getDb().getSheetByName(SHEETS.ATTENDANCE);
    const existing = readRows(sheet);
    const rowsToAppend = [];

    records.forEach(rec => {
      const existingIdx = existing.findIndex(e => e.Date === rec.Date && e.StudentId === rec.StudentId && e.MealType === rec.MealType);
      if (existingIdx !== -1) {
        const rowNum = existingIdx + 2;
        sheet.getRange(rowNum, 6).setValue(rec.Status);
        sheet.getRange(rowNum, 7).setValue(session.email);
        sheet.getRange(rowNum, 8).setValue(formatTimestamp(new Date()));
      } else {
        const id = 'ATT-' + Utilities.getUuid().substring(0, 6).toUpperCase();
        rowsToAppend.push([
          id,
          rec.Date,
          rec.StudentId,
          rec.StudentName,
          rec.MealType,
          rec.Status,
          session.email,
          formatTimestamp(new Date())
        ]);
      }
    });

    if (rowsToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }

    logAuditAction(session.userId, session.email, session.role, 'ATTENDANCE_MARKED', 'ATTENDANCE', '', `Batch attendance processed: ${records.length} items`);
    return { success: true, message: 'Attendance recorded successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ======================== MONTHLY BILLING ========================

function getBills(token, month, year) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER', 'STUDENT']);
    const allBills = readRows(getDb().getSheetByName(SHEETS.BILLS));

    if (session.role === 'STUDENT') {
      const studentSheet = getDb().getSheetByName(SHEETS.STUDENTS);
      const student = readRows(studentSheet).find(s => s.Email.toLowerCase() === session.email.toLowerCase());
      if (!student) return { success: true, data: [] };
      return { success: true, data: allBills.filter(b => b.StudentId === student.StudentId) };
    }

    let filtered = allBills;
    if (month) filtered = filtered.filter(b => b.Month === month);
    if (year) filtered = filtered.filter(b => b.Year === String(year));
    return { success: true, data: filtered };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function generateMonthlyBill(token, month, year, adjustmentNotes) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const session = requireAuth(token, ['ADMIN', 'MANAGER']);
    if (!month || !year) return { success: false, message: 'Month and year are required.' };

    const ss = getDb();
    const students = readRows(ss.getSheetByName(SHEETS.STUDENTS)).filter(s => s.Status === 'ACTIVE');
    const attendance = readRows(ss.getSheetByName(SHEETS.ATTENDANCE));
    const mealRates = readRows(ss.getSheetByName(SHEETS.MEAL_RATES));
    const billsSheet = ss.getSheetByName(SHEETS.BILLS);
    const existingBills = readRows(billsSheet);

    const rates = { BREAKFAST: 0, LUNCH: 0, SNACKS: 0, DINNER: 0 };
    mealRates.forEach(r => {
      rates[r.MealType.toUpperCase()] = Number(r.Rate) || 0;
    });

    const monthAttendance = attendance.filter(a => {
      const d = new Date(a.Date);
      const mStr = d.toLocaleString('default', { month: 'long' });
      const yStr = d.getFullYear().toString();
      return mStr.toLowerCase() === month.toLowerCase() && yStr === String(year) && a.Status === 'PRESENT';
    });

    let generatedCount = 0;

    students.forEach(student => {
      const studentAttendance = monthAttendance.filter(a => a.StudentId === student.StudentId);
      const bCount = studentAttendance.filter(a => a.MealType === 'BREAKFAST').length;
      const lCount = studentAttendance.filter(a => a.MealType === 'LUNCH').length;
      const sCount = studentAttendance.filter(a => a.MealType === 'SNACKS').length;
      const dCount = studentAttendance.filter(a => a.MealType === 'DINNER').length;

      const gross = (bCount * rates.BREAKFAST) + (lCount * rates.LUNCH) + (sCount * rates.SNACKS) + (dCount * rates.DINNER);
      const adjustment = 0;
      const finalAmount = gross + adjustment;

      const existingIndex = existingBills.findIndex(b => b.StudentId === student.StudentId && b.Month === month && b.Year === String(year));

      if (existingIndex !== -1) {
        const rowNum = existingIndex + 2;
        const currentPaid = Number(billsSheet.getRange(rowNum, 17).getValue() || 0);
        const dueAmount = Math.max(0, finalAmount - currentPaid);
        const billStatus = dueAmount === 0 && finalAmount > 0 ? 'PAID' : (currentPaid > 0 ? 'PARTIAL' : 'PENDING');

        billsSheet.getRange(rowNum, 6, 1, 14).setValues([[
          bCount, lCount, sCount, dCount,
          rates.BREAKFAST, rates.LUNCH, rates.SNACKS, rates.DINNER,
          gross, adjustment, finalAmount, currentPaid, dueAmount, billStatus
        ]]);
      } else {
        const billId = 'BIL-' + Utilities.getUuid().substring(0, 6).toUpperCase();
        const dueAmount = finalAmount;
        const billStatus = finalAmount === 0 ? 'PAID' : 'PENDING';

        billsSheet.appendRow([
          billId,
          student.StudentId,
          student.Name,
          month,
          String(year),
          bCount,
          lCount,
          sCount,
          dCount,
          rates.BREAKFAST,
          rates.LUNCH,
          rates.SNACKS,
          rates.DINNER,
          gross,
          adjustment,
          finalAmount,
          0,
          dueAmount,
          billStatus,
          formatDate(new Date())
        ]);
      }
      generatedCount++;
    });

    logAuditAction(session.userId, session.email, session.role, 'BILL_GENERATED', 'BILLS', `${month}-${year}`, `Generated bills for ${generatedCount} students`);
    return { success: true, message: `Successfully generated ${generatedCount} monthly bills for ${month} ${year}.` };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ======================== PAYMENT LEDGER ========================

function getPayments(token) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER', 'STUDENT']);
    const payments = readRows(getDb().getSheetByName(SHEETS.PAYMENTS));

    if (session.role === 'STUDENT') {
      const studentSheet = getDb().getSheetByName(SHEETS.STUDENTS);
      const student = readRows(studentSheet).find(s => s.Email.toLowerCase() === session.email.toLowerCase());
      if (!student) return { success: true, data: [] };
      return { success: true, data: payments.filter(p => p.StudentId === student.StudentId) };
    }

    return { success: true, data: payments };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function recordPayment(token, data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const session = requireAuth(token, ['ADMIN', 'MANAGER']);
    validateRequired(data, ['BillId', 'Amount', 'PaymentMode']);

    const amount = Number(data.Amount);
    if (isNaN(amount) || amount <= 0) return { success: false, message: 'Invalid payment amount.' };

    const ss = getDb();
    const billsSheet = ss.getSheetByName(SHEETS.BILLS);
    const bills = readRows(billsSheet);
    const bIndex = bills.findIndex(b => b.BillId === data.BillId);
    if (bIndex === -1) return { success: false, message: 'Selected Bill does not exist.' };

    const bill = bills[bIndex];
    const finalAmt = Number(bill.FinalAmount || 0);
    const currentPaid = Number(bill.PaidAmount || 0);
    const currentDue = Number(bill.DueAmount || 0);

    if (amount > currentDue) {
      return { success: false, message: `Payment exceeds current due amount of ₹${currentDue}.` };
    }

    const newPaid = currentPaid + amount;
    const newDue = Math.max(0, finalAmt - newPaid);
    const newStatus = newDue === 0 ? 'PAID' : 'PARTIAL';

    const billRowNum = bIndex + 2;
    billsSheet.getRange(billRowNum, 17).setValue(newPaid);
    billsSheet.getRange(billRowNum, 18).setValue(newDue);
    billsSheet.getRange(billRowNum, 19).setValue(newStatus);

    const paymentId = 'PAY-' + Utilities.getUuid().substring(0, 6).toUpperCase();
    const paySheet = ss.getSheetByName(SHEETS.PAYMENTS);
    paySheet.appendRow([
      paymentId,
      bill.BillId,
      bill.StudentId,
      data.Date || formatDate(new Date()),
      amount,
      data.PaymentMode,
      data.ReferenceNumber || 'N/A',
      session.email,
      data.Remarks || 'Offline Collection'
    ]);

    logAuditAction(session.userId, session.email, session.role, 'PAYMENT', 'PAYMENTS', paymentId, `Recorded payment ₹${amount} for bill ${bill.BillId}`);
    return { success: true, message: 'Payment recorded successfully.', paymentId: paymentId };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ======================== INVENTORY MANAGEMENT ========================

function getInventory(token) {
  try {
    requireAuth(token, ['ADMIN', 'MANAGER']);
    return { success: true, data: readRows(getDb().getSheetByName(SHEETS.INVENTORY)) };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function addInventoryItem(token, data) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER']);
    validateRequired(data, ['ItemName', 'Category', 'Unit', 'OpeningStock', 'MinimumStock']);

    const sheet = getDb().getSheetByName(SHEETS.INVENTORY);
    const openStock = Number(data.OpeningStock) || 0;
    const purchased = Number(data.PurchasedQuantity) || 0;
    const used = Number(data.UsedQuantity) || 0;
    const currentStock = openStock + purchased - used;
    const minStock = Number(data.MinimumStock) || 0;

    let status = 'NORMAL';
    if (currentStock <= 0) status = 'OUT OF STOCK';
    else if (currentStock <= minStock) status = 'LOW STOCK';

    const itemId = 'INV-' + Utilities.getUuid().substring(0, 6).toUpperCase();
    sheet.appendRow([
      itemId,
      data.ItemName.trim(),
      data.Category,
      data.Unit,
      openStock,
      purchased,
      used,
      currentStock,
      minStock,
      data.Supplier || '',
      formatDate(new Date()),
      status
    ]);

    logAuditAction(session.userId, session.email, session.role, 'CREATE', 'INVENTORY', itemId, `Added item ${data.ItemName}`);
    return { success: true, message: 'Inventory item added.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateInventoryItem(token, itemId, data) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER']);
    const sheet = getDb().getSheetByName(SHEETS.INVENTORY);
    const items = readRows(sheet);
    const idx = items.findIndex(i => i.ItemId === itemId);
    if (idx === -1) return { success: false, message: 'Item not found.' };

    const openStock = Number(data.OpeningStock) || 0;
    const purchased = Number(data.PurchasedQuantity) || 0;
    const used = Number(data.UsedQuantity) || 0;
    const currentStock = openStock + purchased - used;
    const minStock = Number(data.MinimumStock) || 0;

    let status = 'NORMAL';
    if (currentStock <= 0) status = 'OUT OF STOCK';
    else if (currentStock <= minStock) status = 'LOW STOCK';

    const row = idx + 2;
    sheet.getRange(row, 2).setValue(data.ItemName);
    sheet.getRange(row, 3).setValue(data.Category);
    sheet.getRange(row, 4).setValue(data.Unit);
    sheet.getRange(row, 5).setValue(openStock);
    sheet.getRange(row, 6).setValue(purchased);
    sheet.getRange(row, 7).setValue(used);
    sheet.getRange(row, 8).setValue(currentStock);
    sheet.getRange(row, 9).setValue(minStock);
    sheet.getRange(row, 10).setValue(data.Supplier);
    sheet.getRange(row, 11).setValue(formatDate(new Date()));
    sheet.getRange(row, 12).setValue(status);

    logAuditAction(session.userId, session.email, session.role, 'UPDATE', 'INVENTORY', itemId, `Updated item ${data.ItemName}`);
    return { success: true, message: 'Inventory item updated.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteInventoryItem(token, itemId) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER']);
    const sheet = getDb().getSheetByName(SHEETS.INVENTORY);
    const items = readRows(sheet);
    const idx = items.findIndex(i => i.ItemId === itemId);
    if (idx === -1) return { success: false, message: 'Item not found.' };
    sheet.deleteRow(idx + 2);
    logAuditAction(session.userId, session.email, session.role, 'DELETE', 'INVENTORY', itemId, 'Deleted inventory item');
    return { success: true, message: 'Inventory item deleted.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ======================== COMPLAINTS ========================

function getComplaints(token) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER', 'STUDENT']);
    const complaints = readRows(getDb().getSheetByName(SHEETS.COMPLAINTS));

    if (session.role === 'STUDENT') {
      const studentSheet = getDb().getSheetByName(SHEETS.STUDENTS);
      const student = readRows(studentSheet).find(s => s.Email.toLowerCase() === session.email.toLowerCase());
      if (!student) return { success: true, data: [] };
      return { success: true, data: complaints.filter(c => c.StudentId === student.StudentId) };
    }

    return { success: true, data: complaints };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function createComplaint(token, data) {
  try {
    const session = requireAuth(token, ['STUDENT', 'ADMIN']);
    validateRequired(data, ['Category', 'Subject', 'Description']);

    const ss = getDb();
    let studentId = data.StudentId;
    let studentName = data.StudentName;

    if (session.role === 'STUDENT') {
      const studentSheet = ss.getSheetByName(SHEETS.STUDENTS);
      const student = readRows(studentSheet).find(s => s.Email.toLowerCase() === session.email.toLowerCase());
      if (!student) return { success: false, message: 'Student record profile missing.' };
      studentId = student.StudentId;
      studentName = student.Name;
    }

    const complaintId = 'CMP-' + Utilities.getUuid().substring(0, 6).toUpperCase();
    const sheet = ss.getSheetByName(SHEETS.COMPLAINTS);
    sheet.appendRow([
      complaintId,
      studentId,
      studentName,
      data.Category,
      data.Subject,
      data.Description,
      formatDate(new Date()),
      'OPEN',
      '',
      '',
      ''
    ]);

    logAuditAction(session.userId, session.email, session.role, 'CREATE', 'COMPLAINTS', complaintId, `Logged complaint: ${data.Subject}`);
    return { success: true, message: 'Complaint registered successfully.', complaintId: complaintId };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function resolveComplaint(token, complaintId, responseText, status) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER']);
    const sheet = getDb().getSheetByName(SHEETS.COMPLAINTS);
    const complaints = readRows(sheet);
    const idx = complaints.findIndex(c => c.ComplaintId === complaintId);
    if (idx === -1) return { success: false, message: 'Complaint not found.' };

    const row = idx + 2;
    sheet.getRange(row, 8).setValue(status || 'RESOLVED');
    sheet.getRange(row, 9).setValue(responseText || 'Action taken.');
    sheet.getRange(row, 10).setValue(session.email);
    sheet.getRange(row, 11).setValue(formatDate(new Date()));

    logAuditAction(session.userId, session.email, session.role, 'UPDATE', 'COMPLAINTS', complaintId, `Complaint updated to ${status}`);
    return { success: true, message: 'Complaint updated successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ======================== NOTICES ========================

function getNotices(token) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER', 'STUDENT']);
    const notices = readRows(getDb().getSheetByName(SHEETS.NOTICES));

    if (session.role === 'STUDENT') {
      return { success: true, data: notices.filter(n => n.Status === 'ACTIVE' && (n.Audience === 'ALL' || n.Audience === 'STUDENTS')) };
    }
    return { success: true, data: notices };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function createNotice(token, data) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER']);
    validateRequired(data, ['Title', 'Message', 'Audience', 'Priority']);

    const sheet = getDb().getSheetByName(SHEETS.NOTICES);
    const noticeId = 'NOT-' + Utilities.getUuid().substring(0, 6).toUpperCase();
    sheet.appendRow([
      noticeId,
      data.Title,
      data.Message,
      data.Audience,
      data.Priority,
      formatDate(new Date()),
      session.email,
      'ACTIVE'
    ]);

    logAuditAction(session.userId, session.email, session.role, 'CREATE', 'NOTICES', noticeId, `Published notice ${data.Title}`);
    return { success: true, message: 'Notice published successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteNotice(token, noticeId) {
  try {
    const session = requireAuth(token, ['ADMIN', 'MANAGER']);
    const sheet = getDb().getSheetByName(SHEETS.NOTICES);
    const notices = readRows(sheet);
    const idx = notices.findIndex(n => n.NoticeId === noticeId);
    if (idx === -1) return { success: false, message: 'Notice not found.' };

    sheet.deleteRow(idx + 2);
    logAuditAction(session.userId, session.email, session.role, 'DELETE', 'NOTICES', noticeId, 'Deleted notice');
    return { success: true, message: 'Notice deleted successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ======================== REPORTS ========================

function getReports(token, type, filters) {
  try {
    requireAuth(token, ['ADMIN', 'MANAGER']);
    const ss = getDb();
    filters = filters || {};

    if (type === 'BILLING') {
      let bills = readRows(ss.getSheetByName(SHEETS.BILLS));
      if (filters.month) bills = bills.filter(b => b.Month === filters.month);
      if (filters.year) bills = bills.filter(b => b.Year === String(filters.year));
      return { success: true, data: bills };
    }

    if (type === 'PAYMENTS') {
      let payments = readRows(ss.getSheetByName(SHEETS.PAYMENTS));
      if (filters.startDate) payments = payments.filter(p => p.Date >= filters.startDate);
      if (filters.endDate) payments = payments.filter(p => p.Date <= filters.endDate);
      return { success: true, data: payments };
    }

    if (type === 'ATTENDANCE') {
      let att = readRows(ss.getSheetByName(SHEETS.ATTENDANCE));
      if (filters.startDate) att = att.filter(a => a.Date >= filters.startDate);
      if (filters.endDate) att = att.filter(a => a.Date <= filters.endDate);
      return { success: true, data: att };
    }

    if (type === 'INVENTORY') {
      return { success: true, data: readRows(ss.getSheetByName(SHEETS.INVENTORY)) };
    }

    if (type === 'COMPLAINTS') {
      return { success: true, data: readRows(ss.getSheetByName(SHEETS.COMPLAINTS)) };
    }

    return { success: false, message: 'Invalid report type requested.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ======================== SETTINGS & AUDIT LOGS ========================

function getSettings(token) {
  try {
    requireAuth(token, ['ADMIN']);
    const rows = readRows(getDb().getSheetByName(SHEETS.SETTINGS));
    const obj = {};
    rows.forEach(r => { obj[r.Key] = r.Value; });
    return { success: true, data: obj };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function saveSettings(token, settingsObj) {
  try {
    const session = requireAuth(token, ['ADMIN']);
    const sheet = getDb().getSheetByName(SHEETS.SETTINGS);
    const current = readRows(sheet);
    const keys = Object.keys(settingsObj);

    keys.forEach(k => {
      const idx = current.findIndex(c => c.Key === k);
      if (idx !== -1) {
        sheet.getRange(idx + 2, 2).setValue(settingsObj[k]);
      } else {
        sheet.appendRow([k, settingsObj[k]]);
      }
    });

    logAuditAction(session.userId, session.email, session.role, 'UPDATE', 'SETTINGS', 'APP_CONFIG', 'Updated application settings');
    return { success: true, message: 'Settings saved successfully.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getAuditLogs(token) {
  try {
    requireAuth(token, ['ADMIN']);
    const logs = readRows(getDb().getSheetByName(SHEETS.AUDIT_LOGS));
    return { success: true, data: logs.slice(-200).reverse() };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ======================== HELPER UTILITIES ========================

function readRows(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return [];

  const rawValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = rawValues[0].map(h => String(h).trim());
  const rows = [];

  for (let r = 1; r < rawValues.length; r++) {
    const rowObj = {};
    for (let c = 0; c < headers.length; c++) {
      let val = rawValues[r][c];
      if (val instanceof Date) {
        val = formatDate(val);
      }
      rowObj[headers[c]] = val;
    }
    rows.push(rowObj);
  }
  return rows;
}

function logAuditAction(userId, userName, role, action, moduleName, recordId, description) {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName(SHEETS.AUDIT_LOGS);
    if (!sheet) return;
    const logId = 'LOG-' + Utilities.getUuid().substring(0, 8);
    sheet.appendRow([
      logId,
      formatTimestamp(new Date()),
      userId || 'SYSTEM',
      userName || 'SYSTEM',
      role || 'SYSTEM',
      action,
      moduleName,
      recordId,
      description
    ]);
  } catch (e) {
    // Fail silently to avoid breaking primary transaction
  }
}

function validateRequired(obj, fields) {
  for (let i = 0; i < fields.length; i++) {
    const f = fields[i];
    if (obj[f] === undefined || obj[f] === null || String(obj[f]).trim() === '') {
      throw new Error(`Field '${f}' is required.`);
    }
  }
}

function formatDate(date) {
  if (!date) return '';
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatTimestamp(date) {
  if (!date) return '';
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}