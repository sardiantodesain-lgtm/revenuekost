function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard Revenue Kost')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function getSheet(sheetName) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName); }

// --- AUTENTIKASI ---
function registerUser(username, password) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(username).trim()) return { success: false, message: 'Username sudah digunakan.' };
  }
  const userId = 'USR' + new Date().getTime();
  sheet.appendRow([userId, username, password, new Date()]);
  return { success: true, message: 'Registrasi berhasil!' };
}

function loginUser(username, password) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(username).trim() && String(data[i][2]).trim() === String(password).trim()) {
      return { success: true, userId: data[i][0], username: username, password: password };
    }
  }
  return { success: false, message: 'Username atau password salah.' };
}

function updateAccount(userId, newUsername, newPassword) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      if(newUsername) sheet.getRange(i + 1, 2).setValue(newUsername);
      if(newPassword) sheet.getRange(i + 1, 3).setValue(newPassword);
      return { success: true, message: 'Akun berhasil diperbarui!', newUsername: newUsername };
    }
  }
  return { success: false, message: 'User tidak ditemukan.' };
}

// --- MENU 1: REVENUE ---
function addRevenue(userId, date, room, tenant, category, amount, note) {
  const sheet = getSheet('Revenues');
  sheet.appendRow([userId, date, room, tenant, category, amount, note, new Date()]);
  return { success: true, message: 'Pendapatan berhasil dicatat!' };
}

// --- MENU 2: DASHBOARD & GRAFIK ---
function getDashboardData(userId, month, year) {
  const sheet = getSheet('Revenues');
  if (sheet.getLastRow() < 2) return { total: 0, chartData: [] };
  
  const data = sheet.getDataRange().getValues();
  let totalRevenue = 0;
  let dailyData = {}; // Format: { "1": 1500000, "5": 500000 }
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      const rowDate = new Date(data[i][1]);
      // Periksa apakah bulan dan tahun cocok (bulan di JS dimulai dari 0, jadi bulan inputan - 1)
      if (rowDate.getMonth() + 1 == month && rowDate.getFullYear() == year) {
        const amount = Number(data[i][5]);
        const day = rowDate.getDate();
        totalRevenue += amount;
        
        if (dailyData[day]) dailyData[day] += amount;
        else dailyData[day] = amount;
      }
    }
  }
  
  return { total: totalRevenue, chartData: dailyData };
}

// --- MENU 3: MAINTENANCE ---
function addMaintenance(userId, room, material, price, qty, total) {
  const sheet = getSheet('Maintenance');
  sheet.appendRow([userId, room, material, price, qty, total, new Date()]);
  return { success: true, message: 'Anggaran perbaikan disimpan!' };
}