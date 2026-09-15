const APP_CONFIG = {
  SPREADSHEET_ID: "1OCni4AFz5OcPQc6ZRhGLZdgWx7PdSRcoZaE7PoPnMMM",
  TIME_ZONE: "Asia/Jakarta",
  SHEETS: {
    USERS: "Pengguna",
    ROOMS: "Ruangan",
    BOOKINGS: "Peminjaman"
  },
  ROLES: ["Pengguna", "Penyetuju", "Admin"],
  USER_STATUSES: ["Menunggu", "Aktif", "Ditolak"],
  BOOKING_STATUSES: ["Menunggu", "Disetujui", "Ditolak", "Dikembalikan"]
};

const HEADERS = {
  Pengguna: ["Email", "Nama", "Role", "Status", "TanggalDaftar", "Dept"],
  Ruangan: ["ID", "NamaRuangan", "Status", "Deskripsi"],
  Peminjaman: [
    "ID", "EmailPeminjam", "NamaPeminjam", "IDRuangan", "NamaRuangan",
    "Tanggal", "WaktuMulai", "WaktuSelesai", "Tujuan", "JumlahPeserta",
    "PenyetujuEmail", "PenyetujuNama", "Status", "CatatanPenyetuju",
    "TanggalPengajuan", "StatusPersetujuan"
  ]
};

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = params.action || "";

  if (!action) {
    return HtmlService.createHtmlOutputFromFile("index")
      .setTitle("Peminjaman Ruangan")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  try {
    return outputJson_(routeGet_(action, params), params.callback);
  } catch (error) {
    return outputJson_(errorResponse_(error), params.callback);
  }
}

function doPost(e) {
  let payload = {};

  try {
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    return outputJson_(routePost_(payload), payload.callback);
  } catch (error) {
    return outputJson_(errorResponse_(error), payload.callback);
  }
}

function setupDatabase() {
  const spreadsheet = getSpreadsheet_();

  ensureSheet_(spreadsheet, APP_CONFIG.SHEETS.USERS, HEADERS.Pengguna);
  ensureSheet_(spreadsheet, APP_CONFIG.SHEETS.ROOMS, HEADERS.Ruangan);
  ensureSheet_(spreadsheet, APP_CONFIG.SHEETS.BOOKINGS, HEADERS.Peminjaman);

  const roomSheet = spreadsheet.getSheetByName(APP_CONFIG.SHEETS.ROOMS);
  if (roomSheet.getLastRow() === 1) {
    const rooms = [
      ["R002", "Ruang Business Sense", "Aktif", ""],
      ["R003", "Ruang Excellent Services", "Aktif", ""],
      ["R004", "Ruang MEETING LT.2 (EX FIF)", "Aktif", ""],
      ["R005", "LANTAI 2 HEPS", "Aktif", ""],
      ["R006", "STUDIO PHOTO AMCJ", "Aktif", ""]
    ];
    roomSheet.getRange(2, 1, rooms.length, rooms[0].length).setValues(rooms);
  }

  return {
    ok: true,
    message: "Struktur 3 tab sudah diperiksa. Data yang sudah ada tidak dihapus."
  };
}

function getCurrentUser() {
  const email = getActiveEmail_();
  const user = findUserByEmail_(email);

  return {
    ok: true,
    email: email,
    registered: Boolean(user),
    user: user ? publicUser_(user) : { Email: email }
  };
}

function getBootstrap() {
  const email = getActiveEmail_();
  const user = findUserByEmail_(email);
  const rooms = readRooms_();

  const approvers = readUsers_().filter(function(item) {
    return item.Role === "Penyetuju" && String(item.Status).toLowerCase() === "aktif";
  }).map(publicUser_);

  if (!user) {
    return {
      ok: true,
      registered: false,
      canUse: false,
      user: { Email: email },
      rooms: rooms,
      bookings: [],
      notifications: [],
      approvers: []
    };
  }

  const active = String(user.Status).toLowerCase() === "aktif";
  if (!active) {
    return {
      ok: true,
      registered: true,
      canUse: false,
      user: publicUser_(user),
      rooms: rooms,
      bookings: [],
      notifications: [],
      approvers: approvers
    };
  }

  const bookings = readBookings_();
  const visibleBookings = bookings.filter(function(booking) {
    return user.Role === "Admin" ||
      booking.EmailPeminjam === email ||
      approverEmails_(booking).indexOf(email) !== -1 ||
      booking.Status === "Disetujui";
  });

  return {
    ok: true,
    registered: true,
    canUse: true,
    user: publicUser_(user),
    rooms: rooms,
    bookings: visibleBookings.map(publicBooking_),
    notifications: buildNotifications_(user, bookings),
    approvers: approvers,
    users: user.Role === "Admin" ? readUsers_().map(publicUser_) : []
  };
}

function registerUser(data) {
  const email = getActiveEmail_();
  const name = clean_(data && data.name);
  const dept = clean_(data && data.dept);
  const role = clean_(data && data.role) || "Pengguna";

  if (!name || !dept) {
    throw new Error("Nama dan departemen wajib diisi.");
  }
  if (APP_CONFIG.ROLES.indexOf(role) === -1) {
    throw new Error("Peran yang dipilih tidak valid.");
  }

  const existing = findUserByEmail_(email);
  if (existing) {
    return {
      ok: true,
      alreadyRegistered: true,
      user: publicUser_(existing),
      message: "Email ini sudah terdaftar."
    };
  }

  const status = role === "Pengguna" ? "Aktif" : "Menunggu";
  const sheet = getSheet_(APP_CONFIG.SHEETS.USERS);
  sheet.appendRow([email, name, role, status, new Date(), dept]);

  return {
    ok: true,
    registered: true,
    canUse: status === "Aktif",
    user: {
      Email: email,
      Nama: name,
      Role: role,
      Status: status,
      Dept: dept
    },
    message: status === "Aktif"
      ? "Pendaftaran berhasil. Akun Pengguna langsung aktif."
      : "Pendaftaran berhasil. Silakan menunggu persetujuan Admin."
  };
}

function createBooking(data) {
  const email = getActiveEmail_();
  const user = requireActiveUser_(email);
  const name = clean_(data && data.name);
  const roomId = clean_(data && data.roomId);
  const date = clean_(data && data.date);
  const start = clean_(data && (data.start || data.timeStart));
  const end = clean_(data && (data.end || data.timeEnd));
  const purpose = clean_(data && data.purpose);
  const participants = Number(data && data.participants) || 0;

  if (normalize_(name) !== normalize_(user.Nama)) {
    throw new Error("Nama saat meminjam harus sama dengan nama saat mendaftar.");
  }
  if (!roomId || !date || !start || !end || !purpose) {
    throw new Error("Semua data peminjaman wajib diisi.");
  }
  if (minutes_(end) <= minutes_(start)) {
    throw new Error("Waktu selesai harus lebih besar daripada waktu mulai.");
  }
  if (date < todayKey_()) {
    throw new Error("Tanggal peminjaman tidak boleh lewat.");
  }

  const room = readRooms_().find(function(item) {
    return item.ID === roomId && isRoomActive_(item.Status);
  });
  if (!room) {
    throw new Error("Ruangan tidak tersedia.");
  }

  const activeApprovers = readUsers_().filter(function(item) {
    return item.Role === "Penyetuju" && isRoomActive_(item.Status);
  });
  if (user.Role !== "Admin" && !activeApprovers.length) {
    throw new Error("Belum ada Penyetuju aktif. Pengajuan belum dapat dibuat.");
  }

  const approverEmails = user.Role === "Admin"
    ? []
    : activeApprovers.map(function(item) { return item.Email; });
  const approverNames = user.Role === "Admin"
    ? []
    : activeApprovers.map(function(item) { return item.Nama; });
  const approvalState = {};
  approverEmails.forEach(function(approverEmail) {
    approvalState[approverEmail] = "Menunggu";
  });

  const bookings = readBookings_();
  const conflict = bookings.some(function(item) {
    return item.IDRuangan === room.ID &&
      item.Tanggal === date &&
      ["Menunggu", "Disetujui"].indexOf(item.Status) !== -1 &&
      overlaps_(start, end, item.WaktuMulai, item.WaktuSelesai);
  });
  if (conflict) {
    throw new Error("Jadwal ruangan bentrok dengan peminjaman lain.");
  }

  const id = "PMJ-" + Utilities.formatDate(new Date(), APP_CONFIG.TIME_ZONE, "yyyyMMdd-HHmmss");
  const status = user.Role === "Admin" ? "Disetujui" : "Menunggu";
  getSheet_(APP_CONFIG.SHEETS.BOOKINGS).appendRow([
    id,
    email,
    user.Nama,
    room.ID,
    room.NamaRuangan,
    date,
    start,
    end,
    purpose,
    participants,
    approverEmails.join(", "),
    approverNames.join(", "),
    status,
    "",
    new Date(),
    JSON.stringify(approvalState)
  ]);

  return {
    ok: true,
    message: status === "Disetujui"
      ? "Peminjaman Admin langsung disetujui."
      : "Pengajuan dikirim ke semua Penyetuju aktif.",
    booking: {
      ID: id,
      EmailPeminjam: email,
      NamaPeminjam: user.Nama,
      IDRuangan: room.ID,
      NamaRuangan: room.NamaRuangan,
      Tanggal: date,
      WaktuMulai: start,
      WaktuSelesai: end,
      Tujuan: purpose,
      JumlahPeserta: participants,
      PenyetujuEmail: approverEmails.join(", "),
      PenyetujuNama: approverNames.join(", "),
      Status: status,
      CatatanPenyetuju: "",
      TanggalPengajuan: todayKey_(),
      StatusPersetujuan: JSON.stringify(approvalState)
    }
  };
}

function setBookingStatus(data) {
  const email = getActiveEmail_();
  const user = requireActiveUser_(email);
  const id = clean_(data && data.id);
  const status = clean_(data && data.status);
  const note = clean_(data && data.note);

  if (["Disetujui", "Ditolak"].indexOf(status) === -1) {
    throw new Error("Status persetujuan tidak valid.");
  }

  const sheet = getSheet_(APP_CONFIG.SHEETS.BOOKINGS);
  const values = sheet.getDataRange().getValues();
  let rowNumber = -1;
  let booking = null;

  for (let row = 1; row < values.length; row += 1) {
    const item = bookingFromRow_(values[row], row + 1);
    if (item.ID === id) {
      rowNumber = row + 1;
      booking = item;
      break;
    }
  }

  if (!booking) throw new Error("Data peminjaman tidak ditemukan.");
  if (booking.Status !== "Menunggu") {
    throw new Error("Pengajuan ini sudah diproses.");
  }

  const approvers = approverEmails_(booking);
  const approvalState = approvalState_(booking);

  if (user.Role !== "Admin" && approvers.indexOf(email) === -1) {
    throw new Error("Kamu tidak memiliki izin untuk memproses pengajuan ini.");
  }

  if (user.Role === "Admin") {
    if (status === "Disetujui") {
      const conflict = readBookings_().some(function(item) {
        return item.ID !== booking.ID &&
          item.IDRuangan === booking.IDRuangan &&
          item.Tanggal === booking.Tanggal &&
          item.Status === "Disetujui" &&
          overlaps_(booking.WaktuMulai, booking.WaktuSelesai, item.WaktuMulai, item.WaktuSelesai);
      });
      if (conflict) throw new Error("Pengajuan tidak dapat disetujui karena jadwal sudah terisi.");
      approvers.forEach(function(approverEmail) {
        approvalState[approverEmail] = "Disetujui";
      });
      booking.Status = "Disetujui";
    } else {
      booking.Status = "Ditolak";
    }
  } else {
    approvalState[email] = status;
    if (status === "Ditolak") {
      booking.Status = "Ditolak";
    } else if (allApproversApproved_(booking, approvalState)) {
      const conflict = readBookings_().some(function(item) {
        return item.ID !== booking.ID &&
          item.IDRuangan === booking.IDRuangan &&
          item.Tanggal === booking.Tanggal &&
          item.Status === "Disetujui" &&
          overlaps_(booking.WaktuMulai, booking.WaktuSelesai, item.WaktuMulai, item.WaktuSelesai);
      });
      if (conflict) throw new Error("Pengajuan tidak dapat disetujui karena jadwal sudah terisi.");
      booking.Status = "Disetujui";
    } else {
      booking.Status = "Menunggu";
    }
  }

  sheet.getRange(rowNumber, 13).setValue(booking.Status);
  sheet.getRange(rowNumber, 14).setValue(note);
  sheet.getRange(rowNumber, 16).setValue(JSON.stringify(approvalState));

  let message = "Status " + id + " diubah menjadi " + booking.Status + ".";
  if (user.Role !== "Admin" && status === "Disetujui" && booking.Status === "Menunggu") {
    message = "Persetujuan kamu tersimpan. Menunggu Penyetuju lain.";
  } else if (user.Role !== "Admin" && status === "Disetujui" && booking.Status === "Disetujui") {
    message = "Semua Penyetuju sudah menyetujui peminjaman.";
  }

  return {
    ok: true,
    message: message
  };
}

function returnBooking(data) {
  const email = getActiveEmail_();
  const user = requireActiveUser_(email);
  const id = clean_(data && data.id);

  if (!id) throw new Error("ID peminjaman tidak valid.");

  const sheet = getSheet_(APP_CONFIG.SHEETS.BOOKINGS);
  const values = sheet.getDataRange().getValues();
  let rowNumber = -1;
  let booking = null;

  for (let row = 1; row < values.length; row += 1) {
    const item = bookingFromRow_(values[row], row + 1);
    if (item.ID === id) {
      rowNumber = row + 1;
      booking = item;
      break;
    }
  }

  if (!booking) throw new Error("Data peminjaman tidak ditemukan.");
  if (user.Role !== "Admin" && booking.EmailPeminjam !== email) {
    throw new Error("Kamu tidak memiliki izin untuk mengembalikan ruangan ini.");
  }
  if (booking.Status !== "Disetujui") {
    throw new Error("Hanya peminjaman yang sudah disetujui yang dapat dikembalikan.");
  }

  const endDate = bookingEndDate_(booking);
  if (!endDate || new Date().getTime() < endDate.getTime()) {
    throw new Error("Ruangan baru dapat dikembalikan setelah waktu peminjaman selesai.");
  }

  booking.Status = "Dikembalikan";
  sheet.getRange(rowNumber, 13).setValue(booking.Status);

  return {
    ok: true,
    message: "Ruangan " + booking.NamaRuangan + " berhasil dikembalikan.",
    booking: publicBooking_(booking)
  };
}

function bookingEndDate_(booking) {
  if (!booking || !booking.Tanggal || !booking.WaktuSelesai) return null;

  try {
    return Utilities.parseDate(
      booking.Tanggal + " " + booking.WaktuSelesai,
      APP_CONFIG.TIME_ZONE,
      "yyyy-MM-dd HH:mm"
    );
  } catch (error) {
    return null;
  }
}

function setUserStatus(data) {
  const admin = requireAdmin_();
  const email = clean_(data && data.email);
  const status = clean_(data && data.status);

  if (APP_CONFIG.USER_STATUSES.indexOf(status) === -1) {
    throw new Error("Status pengguna tidak valid.");
  }

  const sheet = getSheet_(APP_CONFIG.SHEETS.USERS);
  const values = sheet.getDataRange().getValues();
  let rowNumber = -1;
  let target = null;

  for (let row = 1; row < values.length; row += 1) {
    const item = userFromRow_(values[row], row + 1);
    if (item.Email === email) {
      rowNumber = row + 1;
      target = item;
      break;
    }
  }

  if (!target) throw new Error("Pengguna tidak ditemukan.");
  if (status === "Aktif" && target.Role === "Penyetuju") {
    const activeApprovers = readUsers_().filter(function(item) {
      return item.Role === "Penyetuju" &&
        item.Status === "Aktif" &&
        item.Email !== email;
    });
    if (activeApprovers.length >= 2) {
      throw new Error("Maksimal hanya 2 Penyetuju aktif.");
    }
  }

  sheet.getRange(rowNumber, 4).setValue(status);
  return {
    ok: true,
    message: "Status " + email + " diubah menjadi " + status + ".",
    admin: publicUser_(admin)
  };
}

function saveRoom(data) {
  requireAdmin_();

  const id = clean_(data && data.id);
  const name = clean_(data && data.name);
  const status = clean_(data && data.status) || "Aktif";
  const description = clean_(data && (data.description || data.desc));

  if (!id || !name) throw new Error("ID dan nama ruangan wajib diisi.");
  if (["Aktif", "Tersedia", "Nonaktif"].indexOf(status) === -1) {
    throw new Error("Status ruangan tidak valid.");
  }

  const sheet = getSheet_(APP_CONFIG.SHEETS.ROOMS);
  const values = sheet.getDataRange().getValues();

  for (let row = 1; row < values.length; row += 1) {
    if (String(values[row][0]).trim() === id) {
      sheet.getRange(row + 1, 1, 1, 4).setValues([[id, name, status, description]]);
      return { ok: true, message: "Data ruangan diperbarui." };
    }
  }

  sheet.appendRow([id, name, status, description]);
  return { ok: true, message: "Ruangan baru ditambahkan." };
}

function markNotificationsRead() {
  const email = getActiveEmail_();
  PropertiesService.getUserProperties().setProperty("NOTIFICATIONS_READ_" + email, new Date().toISOString());
  return { ok: true };
}

function routeGet_(action) {
  switch (action) {
    case "health":
      return { ok: true, message: "Penghubung Peminjaman Ruangan aktif." };
    case "currentUser":
      return getCurrentUser();
    case "bootstrap":
      return getBootstrap();
    default:
      throw new Error("Perintah GET tidak dikenal.");
  }
}

function routePost_(data) {
  const action = clean_(data && data.action);

  switch (action) {
    case "registerUser":
      return registerUser(data);
    case "createBooking":
      return createBooking(data);
    case "setBookingStatus":
      return setBookingStatus(data);
    case "returnBooking":
      return returnBooking(data);
    case "setUserStatus":
      return setUserStatus(data);
    case "saveRoom":
      return saveRoom(data);
    case "markNotificationsRead":
      return markNotificationsRead();
    default:
      throw new Error("Perintah POST tidak dikenal.");
  }
}

function getSpreadsheet_() {
  if (!APP_CONFIG.SPREADSHEET_ID ||
      APP_CONFIG.SPREADSHEET_ID === "GANTI_DENGAN_ID_SPREADSHEET_ANDA") {
    throw new Error("SPREADSHEET_ID belum diisi di Code.gs.");
  }
  return SpreadsheetApp.openById(APP_CONFIG.SPREADSHEET_ID);
}

function getSheet_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error("Tab " + sheetName + " tidak ditemukan.");
  if (HEADERS[sheetName]) ensureHeaders_(sheet, HEADERS[sheetName]);
  return sheet;
}

function ensureSheet_(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  ensureHeaders_(sheet, headers);
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const existing = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(clean_);
  let nextColumn = lastColumn;

  headers.forEach(function(header) {
    if (existing.indexOf(header) === -1) {
      nextColumn += 1;
      sheet.getRange(1, nextColumn).setValue(header);
      existing.push(header);
    }
  });

  sheet.setFrozenRows(1);
}

function readUsers_() {
  return readRows_(APP_CONFIG.SHEETS.USERS, userFromRow_);
}

function readRooms_() {
  return readRows_(APP_CONFIG.SHEETS.ROOMS, roomFromRow_);
}

function readBookings_() {
  return readRows_(APP_CONFIG.SHEETS.BOOKINGS, bookingFromRow_);
}

function readRows_(sheetName, mapper) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return values.map(function(row, index) {
    return mapper(row, index + 2);
  });
}

function userFromRow_(row, rowNumber) {
  return {
    Email: clean_(row[0]),
    Nama: clean_(row[1]),
    Role: clean_(row[2]),
    Status: clean_(row[3]),
    TanggalDaftar: dateValue_(row[4]),
    Dept: clean_(row[5]),
    _row: rowNumber
  };
}

function roomFromRow_(row, rowNumber) {
  return {
    ID: clean_(row[0]),
    NamaRuangan: clean_(row[1]),
    Status: clean_(row[2]),
    Deskripsi: clean_(row[3]),
    _row: rowNumber
  };
}

function bookingFromRow_(row, rowNumber) {
  return {
    ID: clean_(row[0]),
    EmailPeminjam: clean_(row[1]),
    NamaPeminjam: clean_(row[2]),
    IDRuangan: clean_(row[3]),
    NamaRuangan: clean_(row[4]),
    Tanggal: dateValue_(row[5]),
    WaktuMulai: timeValue_(row[6]),
    WaktuSelesai: timeValue_(row[7]),
    Tujuan: clean_(row[8]),
    JumlahPeserta: Number(row[9]) || 0,
    PenyetujuEmail: clean_(row[10]),
    PenyetujuNama: clean_(row[11]),
    Status: clean_(row[12]),
    CatatanPenyetuju: clean_(row[13]),
    TanggalPengajuan: dateValue_(row[14]),
    StatusPersetujuan: clean_(row[15]),
    _row: rowNumber
  };
}

function publicUser_(user) {
  return {
    Email: user.Email,
    Nama: user.Nama,
    Role: user.Role,
    Status: user.Status,
    TanggalDaftar: user.TanggalDaftar,
    Dept: user.Dept
  };
}

function publicBooking_(booking) {
  return {
    ID: booking.ID,
    EmailPeminjam: booking.EmailPeminjam,
    NamaPeminjam: booking.NamaPeminjam,
    IDRuangan: booking.IDRuangan,
    NamaRuangan: booking.NamaRuangan,
    Tanggal: booking.Tanggal,
    WaktuMulai: booking.WaktuMulai,
    WaktuSelesai: booking.WaktuSelesai,
    Tujuan: booking.Tujuan,
    JumlahPeserta: booking.JumlahPeserta,
    PenyetujuEmail: booking.PenyetujuEmail,
    PenyetujuNama: booking.PenyetujuNama,
    Status: booking.Status,
    CatatanPenyetuju: booking.CatatanPenyetuju,
    TanggalPengajuan: booking.TanggalPengajuan,
    StatusPersetujuan: booking.StatusPersetujuan
  };
}

function buildNotifications_(user, bookings) {
  const notifications = [];

  if (user.Role === "Admin") {
    bookings.filter(function(item) {
      return item.Status === "Menunggu";
    }).slice(-8).forEach(function(item) {
      notifications.push({
        icon: "bi-megaphone",
        text: "Pengajuan baru " + item.ID + " dari " + item.NamaPeminjam + ".",
        time: item.TanggalPengajuan,
        read: false
      });
    });
  }

  if (user.Role === "Penyetuju") {
    bookings.filter(function(item) {
      return item.Status === "Menunggu" && approverEmails_(item).indexOf(user.Email) !== -1;
    }).slice(-8).forEach(function(item) {
      notifications.push({
        icon: "bi-clipboard-check",
        text: "Ada pengajuan baru untuk " + item.NamaRuangan + ".",
        time: item.TanggalPengajuan,
        read: false
      });
    });
  }

  if (user.Role === "Pengguna") {
    bookings.filter(function(item) {
      return item.EmailPeminjam === user.Email && item.Status !== "Menunggu";
    }).slice(-8).forEach(function(item) {
      notifications.push({
        icon: item.Status === "Disetujui" ? "bi-check-circle" : "bi-x-circle",
        text: "Pengajuan " + item.ID + " " + item.Status.toLowerCase() + ".",
        time: item.TanggalPengajuan,
        read: false
      });
    });
  }

  return notifications.reverse();
}

function approverEmails_(booking) {
  return clean_(booking && booking.PenyetujuEmail)
    .split(",")
    .map(function(email) { return clean_(email); })
    .filter(Boolean);
}

function approvalState_(booking) {
  let state = {};
  try {
    const parsed = JSON.parse(clean_(booking && booking.StatusPersetujuan) || "{}");
    if (parsed && typeof parsed === "object") state = parsed;
  } catch (error) {
    state = {};
  }

  approverEmails_(booking).forEach(function(email) {
    if (!state[email]) state[email] = "Menunggu";
  });
  return state;
}

function allApproversApproved_(booking, state) {
  const approvers = approverEmails_(booking);
  return approvers.length > 0 && approvers.every(function(email) {
    return state[email] === "Disetujui";
  });
}

function findUserByEmail_(email) {
  return readUsers_().find(function(user) {
    return user.Email.toLowerCase() === email.toLowerCase();
  }) || null;
}

function requireActiveUser_(email) {
  const user = findUserByEmail_(email);
  if (!user) throw new Error("Email belum terdaftar. Silakan daftar terlebih dahulu.");
  if (String(user.Status).toLowerCase() !== "aktif") {
    throw new Error("Akun belum aktif. Silakan menunggu persetujuan Admin.");
  }
  return user;
}

function requireAdmin_() {
  const user = requireActiveUser_(getActiveEmail_());
  if (user.Role !== "Admin") throw new Error("Hanya Admin yang boleh melakukan tindakan ini.");
  return user;
}

function getActiveEmail_() {
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error("Email Google tidak terbaca. Buka aplikasi setelah login dengan akun Google.");
  }
  return email.toLowerCase().trim();
}

function dateValue_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, APP_CONFIG.TIME_ZONE, "yyyy-MM-dd");
  }
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const parsed = new Date(text);
  return isNaN(parsed.getTime())
    ? text
    : Utilities.formatDate(parsed, APP_CONFIG.TIME_ZONE, "yyyy-MM-dd");
}

function timeValue_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, APP_CONFIG.TIME_ZONE, "HH:mm");
  }
  const text = String(value).trim();
  const match = text.match(/^(\d{1,2}):(\d{2})/);
  return match ? String(match[1]).padStart(2, "0") + ":" + match[2] : text;
}

function isRoomActive_(status) {
  return ["aktif", "tersedia", "available"].indexOf(clean_(status).toLowerCase()) !== -1;
}

function todayKey_() {
  return Utilities.formatDate(new Date(), APP_CONFIG.TIME_ZONE, "yyyy-MM-dd");
}

function minutes_(value) {
  const parts = String(value || "").split(":");
  return Number(parts[0]) * 60 + Number(parts[1]);
}

function overlaps_(startA, endA, startB, endB) {
  return minutes_(startA) < minutes_(endB) && minutes_(endA) > minutes_(startB);
}

function clean_(value) {
  return String(value == null ? "" : value).trim();
}

function normalize_(value) {
  return clean_(value).replace(/\s+/g, " ").toLowerCase();
}

function errorResponse_(error) {
  return {
    ok: false,
    error: error && error.message ? error.message : String(error)
  };
}

function outputJson_(data, callback) {
  const json = JSON.stringify(data);

  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return HtmlService.createHtmlOutput(callback + "(" + json + ");")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
