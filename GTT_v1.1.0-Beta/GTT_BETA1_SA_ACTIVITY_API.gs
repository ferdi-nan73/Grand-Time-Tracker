/**
 * GRAND TIME TRACKER (GTT)
 * GTT v1.0.0-Beta.2 — SA DAILY ACTIVITY
 *
 * ADDITIVE / ISOLATED MODULE ONLY.
 * Tidak mengubah fungsi, variabel, class, atau logika existing yang sudah LOCKED.
 *
 * Modul baru:
 * - Tugas Luar multi-outlet
 * - Approval di HP SL/SPV/Manager
 * - Riwayat + durasi per Tugas Luar
 * - Monitoring Tugas Luar + Break per outlet
 * - Geofence Start/Finish outlet asal
 * - Cancel pending/approved + edit yang memicu re-approval
 * - Ringkasan TL harian PUBLIC per outlet (display-only daily reset)
 */

var GTT_ACTIVITY_BETA1 = Object.freeze({
  VERSION: 'v1.0.0-Beta.2',
  SPREADSHEET_ID: '1pjiQEnuox8-onLat4frlj3HKDX7OKAFmDbJMxopiuhk',
  TIME_ZONE: 'Asia/Makassar',
  SHEET: Object.freeze({
    MASTER_SA: 'MASTER_SA',
    MASTER_LOKASI: 'MASTER_LOKASI',
    TUGAS_LUAR: 'TUGAS_LUAR',
    LOG_ISTIRAHAT: 'LOG_ISTIRAHAT'
  }),
  STATUS: Object.freeze({
    PENDING: 'PENDING_APPROVAL',
    APPROVED: 'APPROVED',
    IN_PROGRESS: 'IN_PROGRESS',
    COMPLETED: 'COMPLETED',
    REJECTED: 'REJECTED',
    CANCELLED: 'CANCELLED'
  }),
  APPROVAL: Object.freeze({
    PENDING: 'PENDING',
    APPROVED: 'APPROVED',
    REJECTED: 'REJECTED',
    CANCELLED: 'CANCELLED'
  }),
  APPROVER_POSITIONS: Object.freeze(['SL', 'SPV', 'MANAGER']),
  MAX_ROUTE_STOPS: 10,
  MAX_GPS_ACCURACY_M: 150,
  EXTRA_HEADERS: Object.freeze([
    'SA_ID',
    'TASK_NO_HARI',
    'RUTE_JSON',
    'RUTE_LABEL',
    'APPROVAL_STATUS',
    'APPROVER_ID',
    'APPROVER_NAMA',
    'APPROVER_JABATAN',
    'REQUESTED_AT',
    'APPROVED_AT',
    'REJECTED_AT',
    'REJECTION_REASON',
    'CANCELLED_AT',
    'CANCELLED_BY',
    'CANCEL_REASON',
    'START_TS_MS',
    'END_TS_MS',
    'DURATION_SECONDS',
    'APPROVAL_HISTORY_JSON',
    'SOURCE_VERSION'
  ])
});


/* ============================================================
 * PUBLIC API — FRONTEND GTT_BETA1_SA_ACTIVITY_APPEND.html
 * ============================================================ */

/**
 * Bootstrap seluruh data Menu Aktivitas untuk user login.
 * @param {string|number} pinInput
 * @return {Object}
 */
function apiGttActBootstrap(pinInput) {
  return gttActSafeApi_('BOOTSTRAP', function () {
    var ss = gttActSpreadsheet_();
    var user = gttActGetUserByPin_(ss, pinInput);
    var taskCtx = gttActTaskContext_(ss, true);
    return {
      success: true,
      code: 'GTT_ACT_BOOTSTRAP_OK',
      data: gttActBuildDashboardData_(ss, taskCtx, user)
    };
  });
}

/**
 * Membuat request approval Tugas Luar baru.
 * @param {string|number} pinInput
 * @param {Object} payload {route, purpose, approverId}
 * @return {Object}
 */
function apiGttActSubmitApproval(pinInput, payload) {
  return gttActSafeApi_('SUBMIT_APPROVAL', function () {
    return gttActWithLock_(function () {
      var ss = gttActSpreadsheet_();
      var user = gttActGetUserByPin_(ss, pinInput);
      var taskCtx = gttActTaskContext_(ss, true);
      var now = new Date();
      var dateKey = gttActDateKey_(now);

      gttActAssertNoOpenTask_(taskCtx, user, 'Anda masih memiliki Tugas Luar/permintaan yang belum selesai.');

      var prepared = gttActPrepareRequestPayload_(ss, user, payload);
      var rowObj = {};
      var taskId = gttActNewTaskId_(now, user.saId);

      rowObj['ID TUGAS'] = taskId;
      rowObj['ID ABSENSI'] = '';
      rowObj['TANGGAL'] = now;
      rowObj['PIN'] = user.pin;
      rowObj['NAMA SA'] = user.nama;
      rowObj['OUTLET ASAL'] = user.outlet;
      rowObj['JENIS TUGAS'] = 'TUGAS LUAR';
      rowObj['TUJUAN'] = prepared.routeLabel;
      rowObj['NO REFERENSI'] = '';
      rowObj['KETERANGAN'] = prepared.purpose;
      rowObj['JAM MULAI'] = '';
      rowObj['LAT MULAI'] = '';
      rowObj['LNG MULAI'] = '';
      rowObj['AKURASI MULAI (M)'] = '';
      rowObj['JAM KEMBALI'] = '';
      rowObj['LAT KEMBALI'] = '';
      rowObj['LNG KEMBALI'] = '';
      rowObj['AKURASI KEMBALI (M)'] = '';
      rowObj['JARAK KEMBALI (M)'] = '';
      rowObj['DURASI (MENIT)'] = '';
      rowObj['STATUS'] = GTT_ACTIVITY_BETA1.STATUS.PENDING;
      rowObj['PENANDA DURASI'] = '';
      rowObj['ALASAN BATAL'] = '';
      rowObj['DIBUAT PADA'] = now;
      rowObj['DIPERBARUI PADA'] = now;

      rowObj['SA_ID'] = user.saId;
      rowObj['TASK_NO_HARI'] = '';
      rowObj['RUTE_JSON'] = JSON.stringify(prepared.route);
      rowObj['RUTE_LABEL'] = prepared.routeLabel;
      rowObj['APPROVAL_STATUS'] = GTT_ACTIVITY_BETA1.APPROVAL.PENDING;
      rowObj['APPROVER_ID'] = prepared.approver.saId;
      rowObj['APPROVER_NAMA'] = prepared.approver.nama;
      rowObj['APPROVER_JABATAN'] = prepared.approver.jabatan;
      rowObj['REQUESTED_AT'] = now;
      rowObj['APPROVED_AT'] = '';
      rowObj['REJECTED_AT'] = '';
      rowObj['REJECTION_REASON'] = '';
      rowObj['CANCELLED_AT'] = '';
      rowObj['CANCELLED_BY'] = '';
      rowObj['CANCEL_REASON'] = '';
      rowObj['START_TS_MS'] = '';
      rowObj['END_TS_MS'] = '';
      rowObj['DURATION_SECONDS'] = '';
      rowObj['APPROVAL_HISTORY_JSON'] = '[]';
      rowObj['SOURCE_VERSION'] = GTT_ACTIVITY_BETA1.VERSION;

      gttActAppendRowObject_(taskCtx, rowObj);

      return {
        success: true,
        code: 'GTT_ACT_APPROVAL_REQUESTED',
        message: 'Permintaan persetujuan Tugas Luar berhasil dikirim.',
        taskId: taskId,
        dateKey: dateKey,
        data: gttActBuildDashboardData_(ss, gttActTaskContext_(ss, false), user)
      };
    });
  });
}

/**
 * Mengganti approver untuk request PENDING. Request lama dibatalkan secara audit,
 * tetapi tetap memakai satu row/transaksi yang sama.
 */
function apiGttActChangeApprover(pinInput, taskId, newApproverId) {
  return gttActSafeApi_('CHANGE_APPROVER', function () {
    return gttActWithLock_(function () {
      var ss = gttActSpreadsheet_();
      var user = gttActGetUserByPin_(ss, pinInput);
      var taskCtx = gttActTaskContext_(ss, true);
      var task = gttActFindTaskById_(taskCtx, taskId);
      gttActAssertOwner_(task, user);

      if (task.status !== GTT_ACTIVITY_BETA1.STATUS.PENDING) {
        throw new Error('Approver hanya dapat diganti saat permintaan masih PENDING.');
      }

      var approver = gttActGetValidApprover_(ss, user, newApproverId);
      var now = new Date();
      var history = gttActApprovalHistory_(task);
      history.push({
        action: 'CHANGE_APPROVER',
        fromApproverId: task.approverId || '',
        fromApproverName: task.approverName || '',
        toApproverId: approver.saId,
        toApproverName: approver.nama,
        at: now.toISOString()
      });

      gttActUpdateTask_(taskCtx, task.rowNumber, {
        'APPROVER_ID': approver.saId,
        'APPROVER_NAMA': approver.nama,
        'APPROVER_JABATAN': approver.jabatan,
        'APPROVAL_STATUS': GTT_ACTIVITY_BETA1.APPROVAL.PENDING,
        'REQUESTED_AT': now,
        'APPROVAL_HISTORY_JSON': JSON.stringify(history),
        'DIPERBARUI PADA': now
      });

      return {
        success: true,
        code: 'GTT_ACT_APPROVER_CHANGED',
        message: 'Approver berhasil diganti.',
        data: gttActBuildDashboardData_(ss, gttActTaskContext_(ss, false), user)
      };
    });
  });
}

/**
 * Mengubah rute/keperluan sebelum Tugas Luar dimulai.
 * Jika sudah APPROVED, approval otomatis kembali PENDING (re-approval).
 */
function apiGttActUpdateRequest(pinInput, taskId, payload) {
  return gttActSafeApi_('UPDATE_REQUEST', function () {
    return gttActWithLock_(function () {
      var ss = gttActSpreadsheet_();
      var user = gttActGetUserByPin_(ss, pinInput);
      var taskCtx = gttActTaskContext_(ss, true);
      var task = gttActFindTaskById_(taskCtx, taskId);
      gttActAssertOwner_(task, user);

      if (
        task.status !== GTT_ACTIVITY_BETA1.STATUS.PENDING &&
        task.status !== GTT_ACTIVITY_BETA1.STATUS.APPROVED
      ) {
        throw new Error('Data Tugas Luar hanya dapat diubah sebelum Tugas Luar dimulai.');
      }

      var prepared = gttActPrepareRequestPayload_(ss, user, payload);
      var now = new Date();
      var history = gttActApprovalHistory_(task);
      history.push({
        action: 'EDIT_REQUEST_REAPPROVAL',
        previousApprovalStatus: task.approvalStatus || '',
        at: now.toISOString()
      });

      gttActUpdateTask_(taskCtx, task.rowNumber, {
        'TUJUAN': prepared.routeLabel,
        'KETERANGAN': prepared.purpose,
        'RUTE_JSON': JSON.stringify(prepared.route),
        'RUTE_LABEL': prepared.routeLabel,
        'STATUS': GTT_ACTIVITY_BETA1.STATUS.PENDING,
        'APPROVAL_STATUS': GTT_ACTIVITY_BETA1.APPROVAL.PENDING,
        'APPROVER_ID': prepared.approver.saId,
        'APPROVER_NAMA': prepared.approver.nama,
        'APPROVER_JABATAN': prepared.approver.jabatan,
        'REQUESTED_AT': now,
        'APPROVED_AT': '',
        'REJECTED_AT': '',
        'REJECTION_REASON': '',
        'APPROVAL_HISTORY_JSON': JSON.stringify(history),
        'DIPERBARUI PADA': now
      });

      return {
        success: true,
        code: 'GTT_ACT_REQUEST_UPDATED',
        message: 'Data diperbarui dan dikirim ulang untuk persetujuan.',
        data: gttActBuildDashboardData_(ss, gttActTaskContext_(ss, false), user)
      };
    });
  });
}

/**
 * Approval hanya sah jika dilakukan oleh akun/HP approver yang dipilih.
 */
function apiGttActApprove(pinInput, taskId) {
  return gttActSafeApi_('APPROVE', function () {
    return gttActWithLock_(function () {
      var ss = gttActSpreadsheet_();
      var approverUser = gttActGetUserByPin_(ss, pinInput);
      var taskCtx = gttActTaskContext_(ss, true);
      var task = gttActFindTaskById_(taskCtx, taskId);

      gttActAssertApproverCanAct_(task, approverUser);
      if (task.status !== GTT_ACTIVITY_BETA1.STATUS.PENDING) {
        throw new Error('Permintaan ini sudah tidak berstatus PENDING.');
      }

      var now = new Date();
      var taskNo = Number(task.taskNo || 0) || gttActNextTaskNo_(taskCtx, task.pin, task.dateKey);
      var history = gttActApprovalHistory_(task);
      history.push({
        action: 'APPROVED',
        approverId: approverUser.saId,
        approverName: approverUser.nama,
        at: now.toISOString()
      });

      gttActUpdateTask_(taskCtx, task.rowNumber, {
        'TASK_NO_HARI': taskNo,
        'STATUS': GTT_ACTIVITY_BETA1.STATUS.APPROVED,
        'APPROVAL_STATUS': GTT_ACTIVITY_BETA1.APPROVAL.APPROVED,
        'APPROVED_AT': now,
        'REJECTED_AT': '',
        'REJECTION_REASON': '',
        'APPROVAL_HISTORY_JSON': JSON.stringify(history),
        'DIPERBARUI PADA': now
      });

      return {
        success: true,
        code: 'GTT_ACT_APPROVED',
        message: 'Tugas Luar berhasil disetujui.',
        data: gttActBuildDashboardData_(ss, gttActTaskContext_(ss, false), approverUser)
      };
    });
  });
}

function apiGttActReject(pinInput, taskId) {
  return gttActSafeApi_('REJECT', function () {
    return gttActWithLock_(function () {
      var ss = gttActSpreadsheet_();
      var approverUser = gttActGetUserByPin_(ss, pinInput);
      var taskCtx = gttActTaskContext_(ss, true);
      var task = gttActFindTaskById_(taskCtx, taskId);

      gttActAssertApproverCanAct_(task, approverUser);
      if (task.status !== GTT_ACTIVITY_BETA1.STATUS.PENDING) {
        throw new Error('Permintaan ini sudah tidak berstatus PENDING.');
      }

      var now = new Date();
      var reason = 'Ditolak oleh ' + approverUser.nama;
      var history = gttActApprovalHistory_(task);
      history.push({
        action: 'REJECTED',
        approverId: approverUser.saId,
        approverName: approverUser.nama,
        at: now.toISOString()
      });

      gttActUpdateTask_(taskCtx, task.rowNumber, {
        'STATUS': GTT_ACTIVITY_BETA1.STATUS.REJECTED,
        'APPROVAL_STATUS': GTT_ACTIVITY_BETA1.APPROVAL.REJECTED,
        'REJECTED_AT': now,
        'REJECTION_REASON': reason,
        'APPROVAL_HISTORY_JSON': JSON.stringify(history),
        'DIPERBARUI PADA': now
      });

      return {
        success: true,
        code: 'GTT_ACT_REJECTED',
        message: 'Permintaan Tugas Luar ditolak.',
        data: gttActBuildDashboardData_(ss, gttActTaskContext_(ss, false), approverUser)
      };
    });
  });
}

/**
 * Start hanya setelah APPROVED dan user berada di geofence outlet asal.
 */
function apiGttActStart(pinInput, taskId, location) {
  return gttActSafeApi_('START', function () {
    return gttActWithLock_(function () {
      var ss = gttActSpreadsheet_();
      var user = gttActGetUserByPin_(ss, pinInput);
      var taskCtx = gttActTaskContext_(ss, true);
      var task = gttActFindTaskById_(taskCtx, taskId);
      gttActAssertOwner_(task, user);

      if (
        task.status !== GTT_ACTIVITY_BETA1.STATUS.APPROVED ||
        task.approvalStatus !== GTT_ACTIVITY_BETA1.APPROVAL.APPROVED
      ) {
        throw new Error('Tugas Luar belum mendapat persetujuan SL/SPV/Manager.');
      }

      gttActAssertNoOtherInProgress_(taskCtx, user, task.id);
      var activeBreak = gttActFindActiveBreak_(ss, user.pin, user.outlet, new Date());
      if (activeBreak) {
        throw new Error('Tugas Luar tidak dapat dimulai saat ' + activeBreak.label + ' masih berlangsung.');
      }

      var geo = gttActValidateOriginGeofence_(ss, user.outlet, location);
      var now = new Date();

      gttActUpdateTask_(taskCtx, task.rowNumber, {
        'STATUS': GTT_ACTIVITY_BETA1.STATUS.IN_PROGRESS,
        'JAM MULAI': now,
        'LAT MULAI': geo.latitude,
        'LNG MULAI': geo.longitude,
        'AKURASI MULAI (M)': geo.accuracy,
        'START_TS_MS': now.getTime(),
        'DIPERBARUI PADA': now
      });

      return {
        success: true,
        code: 'GTT_ACT_STARTED',
        message: 'Tugas Luar ' + (task.taskNo || '') + ' berhasil dimulai.',
        geofence: geo,
        data: gttActBuildDashboardData_(ss, gttActTaskContext_(ss, false), user)
      };
    });
  });
}

/**
 * Finish hanya setelah kembali ke geofence outlet asal.
 */
function apiGttActFinish(pinInput, taskId, location) {
  return gttActSafeApi_('FINISH', function () {
    return gttActWithLock_(function () {
      var ss = gttActSpreadsheet_();
      var user = gttActGetUserByPin_(ss, pinInput);
      var taskCtx = gttActTaskContext_(ss, true);
      var task = gttActFindTaskById_(taskCtx, taskId);
      gttActAssertOwner_(task, user);

      if (task.status !== GTT_ACTIVITY_BETA1.STATUS.IN_PROGRESS) {
        throw new Error('Tidak ada Tugas Luar aktif yang dapat diselesaikan.');
      }

      var geo = gttActValidateOriginGeofence_(ss, user.outlet, location);
      var now = new Date();
      var startMs = Number(task.startTsMs || 0);
      if (!startMs) {
        throw new Error('Waktu mulai Tugas Luar tidak ditemukan.');
      }

      var durationSeconds = Math.max(0, Math.floor((now.getTime() - startMs) / 1000));
      var durationMinutes = durationSeconds > 0 ? Math.max(1, Math.ceil(durationSeconds / 60)) : 0;

      gttActUpdateTask_(taskCtx, task.rowNumber, {
        'STATUS': GTT_ACTIVITY_BETA1.STATUS.COMPLETED,
        'JAM KEMBALI': now,
        'LAT KEMBALI': geo.latitude,
        'LNG KEMBALI': geo.longitude,
        'AKURASI KEMBALI (M)': geo.accuracy,
        'JARAK KEMBALI (M)': Math.round(geo.distanceMeters),
        'DURASI (MENIT)': durationMinutes,
        'PENANDA DURASI': 'SELESAI',
        'END_TS_MS': now.getTime(),
        'DURATION_SECONDS': durationSeconds,
        'DIPERBARUI PADA': now
      });

      return {
        success: true,
        code: 'GTT_ACT_COMPLETED',
        message: 'Tugas Luar berhasil diselesaikan. Durasi ' + durationMinutes + ' menit.',
        durationSeconds: durationSeconds,
        durationMinutes: durationMinutes,
        geofence: geo,
        data: gttActBuildDashboardData_(ss, gttActTaskContext_(ss, false), user)
      };
    });
  });
}

/**
 * SA boleh cancel sendiri hanya saat PENDING / APPROVED (belum mulai).
 */
function apiGttActCancel(pinInput, taskId, reasonInput) {
  return gttActSafeApi_('CANCEL', function () {
    return gttActWithLock_(function () {
      var ss = gttActSpreadsheet_();
      var user = gttActGetUserByPin_(ss, pinInput);
      var taskCtx = gttActTaskContext_(ss, true);
      var task = gttActFindTaskById_(taskCtx, taskId);
      gttActAssertOwner_(task, user);

      if (
        task.status !== GTT_ACTIVITY_BETA1.STATUS.PENDING &&
        task.status !== GTT_ACTIVITY_BETA1.STATUS.APPROVED
      ) {
        throw new Error('Tugas Luar yang sudah berjalan tidak dapat dibatalkan sendiri. Hubungi SL/SPV/Manager.');
      }

      var now = new Date();
      var reason = String(reasonInput || 'Dibatalkan oleh SA sebelum Tugas Luar dimulai.').trim();
      if (reason.length > 250) reason = reason.slice(0, 250);
      var history = gttActApprovalHistory_(task);
      history.push({
        action: 'CANCELLED_BY_SA',
        by: user.saId,
        at: now.toISOString()
      });

      gttActUpdateTask_(taskCtx, task.rowNumber, {
        'STATUS': GTT_ACTIVITY_BETA1.STATUS.CANCELLED,
        'APPROVAL_STATUS': GTT_ACTIVITY_BETA1.APPROVAL.CANCELLED,
        'ALASAN BATAL': reason,
        'CANCELLED_AT': now,
        'CANCELLED_BY': user.saId,
        'CANCEL_REASON': reason,
        'APPROVAL_HISTORY_JSON': JSON.stringify(history),
        'DIPERBARUI PADA': now
      });

      return {
        success: true,
        code: 'GTT_ACT_CANCELLED',
        message: 'Tugas Luar dibatalkan.',
        data: gttActBuildDashboardData_(ss, gttActTaskContext_(ss, false), user)
      };
    });
  });
}


/* ============================================================
 * INTERNAL — DASHBOARD DATA
 * ============================================================ */

function gttActBuildDashboardData_(ss, taskCtx, user) {
  var now = new Date();
  var dateKey = gttActDateKey_(now);
  var tasks = gttActReadTasks_(taskCtx);
  var myToday = [];
  var teamOutside = [];
  var outletSummaryToday = [];
  var pendingApprovals = [];
  var currentTask = null;
  var i;

  for (i = 0; i < tasks.length; i++) {
    var task = tasks[i];

    if (task.pin === user.pin && task.dateKey === dateKey) {
      myToday.push(task);
    }

    // Open task tetap dipulihkan walau melewati pergantian tanggal.
    if (
      task.pin === user.pin &&
      (
        task.status === GTT_ACTIVITY_BETA1.STATUS.PENDING ||
        task.status === GTT_ACTIVITY_BETA1.STATUS.APPROVED ||
        task.status === GTT_ACTIVITY_BETA1.STATUS.IN_PROGRESS
      )
    ) {
      if (!currentTask || task.rowNumber > currentTask.rowNumber) currentTask = task;
    }

    if (
      task.outlet === user.outlet &&
      task.status === GTT_ACTIVITY_BETA1.STATUS.IN_PROGRESS
    ) {
      teamOutside.push(gttActTaskPublicView_(task, now));
    }

    // Beta.2: 1 baris = 1 transaksi TL yang SUDAH DIMULAI pada tanggal kerja hari ini.
    // Data lama tetap di database; hanya tampilan yang otomatis terfilter saat tanggal berganti.
    var summaryDateKey = task.startTsMs
      ? gttActDateKey_(new Date(Number(task.startTsMs)))
      : task.dateKey;
    if (
      task.outlet === user.outlet &&
      summaryDateKey === dateKey &&
      Number(task.taskNo || 0) > 0 &&
      (
        task.status === GTT_ACTIVITY_BETA1.STATUS.IN_PROGRESS ||
        task.status === GTT_ACTIVITY_BETA1.STATUS.COMPLETED
      )
    ) {
      outletSummaryToday.push(gttActTaskSummaryView_(task, now));
    }

    if (
      task.outlet === user.outlet &&
      task.status === GTT_ACTIVITY_BETA1.STATUS.PENDING &&
      gttActIsApproverPosition_(user.jabatan)
    ) {
      pendingApprovals.push({
        taskId: task.id,
        taskNo: task.taskNo,
        nama: task.nama,
        saId: task.saId,
        outlet: task.outlet,
        route: task.route,
        routeLabel: task.routeLabel,
        purpose: task.purpose,
        approverId: task.approverId,
        approverName: task.approverName,
        approverPosition: task.approverPosition,
        requestedAtText: gttActFormatDateTime_(task.requestedAt),
        canAct: String(task.approverId || '') === String(user.saId || '')
      });
    }
  }

  myToday.sort(function (a, b) { return a.rowNumber - b.rowNumber; });
  teamOutside.sort(function (a, b) { return String(a.nama).localeCompare(String(b.nama)); });
  // Ringkasan mengikuti urutan transaksi aktual; nama SA boleh muncul lagi pada TL 2/TL 3.
  outletSummaryToday.sort(function (a, b) { return a.rowNumber - b.rowNumber; });
  pendingApprovals.sort(function (a, b) { return String(a.requestedAtText).localeCompare(String(b.requestedAtText)); });

  var breakData = gttActGetTeamBreaks_(ss, user.outlet, now);
  var destinations = gttActGetDestinations_(ss, user.outlet);
  var approvers = gttActGetApprovers_(ss, user);

  return {
    version: GTT_ACTIVITY_BETA1.VERSION,
    serverNowMs: now.getTime(),
    dateKey: dateKey,
    dateLabel: gttActDateLabelId_(now),
    user: {
      saId: user.saId,
      nama: user.nama,
      pin: user.pin,
      outlet: user.outlet,
      jabatan: user.jabatan,
      role: user.role,
      isApprover: gttActIsApproverPosition_(user.jabatan)
    },
    destinations: destinations,
    approvers: approvers,
    currentTask: currentTask ? gttActTaskPrivateView_(currentTask, now) : null,
    myToday: myToday.map(function (task) { return gttActTaskPrivateView_(task, now); }),
    team: {
      outside: teamOutside,
      break1: breakData.break1,
      break2: breakData.break2,
      tlToday: outletSummaryToday
    },
    pendingApprovals: pendingApprovals
  };
}

function gttActTaskPublicView_(task, now) {
  return {
    taskId: task.id,
    taskNo: task.taskNo,
    nama: task.nama,
    outlet: task.outlet,
    route: task.route,
    routeLabel: task.routeLabel,
    startTime: task.startTimeText,
    startTsMs: task.startTsMs,
    elapsedSeconds: gttActElapsedSeconds_(task, now),
    status: task.status
  };
}

function gttActTaskSummaryView_(task, now) {
  return {
    rowNumber: task.rowNumber,
    taskId: task.id,
    taskNo: task.taskNo,
    nama: task.nama,
    saId: task.saId,
    outlet: task.outlet,
    routeLabel: task.routeLabel,
    startTime: task.startTimeText,
    endTime: task.endTimeText,
    startTsMs: task.startTsMs,
    endTsMs: task.endTsMs,
    durationSeconds: task.durationSeconds,
    durationMinutes: task.durationMinutes,
    elapsedSeconds: gttActElapsedSeconds_(task, now),
    status: task.status
  };
}

function gttActTaskPrivateView_(task, now) {
  return {
    taskId: task.id,
    taskNo: task.taskNo,
    dateKey: task.dateKey,
    nama: task.nama,
    saId: task.saId,
    outlet: task.outlet,
    route: task.route,
    routeLabel: task.routeLabel,
    purpose: task.purpose,
    status: task.status,
    approvalStatus: task.approvalStatus,
    approverId: task.approverId,
    approverName: task.approverName,
    approverPosition: task.approverPosition,
    requestedAtText: gttActFormatDateTime_(task.requestedAt),
    approvedAtText: gttActFormatDateTime_(task.approvedAt),
    rejectedAtText: gttActFormatDateTime_(task.rejectedAt),
    rejectionReason: task.rejectionReason,
    startTime: task.startTimeText,
    endTime: task.endTimeText,
    startTsMs: task.startTsMs,
    endTsMs: task.endTsMs,
    durationSeconds: task.durationSeconds,
    durationMinutes: task.durationMinutes,
    elapsedSeconds: gttActElapsedSeconds_(task, now),
    cancelReason: task.cancelReason
  };
}


/* ============================================================
 * INTERNAL — VALIDATION / MASTER DATA
 * ============================================================ */

function gttActPrepareRequestPayload_(ss, user, payload) {
  var source = payload || {};
  var route = gttActValidateRoute_(ss, user, source.route);
  var purpose = String(source.purpose || '').trim();
  if (!purpose) throw new Error('Keperluan Tugas Luar wajib diisi.');
  if (purpose.length > 250) purpose = purpose.slice(0, 250);

  var approver = gttActGetValidApprover_(ss, user, source.approverId);
  return {
    route: route,
    routeLabel: gttActRouteLabel_(route),
    purpose: purpose,
    approver: approver
  };
}

function gttActValidateRoute_(ss, user, routeInput) {
  var raw = Array.isArray(routeInput) ? routeInput : [];
  if (!raw.length) throw new Error('Pilih minimal 1 tujuan Tugas Luar.');
  if (raw.length > GTT_ACTIVITY_BETA1.MAX_ROUTE_STOPS) {
    throw new Error('Maksimal ' + GTT_ACTIVITY_BETA1.MAX_ROUTE_STOPS + ' tujuan dalam satu Tugas Luar.');
  }

  var locations = gttActLocationMap_(ss);
  var result = [];
  for (var i = 0; i < raw.length; i++) {
    var item = raw[i] || {};
    var type = gttActNorm_(item.type || 'OUTLET');

    if (type === 'OTHER' || type === 'LAINNYA') {
      var otherName = String(item.name || item.label || '').trim();
      if (!otherName) throw new Error('Tujuan Lainnya pada urutan ' + (i + 1) + ' wajib diisi.');
      if (otherName.length > 120) otherName = otherName.slice(0, 120);
      result.push({
        order: i + 1,
        type: 'OTHER',
        code: 'LAINNYA',
        name: otherName,
        display: otherName
      });
      continue;
    }

    var code = gttActNorm_(item.code || item.value || '');
    if (!code || !locations[code] || !locations[code].active) {
      throw new Error('Tujuan outlet pada urutan ' + (i + 1) + ' tidak valid/aktif.');
    }
    if (code === gttActNorm_(user.outlet)) {
      throw new Error('Outlet asal ' + user.outlet + ' tidak dapat dipilih sebagai tujuan.');
    }

    // Duplicate destination sengaja DIIZINKAN sesuai business flow.
    result.push({
      order: i + 1,
      type: 'OUTLET',
      code: code,
      name: locations[code].name,
      display: code
    });
  }

  return result;
}

function gttActRouteLabel_(route) {
  return (route || []).map(function (item) {
    return item.type === 'OTHER' ? item.name : item.code;
  }).join(' → ');
}

function gttActGetDestinations_(ss, originOutlet) {
  var map = gttActLocationMap_(ss);
  return Object.keys(map)
    .filter(function (code) {
      return map[code].active && code !== gttActNorm_(originOutlet);
    })
    .map(function (code) {
      return { code: code, name: map[code].name };
    });
}

function gttActGetApprovers_(ss, user) {
  var sheet = ss.getSheetByName(GTT_ACTIVITY_BETA1.SHEET.MASTER_SA);
  if (!sheet) throw new Error('Sheet MASTER_SA tidak ditemukan.');
  var values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return [];
  var map = gttActHeaderMap_(values[0]);
  var result = [];

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var status = gttActNorm_(gttActCell_(row, map, 'STATUS'));
    var outlet = gttActNorm_(gttActCell_(row, map, 'OUTLET'));
    var jabatan = gttActNorm_(gttActCell_(row, map, 'JABATAN'));
    var saId = String(gttActCell_(row, map, 'SA_ID') || '').trim();

    if (status !== 'AKTIF') continue;
    if (outlet !== gttActNorm_(user.outlet)) continue;
    if (!gttActIsApproverPosition_(jabatan)) continue;
    if (!saId) continue;
    if (saId === user.saId) continue; // no self approval

    result.push({
      saId: saId,
      nama: String(gttActCell_(row, map, 'NAMA SA') || '').trim(),
      jabatan: jabatan,
      outlet: outlet
    });
  }

  return result;
}

function gttActGetValidApprover_(ss, user, approverIdInput) {
  var approverId = String(approverIdInput || '').trim();
  if (!approverId) throw new Error('Pilih SL/SPV/Manager untuk persetujuan.');
  var list = gttActGetApprovers_(ss, user);
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].saId) === approverId) return list[i];
  }
  throw new Error('Approver tidak valid, tidak aktif, berbeda outlet, atau merupakan akun Anda sendiri.');
}

function gttActGetUserByPin_(ss, pinInput) {
  var pin = String(pinInput || '').trim();
  if (!pin) throw new Error('Sesi login tidak valid. Silakan login ulang.');

  var sheet = ss.getSheetByName(GTT_ACTIVITY_BETA1.SHEET.MASTER_SA);
  if (!sheet) throw new Error('Sheet MASTER_SA tidak ditemukan.');
  var values = sheet.getDataRange().getDisplayValues();
  if (!values.length) throw new Error('MASTER_SA kosong.');
  var map = gttActHeaderMap_(values[0]);

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (String(gttActCell_(row, map, 'PIN') || '').trim() !== pin) continue;
    if (gttActNorm_(gttActCell_(row, map, 'STATUS')) !== 'AKTIF') {
      throw new Error('Akun SA tidak aktif.');
    }

    return {
      rowNumber: i + 1,
      pin: pin,
      nama: String(gttActCell_(row, map, 'NAMA SA') || '').trim(),
      outlet: gttActNorm_(gttActCell_(row, map, 'OUTLET')),
      saId: String(gttActCell_(row, map, 'SA_ID') || '').trim(),
      jabatan: gttActNorm_(gttActCell_(row, map, 'JABATAN')),
      role: gttActNorm_(gttActCell_(row, map, 'ROLE'))
    };
  }

  throw new Error('Pengguna tidak ditemukan pada MASTER_SA.');
}

function gttActIsApproverPosition_(jabatan) {
  return GTT_ACTIVITY_BETA1.APPROVER_POSITIONS.indexOf(gttActNorm_(jabatan)) !== -1;
}

function gttActLocationMap_(ss) {
  var sheet = ss.getSheetByName(GTT_ACTIVITY_BETA1.SHEET.MASTER_LOKASI);
  if (!sheet) throw new Error('Sheet MASTER_LOKASI tidak ditemukan.');
  var values = sheet.getDataRange().getDisplayValues();
  var result = {};
  if (!values.length) return result;
  var map = gttActHeaderMap_(values[0]);

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var code = gttActNorm_(gttActCell_(row, map, 'KODE'));
    if (!code) continue;
    result[code] = {
      code: code,
      name: String(gttActCell_(row, map, 'NAMA OUTLET') || code).trim(),
      latitude: Number(gttActCell_(row, map, 'LATITUDE')),
      longitude: Number(gttActCell_(row, map, 'LONGITUDE')),
      radius: Number(gttActCell_(row, map, 'RADIUS ABSEN (M)')) || 100,
      active: gttActBool_(gttActCell_(row, map, 'AKTIF'))
    };
  }
  return result;
}


/* ============================================================
 * INTERNAL — GEOFENCE / BREAK CONFLICT
 * ============================================================ */

function gttActValidateOriginGeofence_(ss, originOutlet, locationInput) {
  var map = gttActLocationMap_(ss);
  var origin = map[gttActNorm_(originOutlet)];
  if (!origin || !origin.active) throw new Error('Geofence outlet asal tidak tersedia/aktif.');

  var location = locationInput || {};
  var latitude = Number(location.latitude);
  var longitude = Number(location.longitude);
  var accuracy = Number(location.accuracy);
  if (!isFinite(latitude) || !isFinite(longitude) || !isFinite(accuracy)) {
    throw new Error('Data GPS tidak valid.');
  }
  if (accuracy > GTT_ACTIVITY_BETA1.MAX_GPS_ACCURACY_M) {
    throw new Error('Akurasi GPS ±' + Math.round(accuracy) + ' m belum memadai. Perbarui GPS lalu coba lagi.');
  }

  var distance = gttActDistanceMeters_(latitude, longitude, origin.latitude, origin.longitude);
  if (distance > origin.radius) {
    throw new Error(
      'Anda berada di luar geofence ' + origin.code + '. Jarak ±' +
      Math.round(distance) + ' m; batas ' + Math.round(origin.radius) + ' m.'
    );
  }

  return {
    outlet: origin.code,
    latitude: latitude,
    longitude: longitude,
    accuracy: accuracy,
    radiusMeters: origin.radius,
    distanceMeters: distance
  };
}

function gttActDistanceMeters_(lat1, lon1, lat2, lon2) {
  var rad = Math.PI / 180;
  var dLat = (lat2 - lat1) * rad;
  var dLon = (lon2 - lon1) * rad;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * rad) * Math.cos(lat2 * rad) *
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return 6371000 * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}

function gttActFindActiveBreak_(ss, pinInput, outletInput, now) {
  var sheet = ss.getSheetByName(GTT_ACTIVITY_BETA1.SHEET.LOG_ISTIRAHAT);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var values = sheet.getDataRange().getDisplayValues();
  var map = gttActHeaderMap_(values[0]);
  var dateKey = gttActDateKey_(now || new Date());
  var pin = String(pinInput || '').trim();
  var outlet = gttActNorm_(outletInput);

  for (var i = values.length - 1; i >= 1; i--) {
    var row = values[i];
    if (String(gttActCell_(row, map, 'PIN') || '').trim() !== pin) continue;
    if (gttActNorm_(gttActCell_(row, map, 'OUTLET')) !== outlet) continue;
    if (gttActDateKey_(gttActCell_(row, map, 'TANGGAL')) !== dateKey) continue;

    var s2Start = String(gttActCell_(row, map, 'S2 MULAI') || '').trim();
    var s2End = String(gttActCell_(row, map, 'S2 SELESAI') || '').trim();
    if (s2Start && !s2End) return { session: 2, label: 'Break 2', start: s2Start };

    var s1Start = String(gttActCell_(row, map, 'S1 MULAI') || '').trim();
    var s1End = String(gttActCell_(row, map, 'S1 SELESAI') || '').trim();
    if (s1Start && !s1End) return { session: 1, label: 'Break 1', start: s1Start };
  }
  return null;
}

function gttActGetTeamBreaks_(ss, outletInput, now) {
  var result = { break1: [], break2: [] };
  var sheet = ss.getSheetByName(GTT_ACTIVITY_BETA1.SHEET.LOG_ISTIRAHAT);
  if (!sheet || sheet.getLastRow() < 2) return result;

  var values = sheet.getDataRange().getDisplayValues();
  var map = gttActHeaderMap_(values[0]);
  var dateKey = gttActDateKey_(now);
  var outlet = gttActNorm_(outletInput);
  var seen = {};

  for (var i = values.length - 1; i >= 1; i--) {
    var row = values[i];
    var pin = String(gttActCell_(row, map, 'PIN') || '').trim();
    if (!pin || seen[pin]) continue;
    if (gttActNorm_(gttActCell_(row, map, 'OUTLET')) !== outlet) continue;
    if (gttActDateKey_(gttActCell_(row, map, 'TANGGAL')) !== dateKey) continue;

    var name = String(gttActCell_(row, map, 'NAMA SA') || '-').trim();
    var s2Start = String(gttActCell_(row, map, 'S2 MULAI') || '').trim();
    var s2End = String(gttActCell_(row, map, 'S2 SELESAI') || '').trim();
    var s1Start = String(gttActCell_(row, map, 'S1 MULAI') || '').trim();
    var s1End = String(gttActCell_(row, map, 'S1 SELESAI') || '').trim();

    if (s2Start && !s2End) {
      result.break2.push({
        pin: pin,
        nama: name,
        startTime: s2Start,
        elapsedSeconds: gttActElapsedFromClock_(s2Start, now)
      });
      seen[pin] = true;
      continue;
    }

    if (s1Start && !s1End) {
      result.break1.push({
        pin: pin,
        nama: name,
        startTime: s1Start,
        elapsedSeconds: gttActElapsedFromClock_(s1Start, now)
      });
      seen[pin] = true;
    }
  }

  result.break1.sort(function (a, b) { return a.nama.localeCompare(b.nama); });
  result.break2.sort(function (a, b) { return a.nama.localeCompare(b.nama); });
  return result;
}


/* ============================================================
 * INTERNAL — TASK SHEET
 * ============================================================ */

function gttActTaskContext_(ss, ensureHeaders) {
  var sheet = ss.getSheetByName(GTT_ACTIVITY_BETA1.SHEET.TUGAS_LUAR);
  if (!sheet) throw new Error('Sheet TUGAS_LUAR tidak ditemukan.');

  if (ensureHeaders) gttActEnsureTaskHeaders_(sheet);
  var lastColumn = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  var ctx = {
    sheet: sheet,
    headers: headers,
    map: gttActHeaderMap_(headers),
    lastColumn: lastColumn
  };

  // Rev-1: pastikan kolom Activity baru tidak mewarisi format tanggal dari kolom lain.
  // Ini mencegah TASK_NO_HARI tampil seperti 31/12/1899 dan menjaga nilai numerik stabil.
  if (ensureHeaders) gttActEnsureActivityColumnFormats_(ctx);

  return ctx;
}

function gttActEnsureTaskHeaders_(sheet) {
  var lastColumn = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  var map = gttActHeaderMap_(headers);
  var missing = [];

  for (var i = 0; i < GTT_ACTIVITY_BETA1.EXTRA_HEADERS.length; i++) {
    var h = GTT_ACTIVITY_BETA1.EXTRA_HEADERS[i];
    if (!Object.prototype.hasOwnProperty.call(map, gttActNormHeader_(h))) missing.push(h);
  }
  if (!missing.length) return;

  var requiredLast = lastColumn + missing.length;
  if (sheet.getMaxColumns() < requiredLast) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredLast - sheet.getMaxColumns());
  }
  sheet.getRange(1, lastColumn + 1, 1, missing.length).setValues([missing]);
}

function gttActEnsureActivityColumnFormats_(ctx) {
  var sheet = ctx.sheet;
  var rowCount = Math.max(1, sheet.getMaxRows() - 1);
  var numberHeaders = ['TASK_NO_HARI', 'START_TS_MS', 'END_TS_MS', 'DURATION_SECONDS'];
  var textHeaders = ['RUTE_JSON', 'RUTE_LABEL', 'APPROVAL_HISTORY_JSON', 'SOURCE_VERSION'];

  numberHeaders.forEach(function (header) {
    var idx = gttActHeaderIndex_(ctx.map, header);
    if (idx >= 0) sheet.getRange(2, idx + 1, rowCount, 1).setNumberFormat('0');
  });

  textHeaders.forEach(function (header) {
    var idx = gttActHeaderIndex_(ctx.map, header);
    if (idx >= 0) sheet.getRange(2, idx + 1, rowCount, 1).setNumberFormat('@');
  });
}

function gttActReadTasks_(ctx) {
  var lastRow = ctx.sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = ctx.sheet.getRange(2, 1, lastRow - 1, ctx.lastColumn).getValues();
  var display = ctx.sheet.getRange(2, 1, lastRow - 1, ctx.lastColumn).getDisplayValues();
  var result = [];

  for (var i = 0; i < values.length; i++) {
    var raw = values[i];
    var shown = display[i];
    var id = String(gttActCell_(shown, ctx.map, 'ID TUGAS') || '').trim();
    if (!id) continue;
    result.push(gttActRowToTask_(ctx, raw, shown, i + 2));
  }
  return result;
}

function gttActRowToTask_(ctx, raw, shown, rowNumber) {
  var routeJson = String(gttActCell_(shown, ctx.map, 'RUTE_JSON') || '').trim();
  var route = [];
  if (routeJson) {
    try {
      var parsed = JSON.parse(routeJson);
      route = Array.isArray(parsed) ? parsed : [];
    } catch (ignore) {
      route = [];
    }
  }
  if (!route.length) {
    var legacyTarget = String(gttActCell_(shown, ctx.map, 'TUJUAN') || '').trim();
    if (legacyTarget) route = [{ order: 1, type: 'OTHER', code: 'LEGACY', name: legacyTarget, display: legacyTarget }];
  }

  var dateRaw = gttActCell_(raw, ctx.map, 'TANGGAL');
  var status = gttActNorm_(gttActCell_(shown, ctx.map, 'STATUS'));
  var startTs = gttActFiniteNumber_(gttActCell_(raw, ctx.map, 'START_TS_MS'), gttActCell_(shown, ctx.map, 'START_TS_MS'));
  var endTs = gttActFiniteNumber_(gttActCell_(raw, ctx.map, 'END_TS_MS'), gttActCell_(shown, ctx.map, 'END_TS_MS'));
  var durationSeconds = gttActFiniteNumber_(gttActCell_(raw, ctx.map, 'DURATION_SECONDS'), gttActCell_(shown, ctx.map, 'DURATION_SECONDS'));
  var durationMinutes = gttActFiniteNumber_(gttActCell_(raw, ctx.map, 'DURASI (MENIT)'), gttActCell_(shown, ctx.map, 'DURASI (MENIT)'));

  return {
    rowNumber: rowNumber,
    id: String(gttActCell_(shown, ctx.map, 'ID TUGAS') || '').trim(),
    pin: String(gttActCell_(shown, ctx.map, 'PIN') || '').trim(),
    nama: String(gttActCell_(shown, ctx.map, 'NAMA SA') || '').trim(),
    saId: String(gttActCell_(shown, ctx.map, 'SA_ID') || '').trim(),
    outlet: gttActNorm_(gttActCell_(shown, ctx.map, 'OUTLET ASAL')),
    dateKey: gttActDateKey_(dateRaw || gttActCell_(shown, ctx.map, 'TANGGAL')),
    taskNo: gttActTaskNoValue_(gttActCell_(raw, ctx.map, 'TASK_NO_HARI'), gttActCell_(shown, ctx.map, 'TASK_NO_HARI')),
    route: route,
    routeLabel: String(gttActCell_(shown, ctx.map, 'RUTE_LABEL') || gttActCell_(shown, ctx.map, 'TUJUAN') || '').trim(),
    purpose: String(gttActCell_(shown, ctx.map, 'KETERANGAN') || '').trim(),
    status: status,
    approvalStatus: gttActNorm_(gttActCell_(shown, ctx.map, 'APPROVAL_STATUS')),
    approverId: String(gttActCell_(shown, ctx.map, 'APPROVER_ID') || '').trim(),
    approverName: String(gttActCell_(shown, ctx.map, 'APPROVER_NAMA') || '').trim(),
    approverPosition: gttActNorm_(gttActCell_(shown, ctx.map, 'APPROVER_JABATAN')),
    requestedAt: gttActCell_(raw, ctx.map, 'REQUESTED_AT'),
    approvedAt: gttActCell_(raw, ctx.map, 'APPROVED_AT'),
    rejectedAt: gttActCell_(raw, ctx.map, 'REJECTED_AT'),
    rejectionReason: String(gttActCell_(shown, ctx.map, 'REJECTION_REASON') || '').trim(),
    startTimeText: gttActTimeText_(gttActCell_(raw, ctx.map, 'JAM MULAI'), gttActCell_(shown, ctx.map, 'JAM MULAI')),
    endTimeText: gttActTimeText_(gttActCell_(raw, ctx.map, 'JAM KEMBALI'), gttActCell_(shown, ctx.map, 'JAM KEMBALI')),
    startTsMs: startTs,
    endTsMs: endTs,
    durationSeconds: durationSeconds,
    durationMinutes: durationMinutes,
    approvalHistoryJson: String(gttActCell_(shown, ctx.map, 'APPROVAL_HISTORY_JSON') || '[]'),
    cancelReason: String(gttActCell_(shown, ctx.map, 'CANCEL_REASON') || gttActCell_(shown, ctx.map, 'ALASAN BATAL') || '').trim()
  };
}

function gttActAppendRowObject_(ctx, obj) {
  var row = new Array(ctx.lastColumn);
  for (var i = 0; i < row.length; i++) row[i] = '';
  Object.keys(obj || {}).forEach(function (key) {
    var index = gttActHeaderIndex_(ctx.map, key);
    if (index >= 0) row[index] = obj[key];
  });
  var targetRow = Math.max(2, ctx.sheet.getLastRow() + 1);
  ctx.sheet.getRange(targetRow, 1, 1, ctx.lastColumn).setValues([row]);
  gttActFormatTaskRow_(ctx, targetRow);
}

function gttActUpdateTask_(ctx, rowNumber, updates) {
  var keys = Object.keys(updates || {});
  for (var i = 0; i < keys.length; i++) {
    var index = gttActHeaderIndex_(ctx.map, keys[i]);
    if (index < 0) continue;
    ctx.sheet.getRange(rowNumber, index + 1).setValue(updates[keys[i]]);
  }
  gttActFormatTaskRow_(ctx, rowNumber);
}

function gttActFormatTaskRow_(ctx, rowNumber) {
  var dateHeaders = [
    'TANGGAL', 'JAM MULAI', 'JAM KEMBALI', 'DIBUAT PADA', 'DIPERBARUI PADA',
    'REQUESTED_AT', 'APPROVED_AT', 'REJECTED_AT', 'CANCELLED_AT'
  ];
  for (var i = 0; i < dateHeaders.length; i++) {
    var idx = gttActHeaderIndex_(ctx.map, dateHeaders[i]);
    if (idx < 0) continue;
    var fmt = dateHeaders[i] === 'TANGGAL' ? 'dd/MM/yyyy' :
      (dateHeaders[i] === 'JAM MULAI' || dateHeaders[i] === 'JAM KEMBALI' ? 'HH:mm:ss' : 'dd/MM/yyyy HH:mm:ss');
    ctx.sheet.getRange(rowNumber, idx + 1).setNumberFormat(fmt);
  }
}

function gttActFindTaskById_(ctx, taskIdInput) {
  var taskId = String(taskIdInput || '').trim();
  if (!taskId) throw new Error('ID Tugas Luar tidak valid.');
  var tasks = gttActReadTasks_(ctx);
  for (var i = 0; i < tasks.length; i++) {
    if (tasks[i].id === taskId) return tasks[i];
  }
  throw new Error('Tugas Luar tidak ditemukan.');
}

function gttActAssertOwner_(task, user) {
  if (!task || task.pin !== user.pin || task.outlet !== user.outlet) {
    throw new Error('Anda tidak memiliki akses ke Tugas Luar ini.');
  }
}

function gttActAssertApproverCanAct_(task, user) {
  if (!task) throw new Error('Permintaan tidak ditemukan.');
  if (task.outlet !== user.outlet) throw new Error('Permintaan berasal dari outlet lain.');
  if (!gttActIsApproverPosition_(user.jabatan)) throw new Error('Akun ini bukan SL/SPV/Manager.');
  if (String(task.approverId || '') !== String(user.saId || '')) {
    throw new Error('Permintaan ini ditujukan kepada ' + (task.approverName || 'approver lain') + '.');
  }
}

function gttActAssertNoOpenTask_(ctx, user, message) {
  var tasks = gttActReadTasks_(ctx);
  for (var i = 0; i < tasks.length; i++) {
    if (tasks[i].pin !== user.pin) continue;
    if (
      tasks[i].status === GTT_ACTIVITY_BETA1.STATUS.PENDING ||
      tasks[i].status === GTT_ACTIVITY_BETA1.STATUS.APPROVED ||
      tasks[i].status === GTT_ACTIVITY_BETA1.STATUS.IN_PROGRESS
    ) {
      throw new Error(message || 'Masih ada Tugas Luar yang belum selesai.');
    }
  }
}

function gttActAssertNoOtherInProgress_(ctx, user, currentTaskId) {
  var tasks = gttActReadTasks_(ctx);
  for (var i = 0; i < tasks.length; i++) {
    if (tasks[i].id === currentTaskId) continue;
    if (tasks[i].pin === user.pin && tasks[i].status === GTT_ACTIVITY_BETA1.STATUS.IN_PROGRESS) {
      throw new Error('Anda masih memiliki Tugas Luar yang sedang berlangsung.');
    }
  }
}

function gttActNextTaskNo_(ctx, pinInput, dateKey) {
  var pin = String(pinInput || '').trim();
  var maxNo = 0;
  var tasks = gttActReadTasks_(ctx);
  for (var i = 0; i < tasks.length; i++) {
    if (tasks[i].pin !== pin || tasks[i].dateKey !== dateKey) continue;
    maxNo = Math.max(maxNo, Number(tasks[i].taskNo || 0));
  }
  return maxNo + 1;
}

function gttActApprovalHistory_(task) {
  try {
    var value = JSON.parse(String(task.approvalHistoryJson || '[]'));
    return Array.isArray(value) ? value : [];
  } catch (ignore) {
    return [];
  }
}


function gttActFiniteNumber_(rawValue, displayValue) {
  var candidates = [rawValue, displayValue];
  for (var i = 0; i < candidates.length; i++) {
    var value = candidates[i];
    if (value === '' || value === null || value === undefined) continue;
    if (Object.prototype.toString.call(value) === '[object Date]') {
      var ms = value.getTime();
      if (isFinite(ms)) return ms;
      continue;
    }
    var num = Number(value);
    if (isFinite(num)) return num;
  }
  return 0;
}

function gttActTaskNoValue_(rawValue, displayValue) {
  // Jika kolom sebelumnya terformat sebagai tanggal, Sheets dapat mengembalikan serial
  // tugas (1,2,3...) sebagai Date 1899/1900. Konversikan kembali ke serial hari.
  if (Object.prototype.toString.call(rawValue) === '[object Date]' && !isNaN(rawValue.getTime())) {
    var epoch = Date.UTC(1899, 11, 30);
    var serial = Math.round((Date.UTC(
      rawValue.getFullYear(), rawValue.getMonth(), rawValue.getDate()
    ) - epoch) / 86400000);
    return serial > 0 && serial < 10000 ? serial : 0;
  }

  var rawNum = Number(rawValue);
  if (isFinite(rawNum) && rawNum > 0 && rawNum < 10000) return Math.round(rawNum);

  var text = String(displayValue === null || displayValue === undefined ? '' : displayValue).trim();
  if (/^\d{1,4}$/.test(text)) return Number(text);
  return 0;
}

/* ============================================================
 * INTERNAL — GENERIC HELPERS
 * ============================================================ */

function gttActSpreadsheet_() {
  return SpreadsheetApp.openById(GTT_ACTIVITY_BETA1.SPREADSHEET_ID);
}

function gttActWithLock_(callback) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Sistem sedang memproses transaksi lain. Coba lagi beberapa detik.');
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function gttActSafeApi_(operation, callback) {
  // Rev-1: seluruh response Activity dikirim sebagai JSON string.
  // google.script.run paling stabil untuk payload besar/campuran bila transport-nya
  // benar-benar serializable; frontend Rev-1 akan menormalisasi string ini kembali ke object.
  try {
    var result = callback();
    return JSON.stringify(gttActJsonSafe_(result));
  } catch (error) {
    console.error('[GTT ACTIVITY ' + operation + ']', error);
    return JSON.stringify({
      success: false,
      code: 'GTT_ACT_' + operation + '_FAILED',
      message: error && error.message ? error.message : 'Proses GTT Activity gagal.'
    });
  }
}

function gttActJsonSafe_(value) {
  if (value === null || value === undefined) return null;

  var type = typeof value;
  if (type === 'string' || type === 'boolean') return value;
  if (type === 'number') return isFinite(value) ? value : 0;

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return isNaN(value.getTime()) ? '' : value.toISOString();
  }

  if (Array.isArray(value)) {
    return value.map(function (item) { return gttActJsonSafe_(item); });
  }

  if (type === 'object') {
    var result = {};
    Object.keys(value).forEach(function (key) {
      var child = value[key];
      if (typeof child === 'function' || typeof child === 'undefined') return;
      result[key] = gttActJsonSafe_(child);
    });
    return result;
  }

  return String(value);
}

function gttActHeaderMap_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var key = gttActNormHeader_(headers[i]);
    if (key) map[key] = i;
  }
  return map;
}

function gttActHeaderIndex_(map, header) {
  var key = gttActNormHeader_(header);
  return Object.prototype.hasOwnProperty.call(map, key) ? map[key] : -1;
}

function gttActCell_(row, map, header) {
  var idx = gttActHeaderIndex_(map, header);
  return idx >= 0 ? row[idx] : '';
}

function gttActNormHeader_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function gttActNorm_(value) {
  return String(value === null || value === undefined ? '' : value).trim().toUpperCase();
}

function gttActBool_(value) {
  var text = gttActNorm_(value);
  return value === true || text === 'TRUE' || text === 'YA' || text === 'AKTIF' || text === '1';
}

function gttActNewTaskId_(now, saId) {
  var stamp = Utilities.formatDate(now, GTT_ACTIVITY_BETA1.TIME_ZONE, 'yyyyMMddHHmmss');
  var shortUuid = Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
  return 'TL-' + stamp + '-' + String(saId || 'SA').replace(/[^A-Za-z0-9]/g, '').toUpperCase() + '-' + shortUuid;
}

function gttActDateLabelId_(value) {
  var date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return '';

  var key = Utilities.formatDate(date, GTT_ACTIVITY_BETA1.TIME_ZONE, 'yyyy-MM-dd').split('-');
  var year = Number(key[0]);
  var month = Number(key[1]);
  var day = Number(key[2]);
  var dayIndex = new Date(Date.UTC(year, month - 1, day)).getUTCDay();
  var days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  var months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

  return days[dayIndex] + ', ' + String(day).padStart(2, '0') + ' ' + months[month - 1] + ' ' + year;
}

function gttActDateKey_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, GTT_ACTIVITY_BETA1.TIME_ZONE, 'yyyy-MM-dd');
  }
  var text = String(value || '').trim();
  if (!text) return '';

  var iso = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (iso) return iso[1] + '-' + gttActPad2_(iso[2]) + '-' + gttActPad2_(iso[3]);

  var id = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (id) return id[3] + '-' + gttActPad2_(id[2]) + '-' + gttActPad2_(id[1]);

  var parsed = new Date(text);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, GTT_ACTIVITY_BETA1.TIME_ZONE, 'yyyy-MM-dd');
  return text;
}

function gttActTimeText_(rawValue, displayValue) {
  if (Object.prototype.toString.call(rawValue) === '[object Date]' && !isNaN(rawValue.getTime())) {
    return Utilities.formatDate(rawValue, GTT_ACTIVITY_BETA1.TIME_ZONE, 'HH:mm:ss');
  }
  var text = String(displayValue || rawValue || '').trim();
  var match = text.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!match) return text;
  return gttActPad2_(match[1]) + ':' + match[2] + ':' + gttActPad2_(match[3] || '00');
}

function gttActFormatDateTime_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, GTT_ACTIVITY_BETA1.TIME_ZONE, 'dd/MM/yyyy HH:mm:ss');
  }
  return String(value).trim();
}

function gttActElapsedSeconds_(task, now) {
  if (task.status === GTT_ACTIVITY_BETA1.STATUS.COMPLETED && Number(task.durationSeconds || 0) >= 0) {
    return Number(task.durationSeconds || 0);
  }
  var start = Number(task.startTsMs || 0);
  if (!start) return 0;
  return Math.max(0, Math.floor(((now || new Date()).getTime() - start) / 1000));
}

function gttActElapsedFromClock_(startText, now) {
  var match = String(startText || '').match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!match) return 0;
  var startSec = Number(match[1]) * 3600 + Number(match[2]) * 60 + Number(match[3] || 0);
  var currentText = Utilities.formatDate(now || new Date(), GTT_ACTIVITY_BETA1.TIME_ZONE, 'HH:mm:ss');
  var current = currentText.split(':');
  var nowSec = Number(current[0]) * 3600 + Number(current[1]) * 60 + Number(current[2]);
  if (nowSec < startSec) nowSec += 86400;
  return Math.max(0, nowSec - startSec);
}

function gttActPad2_(value) {
  var text = String(value === null || value === undefined ? '' : value);
  return text.length < 2 ? '0' + text : text;
}
