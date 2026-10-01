/**
 * The Big Family Radio & Autonomous Scheduler
 * Hybrid Live Streams + Google Drive Audio Broadcaster
 */

const FOLDER_NAME = "The Big Family Radio";
const SHEET_NAME = "RadioScheduleDual";
const DEFAULT_HEADERS = ["startTime", "endTime", "title", "category", "language", "url", "sourceType", "fileId", "durationSec"];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Radio Scheduler")
    .addItem("Sync Drive Audios into Schedule", "syncDriveFilesToSchedule")
    .addItem("Reset Stations to Default", "forceResetSheet")
    .addToUi();
}

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  const folder = DriveApp.createFolder(FOLDER_NAME);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return folder;
}

function getSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let schedSheet = ss.getSheetByName(SHEET_NAME);
  if (!schedSheet) {
    schedSheet = ss.insertSheet(SHEET_NAME);
    forceResetSheet();
  }

  let tracksSheet = ss.getSheetByName("Tracks");
  if (!tracksSheet) {
    tracksSheet = ss.insertSheet("Tracks");
    tracksSheet.appendRow(["Seq", "Title", "Contributor", "FileId", "StreamUrl", "CreatedAt", "Category", "ApprovalStatus", "DurationSec"]);
  }

  let listenersSheet = ss.getSheetByName("Listeners");
  if (!listenersSheet) {
    listenersSheet = ss.insertSheet("Listeners");
    listenersSheet.appendRow(["ListenerId", "Username", "Status", "LastSeen"]);
  }

  return { schedSheet, tracksSheet, listenersSheet };
}

function forceResetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  sheet.clear();
  const headerRange = sheet.getRange(1, 1, 1, DEFAULT_HEADERS.length);
  headerRange.setValues([DEFAULT_HEADERS]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#1e293b");
  headerRange.setFontColor("#f8fafc");
  headerRange.setHorizontalAlignment("center");

  // Fallback initial feeds
  const initial = [
    ["05:00", "07:00", "Fun Kids UK (Live Radio)", "Kids", "English", "https://radio.canstream.co.uk:8021/live.mp3", "live", "", 0],
    ["07:00", "09:00", "BBC World Service", "News", "English", "https://stream.live.vc.bbcmedia.co.uk/bbc_world_service", "live", "", 0],
    ["09:00", "12:00", "Radio Mirchi Top 20", "Entertainment", "Hindi", "https://stream.zeno.fm/0r0xa792kwzuv", "live", "", 0],
    ["12:00", "15:00", "AIR Hyderabad Live", "Entertainment", "Telugu", "https://air.pc.cdn.bitgravity.com/air/live/pcradio/pub/airtelugu/playlist.m3u8", "live", "", 0]
  ];

  sheet.getRange(2, 1, initial.length, DEFAULT_HEADERS.length).setValues(initial);
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, 2).setNumberFormat("@");
  sheet.autoResizeColumns(1, DEFAULT_HEADERS.length);
}

// Automatically scans Drive folder & approved Tracks and maps them into timed schedule slots
function syncDriveFilesToSchedule() {
  const { schedSheet, tracksSheet } = getSheets();
  const folder = getOrCreateFolder();
  const files = folder.getFiles();

  const trackRows = tracksSheet.getDataRange().getValues();
  const registeredFileIds = new Set();
  for (let i = 1; i < trackRows.length; i++) {
    registeredFileIds.add(String(trackRows[i][3]).trim());
  }

  // Auto-register any audio file dropped straight into the Google Drive folder
  while (files.hasNext()) {
    const file = files.next();
    const id = file.getId();
    if (!registeredFileIds.has(id)) {
      const name = file.getName().replace(/\.[^/.]+$/, "");
      const nextSeq = tracksSheet.getLastRow();
      tracksSheet.appendRow([nextSeq, name, "Family Drive", id, file.getUrl(), new Date().toISOString(), "Drive Audio", "approved", 180]);
    }
  }

  // Read all approved drive audios
  const updatedTrackRows = tracksSheet.getDataRange().getValues();
  const driveAudios = [];
  for (let i = 1; i < updatedTrackRows.length; i++) {
    const fileId = String(updatedTrackRows[i][3]).trim();
    const status = String(updatedTrackRows[i][7]).trim();
    if (fileId && status === "approved") {
      driveAudios.push({
        title: String(updatedTrackRows[i][1]),
        fileId: fileId,
        duration: Math.round(Number(updatedTrackRows[i][8])) || 180
      });
    }
  }

  if (driveAudios.length === 0) return;

  // Append drive audios sequentially into the schedule sheet
  let startMinutes = 15 * 60; // Start at 15:00 (3 PM)
  const rowsToAppend = [];

  driveAudios.forEach(item => {
    const endMinutes = startMinutes + Math.ceil(item.duration / 60);

    const sH = String(Math.floor(startMinutes / 60) % 24).padStart(2, "0");
    const sM = String(startMinutes % 60).padStart(2, "0");
    const eH = String(Math.floor(endMinutes / 60) % 24).padStart(2, "0");
    const eM = String(endMinutes % 60).padStart(2, "0");

    rowsToAppend.push([
      `${sH}:${sM}`,
      `${eH}:${eM}`,
      item.title,
      "Drive Audio",
      "English",
      "", // URL empty, resolved via fileId proxy
      "drive",
      item.fileId,
      item.duration
    ]);

    startMinutes = endMinutes;
  });

  if (rowsToAppend.length > 0) {
    const startRow = schedSheet.getLastRow() + 1;
    schedSheet.getRange(startRow, 1, rowsToAppend.length, DEFAULT_HEADERS.length).setValues(rowsToAppend);
    schedSheet.getRange(startRow, 1, rowsToAppend.length, 2).setNumberFormat("@");
  }
}

function doGet(e) {
  try {
    const parameter = (e && e.parameter) ? e.parameter : {};
    const action = parameter.action || "getSchedule";

    // 1. Google Drive Audio Stream Proxy (Bypasses range/CORS limitations)
    if (action === "streamAudio") {
      const fileId = parameter.fileId;
      if (!fileId) throw new Error("Missing fileId");

      const file = DriveApp.getFileById(fileId);
      let mime = file.getMimeType() || "audio/mpeg";
      if (mime === "application/octet-stream") mime = "audio/mpeg";

      const bytes = file.getBlob().getBytes();
      const b64 = Utilities.base64Encode(bytes);
      const dataUri = "data:" + mime + ";base64," + b64;

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        dataUri: dataUri,
        mimeType: mime,
        sizeBytes: bytes.length
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Fetch Full Combined Timetable (Live + Drive Audios)
    if (action === "getSchedule") {
      const { schedSheet } = getSheets();
      const rows = schedSheet.getDataRange().getValues();

      if (rows.length < 2) {
        return ContentService.createTextOutput(JSON.stringify({ channels: [] }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const headers = rows[0].map(h => h.toString().trim());
      const data = [];

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row[0]) continue;

        let item = {};
        headers.forEach((header, colIndex) => {
          let val = row[colIndex];
          if (val instanceof Date) {
            const hours = String(val.getHours()).padStart(2, "0");
            const minutes = String(val.getMinutes()).padStart(2, "0");
            val = `${hours}:${minutes}`;
          }
          item[header] = val;
        });

        // Ensure default properties
        item.sourceType = item.sourceType || (item.fileId ? "drive" : "live");
        item.durationSec = Number(item.durationSec) || 0;
        data.push(item);
      }

      return ContentService.createTextOutput(JSON.stringify({ channels: data, serverTime: Date.now() }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Invalid GET" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error("No payload");

    const body = JSON.parse(e.postData.contents);
    const action = body.action || "saveSchedule";

    if (action === "uploadDriveAudio") {
      return ContentService.createTextOutput(JSON.stringify(handleAudioUpload(body)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "heartbeat") {
      return ContentService.createTextOutput(JSON.stringify(handleHeartbeat(body)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Save schedule modifications
    const { schedSheet } = getSheets();
    const newSchedule = body.channels;

    if (!Array.isArray(newSchedule)) throw new Error("Expected channels array");

    const lastRow = schedSheet.getLastRow();
    if (lastRow > 1) {
      schedSheet.getRange(2, 1, lastRow - 1, DEFAULT_HEADERS.length).clearContent();
    }

    const rowsToInsert = newSchedule.map(item => [
      String(item.startTime || "").trim(),
      String(item.endTime || "").trim(),
      String(item.title || "").trim(),
      String(item.category || "General").trim(),
      String(item.language || "English").trim(),
      String(item.url || "").trim(),
      String(item.sourceType || (item.fileId ? "drive" : "live")).trim(),
      String(item.fileId || "").trim(),
      Number(item.durationSec) || 0
    ]);

    if (rowsToInsert.length > 0) {
      schedSheet.getRange(2, 1, rowsToInsert.length, DEFAULT_HEADERS.length).setValues(rowsToInsert);
      schedSheet.getRange(2, 1, rowsToInsert.length, 2).setNumberFormat("@");
    }

    SpreadsheetApp.flush();
    return ContentService.createTextOutput(JSON.stringify({ status: "success", count: rowsToInsert.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleAudioUpload(data) {
  const title = String(data.title || "Family Audio").trim();
  const base64File = data.base64File;
  const fileName = data.fileName || `${title}.mp3`;
  const mimeType = data.mimeType || "audio/mpeg";
  const durationSec = Math.round(Number(data.durationSec)) || 180;

  if (!base64File) return { success: false, message: "Missing audio data" };

  const folder = getOrCreateFolder();
  const decodedData = Utilities.base64Decode(base64File);
  const blob = Utilities.newBlob(decodedData, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const { tracksSheet } = getSheets();
  const nextSeq = tracksSheet.getLastRow();

  tracksSheet.appendRow([nextSeq, title, "Family Upload", fileId, file.getUrl(), new Date().toISOString(), "Drive Audio", "approved", durationSec]);

  // Immediately integrate into the schedule
  syncDriveFilesToSchedule();

  return { success: true, message: "Audio uploaded and integrated into radio schedule!", fileId };
}

function handleHeartbeat(data) {
  const listenerId = String(data.listenerId || "L_guest").trim();
  const username = String(data.username || "Family Member").trim();
  const isPlaying = Boolean(data.isPlaying);

  const { listenersSheet } = getSheets();
  const rows = listenersSheet.getDataRange().getValues();
  const now = Date.now();
  let found = false;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === listenerId) {
      listenersSheet.getRange(i + 1, 2).setValue(username);
      listenersSheet.getRange(i + 1, 3).setValue(isPlaying ? "Listening Live" : "Idle");
      listenersSheet.getRange(i + 1, 4).setValue(now);
      found = true;
      break;
    }
  }

  if (!found) {
    listenersSheet.appendRow([listenerId, username, isPlaying ? "Listening Live" : "Idle", now]);
  }

  let activeCount = 0;
  const updated = listenersSheet.getDataRange().getValues();
  for (let i = 1; i < updated.length; i++) {
    const lastSeen = Number(updated[i][3]) || 0;
    if (now - lastSeen < 60000 && String(updated[i][2]) === "Listening Live") {
      activeCount++;
    }
  }

  return { success: true, activeCount };
}