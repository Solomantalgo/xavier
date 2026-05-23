// ============================================================
//  UNION BIKE LOAN — Apps Script Backend for Vercel
//  Deploy as Web App:
//    Execute as: Me
//    Who has access: Anyone
//  After any code change → create a NEW deployment version
// ============================================================

var SPREADSHEET_ID  = '1TIL1wOSyIR0AJwQpkPYwcBWFmPdm9ky0izkcahCYZD8';
var SHEET_NAME      = 'Applications';
var DRIVE_FOLDER_ID = '1ra70Gg-JFRCbduPluqWR25xPqr2o3Tb1';

// Save a base64-encoded image to Google Drive and return a public URL
function saveImageToDrive(base64Data, filename) {
  if (!base64Data || base64Data.length < 10) return '';
  try {
    // Strip the data URI prefix if present (e.g. "data:image/jpeg;base64,")
    var parts = base64Data.split(',');
    var raw = parts.length > 1 ? parts[1] : parts[0];
    var mimeMatch = base64Data.match(/data:([^;]+);/);
    var mimeType = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    var ext = mimeType.split('/')[1] || 'jpg';
    var fullFilename = filename + '.' + ext;

    var decoded = Utilities.base64Decode(raw);
    var blob = Utilities.newBlob(decoded, mimeType, fullFilename);

    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Return a direct-view URL
    return 'https://drive.google.com/file/d/' + file.getId() + '/view';
  } catch (err) {
    Logger.log('Image upload error: ' + err.message);
    return '';
  }
}

function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

function generateFormId() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var id = 'UB-';
  for(var i=0; i<6; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return id;
}

function formatDate(date) {
  var y = date.getFullYear();
  var m = String(date.getMonth() + 1).padStart(2, '0');
  var d = String(date.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + d;
}

function doGet(e) {
  try {
    if (e.parameter.action !== 'getApp') {
      return ContentService.createTextOutput("GET endpoint ready.").setMimeType(ContentService.MimeType.TEXT);
    }

    var formId = e.parameter.formId;
    var nin = e.parameter.nin;
    var dob = e.parameter.dob;

    if (!formId && !(nin && dob)) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Provide Form ID or NIN & DOB' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet not found');

    var data = sheet.getDataRange().getValues();
    var foundRow = null;

    // Reverse loop to get the most recent submission if NIN/DOB matches multiple
    for (var i = data.length - 1; i > 0; i--) {
      var row = data[i];
      if (formId) {
        if (row[57] === formId) {
          foundRow = row;
          break;
        }
      } else if (nin && dob) {
        var rowNin = String(row[10]).trim();
        var rowDob = String(row[9]).trim();
        if (row[9] instanceof Date) {
          rowDob = formatDate(row[9]);
        }
        if (rowNin.toLowerCase() === String(nin).toLowerCase().trim() && rowDob === String(dob).trim()) {
          foundRow = row;
          break;
        }
      }
    }

    if (!foundRow) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Application not found.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var result = {
      formType: foundRow[1], applicationDate: foundRow[2] instanceof Date ? formatDate(foundRow[2]) : foundRow[2],
      unionBranch: foundRow[3], surname: foundRow[4], otherNames: foundRow[5], village: foundRow[6],
      subCounty: foundRow[7], district: foundRow[8], dob: foundRow[9] instanceof Date ? formatDate(foundRow[9]) : foundRow[9],
      nin: foundRow[10], gender: foundRow[11], phone1: foundRow[12], phone2: foundRow[13], maritalStatus: foundRow[14],
      residence: foundRow[15], bikeType: foundRow[16], exchangePlate: foundRow[17], exchangeLogbook: foundRow[18],
      exchangeBikeType: foundRow[19], workName: foundRow[20], workLocation: foundRow[21], chairmanName: foundRow[22],
      chairmanContact: foundRow[23], incomeSource1: foundRow[24], dailyIncome1: foundRow[25], incomeSource2: foundRow[26],
      dailyIncome2: foundRow[27], totalIncome: foundRow[28], expense1: foundRow[29], dailyCost1: foundRow[30],
      expense2: foundRow[31], dailyCost2: foundRow[32], totalExpense: foundRow[33], ridingPermit: foundRow[34],
      guarantorType: foundRow[35], g1Name: foundRow[36], g1Contact: foundRow[37], g1NIN: foundRow[38],
      g1LoanId: foundRow[39], g1BikePlate: foundRow[40], g1IncomeSource: foundRow[41], g1Location: foundRow[42],
      g2Name: foundRow[43], g2Contact: foundRow[44], g2NIN: foundRow[45], g2IncomeSource: foundRow[46],
      g2Location: foundRow[47], tin: foundRow[48], g1GuarantorID: foundRow[49], g2GuarantorID: foundRow[50],
      agentName: foundRow[51], agentContact: foundRow[52], bmName: foundRow[53], bmContact: foundRow[54],
      bmSigDate: foundRow[55] instanceof Date ? formatDate(foundRow[55]) : foundRow[55], bmBranch: foundRow[56],
      formId: foundRow[57]
    };

    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      data = JSON.parse(e.parameter.data);
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      var headers = getHeaders();
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setBackground('#1a5c45')
           .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    var isUpdate = false;
    var rowToUpdate = -1;
    var formId = data.formId || '';

    if (formId) {
       var dataRange = sheet.getDataRange().getValues();
       for(var i = dataRange.length - 1; i > 0; i--) {
         if (dataRange[i][57] === formId) {
           rowToUpdate = i + 1;
           isUpdate = true;
           break;
         }
       }
    }
    
    if (!isUpdate) {
      formId = generateFormId();
      data.formId = formId;
    }

    // --- Upload images to Google Drive ---
    var namePrefix = (data.surname || 'applicant').replace(/[^a-zA-Z0-9]/g, '_') + '_' + formId;
    if (data.photoBase64 && !isUpdate) {
      data.photoUrl = saveImageToDrive(data.photoBase64, namePrefix + '_photo');
    } else if (data.photoBase64 && isUpdate) {
      data.photoUrl = saveImageToDrive(data.photoBase64, namePrefix + '_photo');
    }
    if (data.sigBase64) {
      data.sigUrl = saveImageToDrive(data.sigBase64, namePrefix + '_signature');
    }
    if (data.bmSigBase64) {
      data.bmSigUrl = saveImageToDrive(data.bmSigBase64, namePrefix + '_manager_sig');
    }

    var row = buildRow(data);

    if (isUpdate) {
      // Overwrite existing row
      sheet.getRange(rowToUpdate, 1, 1, row.length).setValues([row]);
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: 'Application updated successfully!', formId: formId }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      sheet.appendRow(row);
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: 'Application submitted successfully!', formId: formId }))
        .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------- Column headers ----------
function getHeaders() {
  return [
    'Timestamp', 'Form Type', 'Application Date', 'Union Branch',
    'Surname', 'Other Names', 'Village', 'Sub County', 'District',
    'Date of Birth', 'NIN', 'Gender', 'Phone 1', 'Phone 2',
    'Marital Status', 'Type of Residence',
    'Bike Applied For',
    'Exchange Bike Plate', 'Exchange Logbook Names', 'Exchange Bike Type',
    'Work/Stage Name', 'Work/Stage Location',
    'Stage Chairman Name', 'Chairman Contact',
    'Income Source 1', 'Daily Income 1 (UGX)',
    'Income Source 2', 'Daily Income 2 (UGX)',
    'Total Daily Income (UGX)',
    'Expense 1', 'Daily Cost 1 (UGX)',
    'Expense 2', 'Daily Cost 2 (UGX)',
    'Total Daily Expense (UGX)',
    'Has Riding Permit',
    'Guarantor Type',
    'G1 Name', 'G1 Contact', 'G1 NIN',
    'G1 Union Loan ID', 'G1 Union Bike Plate',
    'G1 Income Source', 'G1 Location',
    'G2 Name', 'G2 Contact', 'G2 NIN',
    'G2 Income Source', 'G2 Location',
    'TIN', 'G1 Guarantor ID', 'G2 Guarantor ID',
    'Agent Name', 'Agent Contact',
    'Branch Manager Name', 'Branch Manager Contact', 'Manager Sig Date', 'Manager Branch',
    'Form ID',
    'Photo URL', 'Signature URL', 'Manager Signature URL'
  ];
}

// ---------- Map form data to row ----------
function buildRow(d) {
  return [
    new Date(),
    d.formType || 'Union Bike Loan', d.applicationDate, d.unionBranch,
    d.surname, d.otherNames, d.village, d.subCounty, d.district,
    d.dob, d.nin, d.gender, d.phone1, d.phone2,
    d.maritalStatus, d.residence,
    d.bikeType || '',
    d.exchangePlate || '', d.exchangeLogbook || '', d.exchangeBikeType || '',
    d.workName, d.workLocation, d.chairmanName, d.chairmanContact,
    d.incomeSource1, d.dailyIncome1,
    d.incomeSource2, d.dailyIncome2,
    d.totalIncome,
    d.expense1, d.dailyCost1,
    d.expense2, d.dailyCost2,
    d.totalExpense,
    d.ridingPermit,
    d.guarantorType || '',
    d.g1Name, d.g1Contact, d.g1NIN,
    d.g1LoanId || '', d.g1BikePlate || '',
    d.g1IncomeSource, d.g1Location,
    d.g2Name, d.g2Contact, d.g2NIN,
    d.g2IncomeSource, d.g2Location,
    d.tin || '',
    d.g1GuarantorID || '',
    d.g2GuarantorID || '',
    d.agentName || '',
    d.agentContact || '',
    d.bmName || '',
    d.bmContact || '',
    d.bmSigDate || '',
    d.bmBranch || '',
    d.formId || '',
    d.photoUrl || '',
    d.sigUrl || '',
    d.bmSigUrl || ''
  ];
}
