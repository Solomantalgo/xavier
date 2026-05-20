// ============================================================
//  UNION BIKE LOAN — Apps Script Backend for Vercel
//  Deploy as Web App:
//    Execute as: Me
//    Who has access: Anyone
//  After any code change → create a NEW deployment version
// ============================================================

var SPREADSHEET_ID = '1TIL1wOSyIR0AJwQpkPYwcBWFmPdm9ky0izkcahCYZD8';
var SHEET_NAME     = 'Applications';

function doPost(e) {
  try {
    var data;
    // Handle both raw JSON and form data (in case 'no-cors' is used or standard POST)
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      // Fallback
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

    var row = buildRow(data);
    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ success: true, message: 'Application submitted successfully!' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// To allow preflight requests for CORS (Google handles it somewhat weirdly, but this helps in some setups)
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
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
    'Agent Name', 'Agent Contact'
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
    d.agentContact || ''
  ];
}
