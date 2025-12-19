function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Technical Support Reports')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getReports() {
  const ss = SpreadsheetApp.openById('1TWit7wuvT6T5Smmm5qIFMzuzc4UZleuLnauW5xDRBZY');
  const sheet = ss.getSheetByName('Sheet1');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  // Create case-insensitive header index
  const idx = {};
  headers.forEach((h, i) => {
    if (h) idx[h.toString().trim().toLowerCase()] = i;
  });

  const get = (row, name) => {
    const i = idx[name.toLowerCase()];
    return (i !== undefined && row[i] !== undefined && row[i] !== null) ? String(row[i]).trim() : '';
  };

  const reports = data.map(row => ({
    ticket: get(row, 'Submission ID'),
    submit: convertToMalaysiaTime(get(row, 'Submitted at')),
    outlet: get(row, 'Outlet'),
    description: get(row, 'Description of the issue'),
    image: get(row, 'Upload supporting image'),
    location: get(row, 'Specific area / Detailed location'),
    urgency: get(row, 'Urgency Level'),
    outOfService: get(row, 'Is this equipment fully out of service?'),
    status: get(row, 'Status'),
    etc: get(row, 'ETC'),            // 🟢 Added: from column W
    remarks: get(row, 'REMARKS'),    // ✅ corrected to column AB
    completionImage: row[26] || '',  // column AA
    followUp: get(row, 'Follow-up')  // ✅ added
  }));

  return reports;
}

function getOutlets() {
  const ss = SpreadsheetApp.openById('1TWit7wuvT6T5Smmm5qIFMzuzc4UZleuLnauW5xDRBZY');
  const sheet = ss.getSheetByName('Sheet1');
  const data = sheet.getRange('F2:F').getValues().flat().filter(String);
  return [...new Set(data)].sort();
}

function convertToMalaysiaTime(dateString) {
  if (!dateString) return '';

  let date;
  try {
    if (Object.prototype.toString.call(dateString) === '[object Date]') {
      date = dateString;
    } else {
      date = new Date(dateString.replace(/-/g, '/'));
    }
    if (isNaN(date.getTime())) throw new Error('Invalid date');
  } catch {
    return dateString;
  }

  const malaysiaTime = new Date(date.getTime() + 8 * 60 * 60 * 1000);
  const options = {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  };
  return malaysiaTime.toLocaleString('en-GB', options);
}

function staffRequestFollowUp(ticket) {
  const ss = SpreadsheetApp.openById('1TWit7wuvT6T5Smmm5qIFMzuzc4UZleuLnauW5xDRBZY');
  const sheet = ss.getSheetByName('Sheet1');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const ticketCol = headers.findIndex(h => h && h.toString().trim().toLowerCase() === 'submission id');
  const followUpCol = headers.findIndex(h => h && h.toString().trim().toLowerCase() === 'follow-up');

  if (ticketCol === -1 || followUpCol === -1) throw new Error("Required columns not found!");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][ticketCol]).trim() === ticket) {
      const now = new Date();
      sheet.getRange(i + 1, followUpCol + 1).setValue(now); // +1 for 1-based range
      sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()).setBackground('#fff7d1'); // highlight row
      return { success: true, timestamp: now };
    }
  }
  return { success: false };
}
