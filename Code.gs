function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Import CSV data 👉️")
    .addItem("Import from URL", "importCSVFromUrl")
    .addItem("Import from Drive", "importCSVFromDrive")
    .addToUi();
}


function importCSVFromUrl() {
  var url = promptUserForInput("Please enter the URL of the CSV file:");
  var contents = UrlFetchApp.fetch(url);
  displayToastAlert(contents);
}

function promptUserForInput(promptText) {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt(promptText);
  var response = prompt.getResponseText();
  return response;
}
function findFilesInDrive(filename) {
  var files = DriveApp.getFilesByName(filename);
  var result = [];
  while(files.hasNext())
    result.push(files.next());
  return result;
}

function getTextFromPDF() {
  //var fileName = promptUserForInput("Please enter the name of the CSV file to import from Google Drive:");
  var files = findFilesInDrive("document.pdf");
  if(files.length === 0) {
    displayToastAlert("No files with name \"" + fileName + "\" were found in Google Drive.");
    return;
  } else if(files.length > 1) {
    displayToastAlert("Multiple files with name " + fileName +" were found. This program does not support picking the right file yet.");
    return;
  }
  var file = files[0];
  blob = file.getBlob()
  var resource = {
    title: blob.getName(),
    mimeType: blob.getContentType()
  };
  var options = {
    ocr: true,
    ocrLanguage: "en"
  };
  // Convert the pdf to a Google Doc with ocr.
  var file = Drive.Files.insert(resource, blob, options);
  var doc = DocumentApp.openById(file.id);
  var text = doc.getBody().getText();
  var title = doc.getName();
  
  // Deleted the document once the text has been stored.
  Drive.Files.remove(doc.getId());
  
  return text
}



function here(){
  
  txt = getTextFromPDF()
  
  txt = txt.split("\n").join(' ')
  txt = txt.replace(/day\(s\) added to your holding period as a result of a wash sale./g, "")
  txt = txt.replace(/Page [0-9]+ of [0-9]+ .+? (Yield | Security Identifier)/g, "")
  txt = txt.replace(/(Yield|Unrealized|Gain|Loss|Estimated|Annual Income|Accrued Interest|Cost Basis Market Value)/g, "")

  fund = "NGG"
  const regexp = new RegExp(`${fund}[ ]+CUSIP: [a-zA-Z0-9_]{0,20} [\s\0-9/.,$%]*[Reinvestments to Date]*[\s\0-9/.,$%]*[Total Covered]*[\s\0-9/.,$%]*[Total]*[\s\0-9/.,$%]*`, 'g')
  numbers = txt.match(regexp)[0]
  console.log(numbers)
  numbers = numbers.replace(/[\s\a-zA-Z:]{2,}/g,"")
  console.log(numbers)
  numbers = numbers.trim().split(" ")
  count = 3 //third back
  if (numbers[numbers.length-1].includes("%")){ count = 4}
  console.log(numbers[numbers.length-count])
}

function extractTextFromPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  //Get all PDF files:
  const folder = DriveApp.getFolderById("1G-47eVNyxqI9_bWyzh9M_ePWCAZc_vLV");
  //const files = folder.getFiles();
  const files = folder.getFilesByType("application/pdf");

  
}


function dataCopy(){
  sheet = SpreadsheetApp.getActiveSpreadsheet()
  range = sheet.getDataRange()
  console.log(range.getHeight())
  backgrounds = range.getBackgrounds()
  console.log(backgrounds.length)
  console.log(range.getA1Notation())
  for (var i=sheet.getLastRow(); i>=0;i++){

  }
}
