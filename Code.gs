function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Auto Update")
    .addItem("Copy Data", "dataCopy")
    .addItem("Update Market Values", "addMarketValues")
    .addItem("Automate", "Automate")
    .addToUi();
}

function Automate(){
  dataCopy()
  addMarketValues()
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
  var fileName = promptUserForInput("Please enter the name of the pdf file (ie. document.pdf):");
  var files = findFilesInDrive(fileName);
  if(files.length === 0) {
    displayToastAlert("No files with name \"" + fileName + "\" were found in Google Drive.");
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
  var txt = doc.getBody().getText();
  var title = doc.getName();
  
  // Deleted the document once the text has been stored.
  Drive.Files.remove(doc.getId());
  
  return txt
}



function getMarketValue(txt,ticker){
  
  txt = txt.split("\n").join(' ')
  txt = txt.replace(/day\(s\) added to your holding period as a result of a wash sale./g, "")
  txt = txt.replace(/Page [0-9]+ of [0-9]+ .+? (Yield | Security Identifier)/g, "")
  txt = txt.replace(/(Yield|Unrealized|Gain|Loss|Estimated|Annual Income|Accrued Interest|Cost Basis Market Value)/g, "")

  fund = "NGG"
  const regexp = new RegExp(`${ticker}[ ]+CUSIP: [a-zA-Z0-9_]{0,20} [\s\0-9/.,$%]*[Reinvestments to Date]*[\s\0-9/.,$%]*[Total Covered]*[\s\0-9/.,$%]*[Total]*[\s\0-9/.,$%]*`, 'g')
  numbers = txt.match(regexp)[0]
  numbers = numbers.replace(/[\s\a-zA-Z:]{2,}/g,"")
  numbers = numbers.trim().split(" ")
  count = 3 //third back
  if (numbers[numbers.length-1].includes("%")){ count = 4}
  if (numbers.length>=count){
    marketvalue = numbers[numbers.length-count]
    marketvalue = marketvalue.replace(/\$/g,"")
    return marketvalue
  }
}

function getCash(txt){
  const regexp = new RegExp(`Cash, Money Funds and Bank Deposits [\s\0-9/.,$%]*`, 'g')
  numbers = txt.match(regexp)[0]
  numbers = numbers.trim().split(" ")
  count = 1
  if (numbers[numbers.length-1].includes("%")){ count = 2}
  if (numbers.length>=count){
    return numbers[numbers.length-count]
  }
}

function addMarketValues(){
  sheet = getSheet()
  txt = getTextFromPDF()
  var [row,length] = getRowLengthOfBox()
  cash = getCash(txt)
  sheet.getRange("H"+(row+4)).setValue(cash)
  finalrow = row+length
  row = row+5
  values = sheet.getRange("B"+row+":B"+finalrow).getValues()
  values.forEach((ticker,indx) => { 
    ticker = ticker[0]
    if (ticker != ""){
      marketValue = getMarketValue(txt,ticker)
      v = row+indx
      sheet.getRange("H"+v).setValue(marketValue)
    }
  })
}



function allWhite(row){
  return row.every(function (e) {
    return e === "#ffffff";
  });
}
function getSheet(){
  return SpreadsheetApp.getActive().getSheetByName("Monthly")
}

function getRowLengthOfBox(){
  sheet = getSheet()
  range = sheet.getDataRange()
  backgrounds = range.getBackgrounds()
  lastrow = sheet.getLastRow()
  length = 0
  found = false
  for (var i=lastrow-1; i>=0;i--){
    if (allWhite(backgrounds[i]) & found){
      row = i+1
      break
    }
    if (!allWhite(backgrounds[i])){
      length = length + 1
      found = true
    }
  }
  length = length + 3
  row = row -2
  console.log(row)
  return [row, length]
}

function dataCopy(){
  sheet = getSheet()
  lastrow = sheet.getLastRow()
  lastcol = sheet.getLastColumn()
  
  var [row,length] = getRowLengthOfBox()
  console.log(row)
  oldRange = sheet.getRange(row,1,length,lastcol)
  newRange = sheet.getRange(row+length,1,length,lastcol)
  oldRange.copyTo(newRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  date = sheet.getRange(row+length+1,1,1,1).getValue()
  var firstDay = new Date(date.getFullYear(), date.getMonth()+1, 1);
  var lastDay = new Date(date.getFullYear(), date.getMonth() + 2, 0); 
  sheet.getRange(lastrow+2,1,1,3).setValues([[firstDay,"",lastDay]])
}

