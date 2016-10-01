var SHEET_BOOKPROFIT = "Report-BookProfitEmail";

var SHEET_SUMMARY = "Report-Summary";

var SHEET_NPA = "Report-SellNPAEmail";

var subject ="";

var INV_SHEET_ID ="1odTnC9w5iYgXoNYrdrz1fKtqWQfHtGWJZw7j4nskA9k";

var EMAIL_DATA_SHEET = "";


function sendSummaryEmail(){

  subject ="SCRIPTS : Summary Email";
  
  EMAIL_DATA_SHEET=SHEET_SUMMARY;
  
  sendEmails();

}


function sendBookProfitEmail(){

  subject ="SCRIPTS : Book Profit";
  
  EMAIL_DATA_SHEET=SHEET_BOOKPROFIT;
  
  sendEmails();

}


function sendSellNPAEmail(){

  subject ="SCRIPTS : Sell NPA";
  
  EMAIL_DATA_SHEET=SHEET_NPA;
  
  sendEmails();

}



function sendEmails() {  
  
  // Fetch the template
  var template = HtmlService.createTemplateFromFile('Reports-Email-Template');
   
  // Fetch the data the needs to be passed onto html page
  var data = getData();
  
  // IF : Data is empty.Skip Sending Email
  if(data != null && data != '' && data.length > 1){
  
    // Assign the data range the variable 'report' that is being referred in the html page.
    template.report = data;
    
    // This statement actually loads the html page
    var htmlBody = template.evaluate().getContent();
    
    // Read html from a file   
    //var htmlBody = HtmlService.createTemplateFromFile('Reports-Email-Template').evaluate().getContent();
    
    // set the emailAddress here
    var emailAddress = "sukumar.yadagiri@gmail.com";  
    
    // set the subject
    //var subject = "Script : Sending from Investments Sheet";
    
    // this function sends the email  
    MailApp.sendEmail({to: emailAddress,
                       subject: subject,
                       htmlBody: htmlBody,});
                       
  }// End IF : Data is empty.Skip Sending Email
  
}


function getData(){
  
  // get Active Sheets
  var sheet = SpreadsheetApp.openById(INV_SHEET_ID).getSheetByName(EMAIL_DATA_SHEET);
  
  var data = [];
  
  var startRow = 1;  // First row of data to process
  
  var startCol = 1;  // First Col of data to process    
  
  var lastRowNum = sheet.getLastRow();   // Get Last Row
    
  //Get the range only if we there are more than ONE row
  if(lastRowNum > 1) {
    
    // Get the last Column only if there are more rows to process.
     var lastColNum = sheet.getLastColumn(); // Get Last Column
    
    // Fetch the entire range of Report-BookProfit
    var dataRange = sheet.getRange(startRow, startCol, lastRowNum, lastColNum);
    
    // Fetch values for each row in the Range.
    data = dataRange.getDisplayValues();
  }
  
  return data;

}


