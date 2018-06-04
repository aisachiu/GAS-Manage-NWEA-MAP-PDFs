var thisUser = Session.getActiveUser().getEmail();
var masterListURL = SpreadsheetApp.getActiveSpreadsheet().getUrl();//Sets this Spreadsheet's URL.
var masterSheetName = "Sheet 1";
var myDefaultLink = "http://www.google.com";
var appTitle = "AIS MAP Reports Parent Portal";
var sfID = '1K_awhmCHICpW3UWbSnOMdW6ZbIoh2QNK'
var studentEmailListSheetName = "Master Email List"


function testthis(){
  var j = 'GOOGLE_DRAWINGS';
  var k = MimeType[j];
  Logger.log(k);
}

// function onOpen() - creates the menu item "Map Reports" in the spreadsheet
function onOpen() {// Or DocumentApp or FormApp
  SpreadsheetApp.getUi()
      .createMenu('MAP Reports')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Step 1 - Get Files')
      .addItem('NWEA MAP PDFs From Folder', 'listPDFsStart')
      .addItem('List Files of Type From Folder', 'listAnyFilesStart'))
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Step 2 - Share Files with Email columns')
      .addItem('Read-Only', 'shareViewDocWithEmailCols')
      .addItem('As Editor', 'shareEditDocWithEmailCols'))
      .addToUi();
}

// function listPDFsStart() - Creates a list of all the PDF files within a given Google folder
function listPDFsStart() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Please paste in the Google Drive Folder containing the PDFs',
      'Google Drive Folder URL:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if ((button == ui.Button.OK) && (text !== "")) {
    var myID = getIdFromUrl(text);
    getMAPReportIDsAIS(myID); //Use the AIS version
  }
}

// function listAnyFilesStart() - Creates a list of all the files within a given Google folder of a given filetype
function listAnyFilesStart() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Please paste in the Google Drive Folder containing the PDFs',
      'Google Drive Folder URL:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if ((button == ui.Button.OK) && (text !== "")) {
    var myID = getIdFromUrl(text);
      var result2 = ui.prompt(
      'Please enter the file (MIME) type. Possible types are GOOGLE_DRAWINGS, GOOGLE_DOCS, GOOGLE_FORMS, GOOGLE_SHEETS, GOOGLE_SITES, GOOGLE_SLIDES, FOLDER, BMP, GIF, JPEG, PNG, SVG, PDF ',
      'Mime Type:',
      ui.ButtonSet.OK_CANCEL);
      var mimeL = result2.getResponseText();
      if ((button == ui.Button.OK) && (typeof MimeType[mimeL] != 'undefined' )){
            getDocsFromFolder(myID, MimeType[mimeL]); //Use the AIS version
      }
  }
}

// ------
//
// function getDocsFromFolder
//
// this function goes through all PDF files in the designated Google folder and extracts the student ID from the end of the filename.
// It is intended to work with NWEA MAP student growth reports, which when exported as individual PDF files for each student, have
// a file name in the following pattern: "SSS_LLLLNNNN_XXX.pdf" Where XXX is the student number.
// The function will write the results onto a new sheet, with a link and the ID of the doc next to the student ID.
//
// ------
function getDocsFromFolder(sourceFolderID, MimeL) {
  sourceFolderID = sourceFolderID ? sourceFolderID : sfID ;
  //var studentIdLength = 6;
  
  var mySs = SpreadsheetApp.getActive();
  var mySourceFolder = DriveApp.getFolderById(sourceFolderID);
   var files = mySourceFolder.getFilesByType(MimeL);

   var fileList = [["Email 1", "Email 2",  "Link", "Filename", "Row Id", "Doc ID"]];
   var iterator = 1;
   while (files.hasNext()) {
   var file = files.next();
   var myName = file.getName();
   fileList.push(["" , "", file.getName(), file.getUrl(), iterator, file.getId()]);
   iterator++;
   }
   mySs.insertSheet().getRange(1,1,fileList.length, fileList[0].length).setValues(fileList)
}

// ------
//
// function getMAPReportIDs()
//
// this function goes through all PDF files in the designated Google folder and extracts the student ID from the end of the filename.
// It is intended to work with NWEA MAP student growth reports, which when exported as individual PDF files for each student, have
// a file name in the following pattern: "SSS_LLLLNNNN_XXX.pdf" Where XXX is the student number.
// The function will write the results onto a new sheet, with a link and the ID of the doc next to the student ID.
//
// ------
function getMAPReportIDs(sourceFolderID) {
  sourceFolderID = sourceFolderID ? sourceFolderID : sfID ;
  var studentIdLength = 6;
  
  var mySs = SpreadsheetApp.getActive();
  var mySourceFolder = DriveApp.getFolderById(sourceFolderID);
   var files = mySourceFolder.getFilesByType(MimeType.PDF);

   var fileList = [["Student ID", "Doc ID", "Filename", "Link"]];
   while (files.hasNext()) {
   var file = files.next();
   var myName = file.getName();
   var StudentId = myName.substr(myName.length - (4+studentIdLength), studentIdLength);
   fileList.push([StudentId, file.getId(), file.getName(), file.getUrl()]);
   }
   mySs.insertSheet().getRange(1,1,fileList.length, fileList[0].length).setValues(fileList)
}


// ------
//
// function getMAPReportIDsAIS()
//
// this function goes through all PDF files in the designated Google folder and extracts the student ID from the end of the filename.
// It is intended to work with NWEA MAP student growth reports, which when exported as individual PDF files for each student, have
// a file name in the following pattern: "SSS_LLLLNNNN_XXX.pdf" Where XXX is the student number.
// The function will write the results onto a new sheet, with a link and the ID of the doc next to the student ID.
//
// AIS version has some changes to the columns - inserting student and parent ID
//
// ------
function getMAPReportIDsAIS(sourceFolderID) {
  sourceFolderID = sourceFolderID ? sourceFolderID : sfID ;
  var studentIdLength = 6;
  
  var mySs = SpreadsheetApp.getActive();
  var emailDir = mySs.getSheetByName(studentEmailListSheetName).getDataRange().getValues();
  var myPEmailCol = -1;
  var mySNumberCol = -1;
  var myGradeLevelCol = -1
  for (var x = 0; x < emailDir[0].length; x++){
    if (emailDir[0][x] == "Parent Email") myPEmailCol = x;
    if (emailDir[0][x] == "Student_Number") mySNumberCol = x;
    if (emailDir[0][x] == "Grade_Level") myGradeLevelCol = x;
  }
  if((myPEmailCol == -1) || (mySNumberCol == -1) || (myGradeLevelCol == -1)) throw "Student_Number, Parent Email or Gradelevel column missing"
  var mySourceFolder = DriveApp.getFolderById(sourceFolderID);
   var files = mySourceFolder.getFilesByType(MimeType.PDF);

   var fileList = [["Student Email", "Parent Email",  "Link", "Filename", "Student ID", "Doc ID", "GradeLevel"]];
   while (files.hasNext()) {
     var file = files.next();
     var myName = file.getName();
     var StudentId = myName.substr(myName.length - (4+studentIdLength), studentIdLength);
     //find parent email
     var parentEmail = "unknown";
     for (var y = 1; y < emailDir.length; y++){
       if (emailDir[y][mySNumberCol] == StudentId){
         parentEmail = emailDir[y][myPEmailCol];
         break;
       }
     }
     fileList.push([ StudentId+"@ais.edu.hk",parentEmail,file.getUrl(), file.getName(),StudentId, file.getId()]);
   }
   mySs.insertSheet().getRange(1,1,fileList.length, fileList[0].length).setValues(fileList)
}

// ------
//
// function shareViewDocWithEmailCols()
//
// Calls shareDocWithEmailCols to share docs with read-only permissions.
// ------
function shareViewDocWithEmailCols(){
  shareDocWithEmailCols('reader');
}

// ------
//
// function shareEditDocWithEmailCols()
//
// Calls shareDocWithEmailCols to share docs with read-only permissions.
// ------
function shareEditDocWithEmailCols(){
  shareDocWithEmailCols('writer');
}

// ------
//
// function shareDocWithEmailCols(role)
// This function goes through the current spreadsheet and shares the doc with ID in col "Doc ID" with 
// any emails in any columns with "Email" in their header (row 1).
// The share is "silent" (ie no notification is sent to the user)
// ------
function shareDocWithEmailCols(role){
  var mySs = SpreadsheetApp.getActiveSheet();
  var myData = mySs.getDataRange().getValues();
  
  //find Email Cols and ID cols
  var emailCols = [];
  var idCol = -1
  for (var c = 0; c < myData[0].length; c++){
    if(myData[0][c].toLowerCase().search("email") > -1){ //found word email
      emailCols.push(c);
    }
    if(myData[0][c] == "Doc ID") idCol = c;
  }
  if (idCol < 0) throw "no Doc ID column found.";
  
  //for each line
  for (var r = 1; r < myData.length; r++){
    for (var e = 0; e < emailCols.length; e++){
      var thisEmail = myData[r][emailCols[e]];
      if (validateEmail(thisEmail)){
        Drive.Permissions.insert(
          {
            'role': role,
            'type': 'user',
            'value': thisEmail
          },
          myData[r][idCol],
          {
            'sendNotificationEmails': 'false'
          });
       }// End If email valid
    }//End for email cols
  }// end for each row
}

function copyTheseReports(){
var mySs = SpreadsheetApp.getActiveSheet();
  var myData = mySs.getDataRange().getValues();
  var output = [];
  var mySourceFolder = DriveApp.getFolderById('1XWvdKbXDBpOq-nUCnoNNMVfmnba_ob9T');
  var myOutputFolder = DriveApp.getFolderById('1R4oRf2oVYqxU5TM_WTjb4rCBzSGvoy3D');
  
   var files = mySourceFolder.getFilesByType(MimeType.PDF);
    Logger.log(files.hasNext());
   var fileList = {};
   while (files.hasNext()) {
   var file = files.next();
   var myName = file.getName();
   Logger.log(myName);
   var StudentId = myName.substr(myName.length - 10,6);
   Logger.log(StudentId);
   fileList[StudentId] = file.getId();
   }
  
  for(var r=1; r < myData.length; r++){
    output.push([fileList[myData[r][0]]]);
    myOutputFolder.addFile(DriveApp.getFileById(fileList[myData[r][0]]));
  }
  mySs.getRange(2,2,output.length,output[0].length).setValues(output);
}


function validateEmail(email) {
    var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(String(email).toLowerCase());
}

function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); }


function setDefaults() {
  var mySs = SpreadsheetApp.openByUrl(masterListURL);
  var mySettings = mySs.getSheetByName("Settings").getDataRange().getValues();
  myDefaultLink = mySettings[1][1]; // get default link from settings sheet
  masterSheetName = mySettings[2][1]; // get sheet name from settings sheet
 // appTitle = (typeof mySettings[3][1] !== 'undefined' && mySettings[3][1] > 0) ? mySettings[3][1] : appTitle;
  
}

// doGet() - Serves the HTML landing page.
function doGet() {
  var myDoc = 'landing';  
  return HtmlService.createTemplateFromFile(myDoc).evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// ----
// getMyLinks() - Called by the landing.html file on load.
//                This goes through the spreadsheet and seeks any rows that contain the logged-in user's email in column 1 and 2.
//                It returns an array containing the link and the title to the link.
// ----
function getMyLinks(){
  setDefaults();
  var mySs = SpreadsheetApp.openByUrl(masterListURL).getSheetByName(masterSheetName);
  var myData = mySs.getDataRange().getValues(); //Get the data in the spreadsheet
  var found = false;
  var myLink = []; //create a blank array to save all found data.
  for (var i=1; i < myData.length; i++){ //for each row
    if ((myData[i][0] == thisUser)||(myData[i][1] == thisUser)){ //if the logged in user email matches col 1 or col 2
      myLink.push([myData[i][2], myData[i][3]]); //add the link and title to the array
      found = true; //indicates that we found a link
    }
  }
  if (!found) myLink.push([myDefaultLink, "Sorry, no links found for this user "+ thisUser]); //Provide a message in form of link if no links found.
  return myLink;
}