var sfID = '1K_awhmCHICpW3UWbSnOMdW6ZbIoh2QNK'

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('MAP Reports')
      .addItem('List PDFs From Folder', 'listPDFsStart')
      .addItem('Share these Files with Email columns', 'shareViewDocWithEmailCols')
      .addToUi();
}

function listPDFsStart() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Please paste in the Google Drive Folder containing the PDFs',
      'Google Drive Folder URL:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  var myID = getIdFromUrl(text)
  if ((button == ui.Button.OK) && (myID)) {
      var result2 = ui.prompt(
      'How long are Student IDs?',
      'Length of Student IDs:',
      ui.ButtonSet.OK_CANCEL);
    var button2 = result2.getSelectedButton();
    var sid = parseInt(result2.getResponseText());
    if (button2== ui.Button.OK){
      if(!isNaN(sid)){
        getMAPReportIDsAIS(myID, sid);
      } else {
        throw "Error - is Student ID length entered a number? Cancelled"
      }
    }
  }
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
function getMAPReportIDs(sourceFolderID, studentIdLength) {
  sourceFolderID = sourceFolderID ? sourceFolderID : sfID ;
  
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
function getMAPReportIDsAIS(sourceFolderID, studentIdLength) {
  sourceFolderID = sourceFolderID ? sourceFolderID : sfID ;
//  var studentIdLength = 6;
  
  var mySs = SpreadsheetApp.getActive();
  var emailDir = mySs.getSheetByName("Master Ss List").getDataRange().getValues();
  var myPEmailCol = -1;
  var mySNumberCol = -1
  for (var x = 0; x < emailDir[0].length; x++){
    if (emailDir[0][x] == "Parent Email") myPEmailCol = x;
    if (emailDir[0][x] == "Student_Number") mySNumberCol = x;
  }
  if((myPEmailCol == -1) || (mySNumberCol == -1)) throw "Student_Number or Parent Email column missing"
  var mySourceFolder = DriveApp.getFolderById(sourceFolderID);
   var files = mySourceFolder.getFilesByType(MimeType.PDF);

   var fileList = [["Student ID", "Doc ID", "Student Email", "Parent Email",  "Link", "Filename"]];
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
   fileList.push([StudentId, file.getId(), StudentId+"@ais.edu.hk",parentEmail,file.getUrl(), file.getName()]);
   }
   mySs.insertSheet().getRange(1,1,fileList.length, fileList[0].length).setValues(fileList)
}

// ------
//
// function shareViewDocWithEmailCols()
//
// This function goes through the current spreadsheet and shares the doc with ID in col "Doc ID" with 
// any emails in any columns with "Email" in their header (row 1).
// The share is "silent" (ie no notification is sent to the user)
//
// ------
function shareViewDocWithEmailCols(){
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
            'role': 'reader',
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
