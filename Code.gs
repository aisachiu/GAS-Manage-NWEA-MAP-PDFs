function myFunction() {
  var mySs = SpreadsheetApp.getActiveSpreadsheet();
  var myData = mySs.getDataRange().getValues();
  var output = [];
  var mySourceFolder = DriveApp.getFolderById('1XWvdKbXDBpOq-nUCnoNNMVfmnba_ob9T');
  var myOutputFolder = DriveApp.getFolderById('1R4oRf2oVYqxU5TM_WTjb4rCBzSGvoy3D');
  
   var files = mySourceFolder.getFilesByType(MimeType.GOOGLE_DOCS);
   var fileList = {};
 while (files.hasNext()) {
   var file = files.next();
   var myName = file.getName();
   Logger.log(myName);
   var StudentId = myName.substr(myName.length - 10,6);
   fileList.push([file.getName(), file.getId()]);
 }
  
  for(var r=1; i < myData.length; i++){
    
  }
}
