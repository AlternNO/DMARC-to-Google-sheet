var folderID= ''; // Insert ID here

function saveAttachment(){
  var tittelen = "Report Domain: askerspeiderne.no"; //Edit name of domain here
  var tittelen2 = "Report domain: askerspeiderne.no"; //Edit name of domain here
  var tittelLengde = tittelen.length;
  var startdate = new Date(2021,4,10); //Remember the month count starts at 0, so 4 is May
  
  var root = DriveApp.getRootFolder();
  var parentFolder = DriveApp.getFolderById(folderID);
  var antall = 100; //How many Gmail threads will you check for DMARC-reports
  
  var threads = GmailApp.getInboxThreads(0, antall);
  for(var i in threads){
    var thread = threads[i];
    var message = thread.getMessages()[0]; // Get first message
    var tittel = message.getSubject();
    var tittelKort = tittel.substring(0,tittelLengde); //Shorten the subject to be the same length 
    var avsender = message.getFrom(); 
    var dato = message.getDate(); 
   
    if((tittelKort==tittelen || tittelKort==tittelen2)  && dato > startdate){
      Logger.log(avsender); // Log from address of the message
      var attachments = message.getAttachments();
      for(var k in attachments){
        var attachment = attachments[k];
        var attachmentBlob = attachment.copyBlob();
        var vedleggsnavn = attachment.getName();
        Logger.log(vedleggsnavn)
        var vedleggstype = attachment.getContentType();
        Logger.log(vedleggstype)
        if(vedleggstype=='application/gzip'){
          Logger.log("ja gzip");
          attachment.setContentType('application/x-gzip'); //https://stackoverflow.com/questions/67684736/unzip-gz-files-using-gas-throws-error-exception-invalid-argument/67687867#67687867
          var attachmentBlob = attachment.copyBlob();
          var files = Utilities.ungzip(attachmentBlob);
          parentFolder.createFile(files);
          }
        if(vedleggstype=='application/zip'){
          Logger.log("ja zip");
          var files = Utilities.unzip(attachmentBlob);
          files.forEach(function(file) {
           Logger.log(file)
           Logger.log(parentFolder)
           parentFolder.createFile(file);
          })
        }
        Logger.log(files)
      }
      thread.moveToArchive();
      Logger.log(tittel + " moved")
    }
  }
  listFilesInFolder(folderID)

}
