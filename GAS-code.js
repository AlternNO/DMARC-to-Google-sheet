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

function listFilesInFolder(folderID) {
  var folder = DriveApp.getFolderById(folderID); 
  var filer = folder.getFiles();
  while(filer.hasNext()){
    var filen = filer.next();
    var filID = filen.getId();
    Logger.log(filen + filID)
    xmlinnhold(filID);
  }
}

function xmlinnhold(filID){ 
  var arkivmappe = ''; //archive folder ID
  var regnearkID = ''; // Google sheet ID
  var regneark = SpreadsheetApp.openById(regnearkID);
  Logger.log(regneark)
  var sheet = regneark.getSheetByName('report');
  Logger.log(filID)
     var data = DriveApp.getFileById(filID).getBlob().getDataAsString();
     var xml = XmlService.parse(data);
     var root = xml.getRootElement();

     var metadata = root.getChild("report_metadata");
     var email = metadata.getChild("report_id").getValue();
     var reportID = metadata.getChild("email").getValue();
     var datorom = metadata.getChild("date_range");
     var start = new Date(datorom.getChild("begin").getValue()*1000);
     var slutt = new Date(datorom.getChild("end").getValue()*1000);

     Logger.log(email)

     var records = root.getChildren("record");

     for(i in records){
      record = records[i];
      Logger.log(record);
      var rad = record.getChild("row");
      var ip = rad.getChild("source_ip").getValue();
      var antall = rad.getChild("count").getValue();
      var evaluering = rad.getChild("policy_evaluated");

      var disp = evaluering.getChild("disposition").getValue();
      var dkimpass = evaluering.getChild("dkim").getValue();
      var spfverdi = evaluering.getChild("spf").getValue();
      
      var fra = record.getChild("identifiers").getChild("header_from");

      var ar = record.getChild("auth_results");
      var dkim = ar.getChild("dkim");
      Logger.log(dkim)
      if(dkim != null){
        var dkimDomain = dkim.getChild("domain").getValue();
        var dkimResult = dkim.getChild("result").getValue();
        var dkimSelector = dkim.getChild("selector").getValue();
      }
      
      var spf = ar.getChild("spf");
      var spfDomain = spf.getChild("domain").getValue();
      var spfResult = spf.getChild("result").getValue();

      sheet.appendRow([reportID,email,start,slutt,ip,antall,disp,dkimpass,spfverdi,dkimDomain,dkimResult,dkimSelector,spfDomain,spfResult]);
      
      Logger.log(ip)
      var file = DriveApp.getFileById(filID);
      var folder = DriveApp.getFolderById(arkivmappe);
      file.moveTo(folder);
     }   

}
