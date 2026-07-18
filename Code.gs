var building_info = '1KYzn9vwq25lXk4xZN18cN2aZQZri4j9xtFwcWqBCoJc';
var buildingstab = 'BuildingInfo';
var LeaveType = 'LeaveType';
var misc ='Misc';
var id = '';
var filter = '';
var email = Session.getActiveUser().getEmail();
var leaveTab = '';
var Mile = '';

function doGet(e) {
   
  if (e.parameter.ID) { // add an and statement checking to see it Approve is not also a parameter
    id = e.parameter.ID;

    var test = isUser();
  
    if (test == 0) {
    var html = HtmlService.createTemplateFromFile('LeaveForm').evaluate().setTitle('Leave Request ' + id).setSandboxMode(HtmlService.SandboxMode.IFRAME); 
    return html;
    }
    else if (test == 1) {
      var html = HtmlService.createTemplateFromFile('NoView').evaluate().setTitle('Leave Request ' + id).setSandboxMode(HtmlService.SandboxMode.IFRAME); 
    return html;
    }
    else if (test == 3){
      var html = HtmlService.createTemplateFromFile('NoPermission').evaluate().setTitle('Leave Request ' + id).setSandboxMode(HtmlService.SandboxMode.IFRAME); 
    return html;
    }
  
    
  }// end if parameter
  else if (e.parameter.Filter) {
    
    filter = e.parameter.Filter
    var html2 = HtmlService.createTemplateFromFile('Leave').evaluate().setTitle('Leave').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
     return html2;  
      
  } // else mileage
  else if (e.parameter.Mile) {
    Mile = e.parameter.Mile;
   var htmlMile = HtmlService.createTemplateFromFile('Mileage-Forms').evaluate().setTitle('Mileage Forms').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
     return htmlMile; 
    
  }
  else {
    //var html3 = HtmlService.createTemplateFromFile('Redirect').evaluate().setTitle('Redirecting').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    // return html3;
  }
}

//--------------------------------------------------------------------

function getUser() {
 
  return email;
  
}

//-------------------------------------------------------------------
function isAdminOffice() {
    var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
    var miscData = miscSheet.getDataRange().getValues();
  
    var HREmail = miscData[1][1];
    var Super = miscData[2][1];
    var perEmail = miscData[8][1];
    var payEmail = miscData[9][1];
  
    if (email == HREmail || email == Super || email == perEmail || email == payEmail)
      return true;
  
   else return false;
  
}

//-------------------------------------------------------------------

//-------------------------------------------------------------------

function isUser() {
    var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
    var miscData = miscSheet.getDataRange().getValues();
    var requestDB = miscData[3][1];
    var requestTab = miscData[4][1];
    var approveTab = miscData[5][1];
    var denyTab = miscData[6][1];
    var ssDB = SpreadsheetApp.openById(requestDB).getSheetByName(requestTab);
    var rdata = ssDB.getDataRange().getValues();
    
    var i = 1;
    var ai = 1;
    var di = 1;
  
  
   while (i < rdata.length && rdata[i][0] != id)
      i++;
    
  if (i < rdata.length && rdata[i][0] == id) {
    if (email == rdata[i][3]) {
      leaveTab = requestTab;
      return 0; 
    }//if user
  } //if is ID
  else {
   var saDB = SpreadsheetApp.openById(requestDB).getSheetByName(approveTab);
   var adata = saDB.getDataRange().getValues(); 
    
   while (ai < adata.length && adata[ai][0] != id)
      ai++;
    
   if (ai < adata.length && adata[ai][0] == id) {
    if (email == adata[ai][3]) {
      leaveTab = approveTab
      return 0; 
    }//if approver
  } //if is ID
  else {
     var sdDB = SpreadsheetApp.openById(requestDB).getSheetByName(denyTab);
     var ddata = sdDB.getDataRange().getValues();    
    
      while (di < ddata.length && ddata[di][0] != id)
         di++;
    
  if (di < ddata.length && ddata[di][0] == id) {
    if (email == ddata[di][3]) {
      leaveTab = denyTab;
      return 0; 
    }//if approver
    else 
      return 1;
  } //if is ID
  else return 3;
 } 
} 
 return 1;
}

//------------------------------------------------------------------------

function getSchools() {
  var ss = SpreadsheetApp.openById(building_info).getSheetByName(buildingstab);
  var sdata = ss.getDataRange().getValues();
  
  return sdata;
}

//--------------------------------------------------------------------

function getUserLeave() {
  var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
  var miscData = miscSheet.getDataRange().getValues();
  
  var requestDB = miscData[3][1];
  
  var ssDB = SpreadsheetApp.openById(requestDB).getSheetByName(leaveTab);
  var rLeave = ssDB.getDataRange().getValues();
  var lastRow = ssDB.getDataRange().getValues().length ;
  //lastRow++;
  
  var i = 0;
  
  while (i < lastRow && rLeave[i][0] != id) {
      i++;
  }//end while
  
  if (i < lastRow && rLeave[i][0] == id) {
    var dsubmit = rLeave[i][1];
    var dsumbit = dsubmit.toString().split(" ");
    rLeave[i][1] = dsumbit[0] + " " + dsumbit[1] + " " + dsumbit[2] + " " + dsumbit[3] + " at " + dsumbit[4];
    
    var dleave = rLeave[i][6].toString().split(" "); 
    rLeave[i][6] = dleave[0] + " " + dleave[1] + " " + dleave[2] + " " + dleave[3];
    
    if (rLeave[i][8]) {
     var lleave = rLeave[i][8].toString().split(" ");
     rLeave[i][8] = lleave[0] + " " + lleave[1] + " " + lleave[2] + " " + lleave[3];
    }
    rLeave[i][26] = leaveTab;
  return rLeave[i];
  }
  else {
    
    return 0;
  }//end else

}//end function

//--------------------------------------------------------------------

function getName() {
  var name = ContactsApp.getContact(email).getFullName();  
  return name;
  
}

//--------------------------------------------------------------------

function cancelLeave(e, reason) {
  id = e;
  
  var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
  var miscData = miscSheet.getDataRange().getValues();
  
  var buildSheet = SpreadsheetApp.openById(building_info).getSheetByName(buildingstab);
  var buildData = buildSheet.getDataRange().getValues();

  var requestDB = miscData[3][1];
  var sheet = useSheet();
  var ssDB = SpreadsheetApp.openById(requestDB).getSheetByName(leaveTab);

 
  var rLeave = ssDB.getDataRange().getValues();
  var lastRow = ssDB.getDataRange().getValues().length ;
  //lastRow--;
  
  var i = 1;
  //var leave = getUserLeave();
  
  while (i < lastRow && rLeave[i][0] != id) {
      i++;
  }//end while
  
  if (rLeave[i][0] == id) {
    ssDB.deleteRow(++i);
    i--;
    rLeave[i][26] = reason;
    
    var values = [[rLeave[i][0], rLeave[i][1], rLeave[i][2], rLeave[i][3], rLeave[i][4], rLeave[i][5], rLeave[i][6], rLeave[i][7], rLeave[i][8], rLeave[i][9], rLeave[i][10], rLeave[i][11], rLeave[i][12], rLeave[i][13], rLeave[i][14], rLeave[i][15], rLeave[i][16], rLeave[i][17], rLeave[i][18], rLeave[i][19], rLeave[i][20], rLeave[i][21], rLeave[i][22], rLeave[i][23], rLeave[i][24], rLeave[i][25], rLeave[i][26]]];
    var scDB = SpreadsheetApp.openById(requestDB).getSheetByName('Cancelled');
    var clastrow = scDB.getDataRange().getValues().length;
    var range = scDB.getRange(clastrow+1, 1, 1, 27);
    range.setValues(values);
    
 }

  // send out emails
  
  //fix time/date formatting
    var dsubmit = rLeave[i][1];
    var dsumbit = dsubmit.toString().split(" ");
    rLeave[i][1] = dsumbit[0] + " " + dsumbit[1] + " " + dsumbit[2] + " " + dsumbit[3] + " at " + dsumbit[4];
    
    var dleave = rLeave[i][6].toString().split(" "); 
    rLeave[i][6] = dleave[0] + " " + dleave[1] + " " + dleave[2] + " " + dleave[3];
    
    if (rLeave[i][8]) {
     var lleave = rLeave[i][8].toString().split(" ");
     rLeave[i][8] = lleave[0] + " " + lleave[1] + " " + lleave[2] + " " + lleave[3];
    }
  
  //Build Message
  var subject = rLeave[i][2] + " has cancelled leave for: " + rLeave[i][6];
  var message= "<div>" + email + " Leave has been cancelled<br>&nbsp;</div>";
  for (var mCount = 0; mCount < rLeave[i].length; mCount++) { //building the message 
   if (rLeave[i][mCount]) {
   message = message + "<div>" + rLeave[0][mCount] + ": " + rLeave[i][mCount] + "<br>&nbsp;</div>";
   }
  }
  
  //finding the supervisor & sec emails
  mCount = 2;
  while ( mCount < buildData.length && rLeave[i][4] != buildData[mCount][0]) {
    mCount++;
  } 
  
  if (rLeave[i][20] == "Stage1 Approved") {
    //Email Linda, Princ, & Sec
    var hrSec = miscData[1][1];
    var prin = buildData[mCount][2];
    var sec = buildData[mCount][1];

    MailApp.sendEmail(hrSec, subject, message, {noReply: true, htmlBody: message});
    MailApp.sendEmail(prin, subject, message, {noReply: true, htmlBody: message});
    MailApp.sendEmail(sec, subject, message, {noReply: true, htmlBody: message});
  }
  else if (rLeave[i][20] == "Stage2 Approved") {
    //Email Super, Linda, Princ, & sec
    var super1 = miscData[2][1];
    var hrSec = miscData[1][1];
    var prin = buildData[mCount][2];
    var sec = buildData[mCount][1];
    
    MailApp.sendEmail(super1, subject, message, {noReply: true, htmlBody: message});
    MailApp.sendEmail(hrSec, subject, message, {noReply: true, htmlBody: message});
    MailApp.sendEmail(prin, subject, message, {noReply: true, htmlBody: message});
    MailApp.sendEmail(sec, subject, message, {noReply: true, htmlBody: message});
  }
  else if (rLeave[i][20] == "Superintendent Approved" || rLeave[i][20] == "Approved") {
   //Email Super, Linda, Princ, & sec
    var super1 = miscData[2][1];
    var hrSec = miscData[1][1];
    var prin = buildData[mCount][2];
    var sec = buildData[mCount][1];
    
    MailApp.sendEmail(super1, subject, message, {noReply: true, htmlBody: message});
    MailApp.sendEmail(hrSec, subject, message, {noReply: true, htmlBody: message});
    MailApp.sendEmail(prin, subject, message, {noReply: true, htmlBody: message});
    MailApp.sendEmail(sec, subject, message, {noReply: true, htmlBody: message});
  }
  else {
   //Email Princ, & Sec
    var prin = buildData[mCount][2];
    var sec = buildData[mCount][1];
    
    MailApp.sendEmail(prin, subject, message, {noReply: true, htmlBody: message});
    MailApp.sendEmail(sec, subject, message, {noReply: true, htmlBody: message});
  }
  
  if (rLeave[i][5] == "Yes") {
   //Email Spec Dir, & Spec Sec
   var specdir = buildData[mCount][4];
   var specsec = buildData[mCount][3];
    
    MailApp.sendEmail(specdir, subject, message, {noReply: true, htmlBody: message});
    MailApp.sendEmail(specsec, subject, message, {noReply: true, htmlBody: message});
  }
  
  //Email rLeave[i][3]
  MailApp.sendEmail(email, subject, message, {noReply: true, htmlBody: message});
  
}

//-------------------------------------------------------------------------------

function useSheet() {
    var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
    var miscData = miscSheet.getDataRange().getValues();
    var requestDB = miscData[3][1];
    var requestTab = miscData[4][1];
    var approveTab = miscData[5][1];
    var denyTab = miscData[6][1];
    var ssDB = SpreadsheetApp.openById(requestDB).getSheetByName(requestTab);
    var saDB = SpreadsheetApp.openById(requestDB).getSheetByName(approveTab);
    var sdDB = SpreadsheetApp.openById(requestDB).getSheetByName(denyTab);
    var rdata = ssDB.getDataRange().getValues();
    var adata = saDB.getDataRange().getValues();
    var ddata = sdDB.getDataRange().getValues();
    var i = 1;
    var ai = 1;
    var di = 1;
  
    
  while(i < rdata.length && rdata[i][0] != id) {
    i++;
  }
  
  if (rdata.length-1 == i) {
  
  while(ai < adata.length && adata[ai][0] != id) {
    ai++;
  } 
   if (adata.length-1 == ai) {
  while(di < ddata.length && ddata[di][0] != id) {
    di++;
  }
  }
    
  }
  if (i == rdata.length && di == ddata.length && ai == adata.length) {
  return 1;
  }
  else if ( rdata.length != i && rdata[i][0] == id) {
    leaveTab = "Requests";
    return "Requests"; 
  }
  else if (adata.length != ai && adata[ai][0] == id) {
    leaveTab = "Approved";
    return "Approved"; 
  }
  else if (ddata.length != di && ddata[di][0] == id) {
    leaveTab = "Denied";
    return 0; 
  }
  else {
   leaveTab = "Approved";
   return "Approved"; 
  }
  
}

//-----------------------------------------------------------------------


function getUserLeaves() {
  var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
  var miscData = miscSheet.getDataRange().getValues();
  
  var requestDB = miscData[3][1];
  var requestTab = miscData[4][1];
  
  var ssDB = SpreadsheetApp.openById(requestDB).getSheetByName(requestTab);
  var rLeave = ssDB.getDataRange().getValues();
  var lastRow = ssDB.getDataRange().getValues().length ;
  //lastRow++;
  var dsubmit;
  var dsumbit;
  var allLeave = [[],[]];
  var count = 0; 
  
  for (var i = 1; i < lastRow; i++) {
    if (rLeave[i][3] == email) {
      dsubmit = rLeave[i][1];
      dsumbit = dsubmit.toString().split(" ");
      rLeave[i][1] = dsumbit[0] + " " + dsumbit[1] + " " + dsumbit[2] + " " + dsumbit[3] + " at " + dsumbit[4];
    
      var dleave = rLeave[i][6].toString().split(" "); 
      rLeave[i][6] = dleave[0] + " " + dleave[1] + " " + dleave[2] + " " + dleave[3];
    
      if (rLeave[i][8]) {
        var lleave = rLeave[i][8].toString().split(" ");
        rLeave[i][8] = lleave[0] + " " + lleave[1] + " " + lleave[2] + " " + lleave[3];
      }
      
      allLeave[count]=rLeave[i];
      count++;
    }
  }
  
  if (count == 0) { 
    return 0;
  }
  else {
  return allLeave;
  }

}//end function

//-----------------------------------------------------------------------
function getUserApprovedLeaves() {
  var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
  var miscData = miscSheet.getDataRange().getValues();
  
  var requestDB = miscData[3][1];
  var approvedTab = miscData[5][1];
  
  var ssDB = SpreadsheetApp.openById(requestDB).getSheetByName(approvedTab);
  var rLeave = ssDB.getDataRange().getValues();
  var lastRow = ssDB.getDataRange().getValues().length ;
  //lastRow++;
  var dsubmit;
  var dsumbit;
  var allLeave = [[],[]];
  var count = 0; 
  var leaveType = getLeave();
  
  
  for (var i = 1; i < lastRow; i++) {
    if (rLeave[i][3] == email && (Mile == 'True' || filter == 100 || leaveType[filter][0] == rLeave[i][11])) {
      dsubmit = rLeave[i][1];
      dsumbit = dsubmit.toString().split(" ");
      rLeave[i][1] = dsumbit[0] + " " + dsumbit[1] + " " + dsumbit[2] + " " + dsumbit[3] + " at " + dsumbit[4];
    
      var dleave = rLeave[i][6].toString().split(" "); 
      rLeave[i][6] = dleave[0] + " " + dleave[1] + " " + dleave[2] + " " + dleave[3];
    
      if (rLeave[i][8]) {
        var lleave = rLeave[i][8].toString().split(" ");
        rLeave[i][8] = lleave[0] + " " + lleave[1] + " " + lleave[2] + " " + lleave[3];
      }
      
      allLeave[count]=rLeave[i];
      count++;
    }
  }
  
  if (count == 0) { 
    return 0;
  }
  else {
  return allLeave;
  }

}//end function

//-----------------------------------------------------------------------
function getUserDeniedLeaves() {
  var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
  var miscData = miscSheet.getDataRange().getValues();
  
  var requestDB = miscData[3][1];
  var denyTab = miscData[6][1];
  
  var ssDB = SpreadsheetApp.openById(requestDB).getSheetByName(denyTab);
  var rLeave = ssDB.getDataRange().getValues();
  var lastRow = ssDB.getDataRange().getValues().length ;
  //lastRow++;
  var dsubmit;
  var dsumbit;
  var allLeave = [[],[]];
  var count = 0; 
  
  for (var i = 1; i < lastRow; i++) {
    if (rLeave[i][3] == email) {
      dsubmit = rLeave[i][1];
      dsumbit = dsubmit.toString().split(" ");
      rLeave[i][1] = dsumbit[0] + " " + dsumbit[1] + " " + dsumbit[2] + " " + dsumbit[3] + " at " + dsumbit[4];
    
      var dleave = rLeave[i][6].toString().split(" "); 
      rLeave[i][6] = dleave[0] + " " + dleave[1] + " " + dleave[2] + " " + dleave[3];
    
      if (rLeave[i][8]) {
        var lleave = rLeave[i][8].toString().split(" ");
        rLeave[i][8] = lleave[0] + " " + lleave[1] + " " + lleave[2] + " " + lleave[3];
      }
      
      allLeave[count]=rLeave[i];
      count++;
    }
  }
  
  if (count == 0) { 
    return 0;
  }
  else {
  return allLeave;
  }

}//end function

//--------------------------------------------------------------------

function getLeave() {
  var ls = SpreadsheetApp.openById(building_info).getSheetByName(LeaveType);
  var leave = ls.getDataRange().getValues();
  
  return leave;
}

//--------------------------------------------------------------------

function returnParameter() {
  if (id != "") {
    return id;
  }
  else if (filter != "") {
    return filter;
  }
  else {
    return 0;
  }
}

//----------------------------------------------------------------------

function createLeaveSpread() {
    email = Session.getActiveUser().getEmail();
  
   var title = email + ' Leave Requests Export- ' + Utilities.formatDate(new Date(), "GMT-04:00", "E MM/dd/yyyy");
   var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
   var miscData = miscSheet.getDataRange().getValues();
  
   var RequestDB = miscData[3][1];
    var approvetab = miscData[5][1];
    var denytab = miscData[6][1];
    var RequestTab = miscData[4][1];
  
    var ss = SpreadsheetApp.openById(RequestDB).getSheetByName(RequestTab);
    var requestData = ss.getDataRange().getValues();
  
    var ss2 = SpreadsheetApp.openById(RequestDB).getSheetByName(approvetab);
    var approveData = ss2.getDataRange().getValues();
  
    var ss3 = SpreadsheetApp.openById(RequestDB).getSheetByName(denytab);
    var denyData = ss3.getDataRange().getValues();
  
    var ssNew = SpreadsheetApp.create(title);
    var url = ssNew.getUrl();
    var docid = ssNew.getId();
    var file = DriveApp.getFileById(docid);
    Logger.log(docid);
    DriveApp.getFolderById("1MMffAi2pQ2XfQJe3meUyUK2YQgMkJV6B").addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    ssNew.addEditor(email);
    ssNew.getActiveSheet().setName('Open Requests');
    //ssNew.getActiveSheet();
    
    ssNew.insertSheet('Approved');
    ssNew.insertSheet('Denied');
   
   var eheader = [[requestData[0][0], requestData[0][1], requestData[0][2], requestData[0][3], requestData[0][4], requestData[0][5], requestData[0][6], requestData[0][7], requestData[0][8], requestData[0][9], requestData[0][10], requestData[0][11], requestData[0][12], requestData[0][13], requestData[0][14], requestData[0][15], requestData[0][16], requestData[0][17], requestData[0][18], requestData[0][19], requestData[0][20], requestData[0][21], requestData[0][22], requestData[0][23], requestData[0][24], requestData[0][25]]]
  
   var sslocal = ssNew.getSheetByName('Open Requests');
   var hrange = sslocal.getRange(sslocal.getLastRow()+1, 1, 1, 26);
   hrange.setValues(eheader);
  
   var ssApprove = ssNew.getSheetByName('Approved');
   var arange = ssApprove.getRange(ssApprove.getLastRow()+1, 1, 1, 26);
   arange.setValues(eheader);
  
   var ssDeny = ssNew.getSheetByName('Denied');
   var drange = ssDeny.getRange(ssDeny.getLastRow()+1, 1, 1, 26);
   drange.setValues(eheader);
  
  var rcount = 1;
  var store = requestData[0];
  while (rcount < requestData.length) {
    if (requestData[rcount][3] == email) {
      hrange = sslocal.getRange(sslocal.getLastRow()+1, 1, 1, 26);
      store = [[requestData[rcount][0],requestData[rcount][1], requestData[rcount][2], requestData[rcount][3], requestData[rcount][4], requestData[rcount][5], requestData[rcount][6], requestData[rcount][7], requestData[rcount][8], requestData[rcount][9], requestData[rcount][10], requestData[rcount][11], requestData[rcount][12], requestData[rcount][13], requestData[rcount][14], requestData[rcount][15], requestData[rcount][16], requestData[rcount][17], requestData[rcount][18], requestData[rcount][19], requestData[rcount][20], requestData[rcount][21], requestData[rcount][22], requestData[rcount][23], requestData[rcount][24], requestData[rcount][25]]]
      hrange.setValues(store);
    }//end if email
    rcount++;
  }//end while loop
  
  rcount =1;
  while (rcount < approveData.length) {
    if (approveData[rcount][3] == email) {
      arange = ssApprove.getRange(ssApprove.getLastRow()+1, 1, 1, 26);
      store = [[approveData[rcount][0],approveData[rcount][1], approveData[rcount][2], approveData[rcount][3], approveData[rcount][4], approveData[rcount][5], approveData[rcount][6], approveData[rcount][7], approveData[rcount][8], approveData[rcount][9], approveData[rcount][10], approveData[rcount][11], approveData[rcount][12], approveData[rcount][13], approveData[rcount][14], approveData[rcount][15], approveData[rcount][16], approveData[rcount][17], approveData[rcount][18], approveData[rcount][19], approveData[rcount][20], approveData[rcount][21], approveData[rcount][22], approveData[rcount][23], approveData[rcount][24], approveData[rcount][25]]]
      arange.setValues(store);
    }//end if email
     rcount++;
  }//end while loop
  
  rcount =1;
  while (rcount < denyData.length) {
    if (denyData[rcount][3] == email) {
      drange = ssDeny.getRange(ssDeny.getLastRow()+1, 1, 1, 26);
      store = [[denyData[rcount][0],denyData[rcount][1], denyData[rcount][2], denyData[rcount][3], denyData[rcount][4], denyData[rcount][5], denyData[rcount][6], denyData[rcount][7], denyData[rcount][8], denyData[rcount][9], denyData[rcount][10], denyData[rcount][11], denyData[rcount][12], denyData[rcount][13], denyData[rcount][14], denyData[rcount][15], denyData[rcount][16], denyData[rcount][17], denyData[rcount][18], denyData[rcount][19], denyData[rcount][20], denyData[rcount][21], denyData[rcount][22], denyData[rcount][23], denyData[rcount][24], denyData[rcount][25]]]
      drange.setValues(store);
    }//end if email
     rcount++;
  }//end while loop

  var message = "Your leave requests have been exported to this <a href='" + url + "'>spreadsheet</a>";
  MailApp.sendEmail(email, title, message, {noReply: true, htmlBody: message});
   //return ssNew.getId();
}

//----------------------------------------------------------------------------------

function getMileage(mID) {
  id = mID;
  
  var title = mID + " Mileage Form";
  leaveTab = 'Approved';
  var test = DriveApp.getFilesByName(title);
  
  if (!test.hasNext()) {
  
  var doc = DriveApp.getFileById("1nJR2O92jcEhHAi-ETv0L2s9O1TFn-bi8ZFT2uTTPwTo");
  var newdoc = doc.makeCopy(title);
  
  var leave = getUserLeave();
  
  var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
  var miscData = miscSheet.getDataRange().getValues();
  var mileage = miscData[7][1];
  var cost = miscData[7][1] * leave[15];
  cost = parseFloat(cost).toFixed(2);
  
  var date = Utilities.formatDate(new Date(), "GMT-04:00", "E MM/dd/yyyy");
 
 var num = newdoc.getId();
 var final = DocumentApp.openById(num);
 var body = final.getBody();
  body.replaceText('{Date}', date)
  body.replaceText('{ID}', mID);
  body.replaceText('{Name}', leave[2]);
  body.replaceText('{Email}', leave[3]);
  body.replaceText('{DLeave}', leave[6]);
  body.replaceText('{Backup}', leave[12]);
  body.replaceText('{Destination}', leave[16]);
  body.replaceText('{Miles}', leave[15]);
  body.replaceText('{Rate}', mileage);
  body.replaceText('{Cost}', cost);
 
 final.saveAndClose();
 final.addEditor(email);
 
 /*put in email for access to URL*/
  var url = final.getUrl();
  var message = 'The mileage form for ' + mID + ' has been shared with you. The link to ' + mID + ' is: ' + url;
  var subject = id + ' mileage form has been shared with you';
  
  MailApp.sendEmail(email, subject, message);
  
  //put URL in spreadsheet
  
  return url;
  }
  else {
   var final = DriveApp.getFilesByName(title);
   var file = final.next();
   var url = file.getUrl();
   //var message = 'The mileage form for ' + mID + ' has been shared with you. The link to ' + mID + ' is: ' + url;
   //var subject = id + ' mileage form has been shared with you';
  
  //MailApp.sendEmail(email, subject, message);
   
    return url;
  }
}

//------------------------------------------------------------------------------


