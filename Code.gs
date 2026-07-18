var building_info = '1KYzn9vwq25lXk4xZN18cN2aZQZri4j9xtFwcWqBCoJc';
//var building_info = '1kHPg2RP72zVIH4FXZw0TsVhnAV3qZ9-ECDTLo0yMp0c';
var buildingstab = 'BuildingInfo';
var LeaveType = 'LeaveType';
var misc ='Misc';
var id = '';
var filter = '';
var email = Session.getActiveUser().getEmail();
var leaveTab = '';
var Mile = '';
var Approve = '';
var emp = '';

function doGet(e) {
   
  if (e.parameter.ID && !e.parameter.Approve) { // add an and statement checking to see it Approve is not also a parameter
    id = e.parameter.ID;

    var test = isUser();
    
   if (test == 2) {
      var html = HtmlService.createTemplateFromFile('AdminView').evaluate().setTitle('Leave Request ' + id).setSandboxMode(HtmlService.SandboxMode.IFRAME); 
    return html;
    }
    else if (test == 1) {
      var html = HtmlService.createTemplateFromFile('NoView').evaluate().setTitle('Leave Request ' + id).setSandboxMode(HtmlService.SandboxMode.IFRAME); 
    return html;
    }
    else {
      var html = HtmlService.createTemplateFromFile('NoPermission').evaluate().setTitle('Leave Request ' + id).setSandboxMode(HtmlService.SandboxMode.IFRAME); 
    return html;
    }
  
    
  }// end if parameter
  
  else if (e.parameter.Approve) {
    Approve = e.parameter.Approve;
   if (e.parameter.ID) 
    id = e.parameter.ID;
    
   if (e.parameter.Filter)
    filter = e.parameter.Filter 
    
   if (e.parameter.EMP)
     emp = e.parameter.EMP;
    
    //if ID approve leave
    if (Approve == "True" && getApprover() == 1) {
     
     var htmlApprove = HtmlService.createTemplateFromFile('Approved').evaluate().setTitle('Approve').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
     return htmlApprove;
    } 
    //else show queue if action
    else if (Approve == "Queue") {
      var htmlQueue = HtmlService.createTemplateFromFile('LeaveQueue').evaluate().setTitle('Leave Queue').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
     return htmlQueue;
    }
    
    //else show request leave if view
    else if (Approve == "Requests" || Approve == "Approved" || Approve == "Denied" || Approve == "Cancelled") {
      var htmlRequest = HtmlService.createTemplateFromFile('Requests').evaluate().setTitle(Approve).setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
     return htmlRequest;
    }
    
  }
  else {
    var html3 = HtmlService.createTemplateFromFile('AdminLinks').evaluate().setTitle('Admin View Links').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
     return html3;
  }
}

//--------------------------------------------------------------------

function getUser() {
 
  return email;
  
}

//-------------------------------------------------------------------
function returnApprove() {
  
  return Approve;
  
}

//-------------------------------------------------------------------
function returnEMP() {
  
  return emp;
  
}

//-------------------------------------------------------------------
function isAdminOffice() {
    var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
    var miscData = miscSheet.getDataRange().getValues();
  
    var HREmail = miscData[1][1];
    var Superint = miscData[2][1];
    var perEmail = miscData[8][1];
    var payEmail = miscData[9][1];
  
    if (email == HREmail || email == Superint || email == perEmail || email == payEmail)
      return true;
  
   else return false;
  
}

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
    if (isApprover(requestTab, i) > 0) {
     return 2; 
    }//if approver
  } //if is ID
  else {
    var saDB = SpreadsheetApp.openById(requestDB).getSheetByName(approveTab);
    var adata = saDB.getDataRange().getValues();
    
   while (ai < adata.length && adata[ai][0] != id)
      ai++;
   if (ai < adata.length && adata[ai][0] == id) {
    if (isApprover(approveTab, ai) > 0) {
     return 2; 
    }//if approver
  } //if is ID
    else {
     var sdDB = SpreadsheetApp.openById(requestDB).getSheetByName(denyTab);
     var ddata = sdDB.getDataRange().getValues();
      
      while (di < ddata.length && ddata[di][0] != id)
         di++;
    
  if (di < ddata.length && ddata[di][0] == id) {
    if (isApprover(denyTab, di) > 0) {
     return 2; 
    }//if approver
  } //if is ID
      
    }
    
    
  }
  
  
}

//--------------------------------------------------------------------

function isApprover(tab, i) {
    var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
    var miscData = miscSheet.getDataRange().getValues();
    var requestDB = miscData[3][1];
    var ssDB = SpreadsheetApp.openById(requestDB).getSheetByName(tab);
    var rdata = ssDB.getDataRange().getValues();
    var buildSheet = SpreadsheetApp.openById(building_info).getSheetByName(buildingstab);
    var buildData = buildSheet.getDataRange().getValues();
    var count = 1; 
    var HREmail = miscData[1][1];
    var Superint = miscData[2][1];
    var perEmail = miscData[8][1];
    var payEmail = miscData[9][1];
    var techsu = '';
  
   while(count < buildData.length && rdata[i][4] != buildData[count][0]) {
      count++;
     }   
  
  if (count >= buildData.length || rdata.length <= i) {
    return 0;
  }
  else {
  
    if (buildData[count][1] == email || buildData[count][2] == email || HREmail == email || Superint == email || perEmail == email || payEmail == email || email == techsu) {
      leaveTab = tab;
      return count;
    }
    else if (rdata[i][5] == "Yes") {
      if (buildData[count][3] == email || buildData[count][4] == email) {
        leaveTab = tab;
        return count; 
      }
      else {
       return 0; 
      }
    }
    else { 
     return 0; 
    }
    
  }
  
  
  
       
     
}

//--------------------------------------------------------------------

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
  email = Session.getActiveUser().getEmail();
  var result = AdminDirectory.Users.get(email, {fields:'name', viewType:'domain_public'});
  var fullname = result.name.fullName;
  Logger.log(fullname);
  return fullname;
  
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
   return 3; 
  }
  
}

//-----------------------------------------------------------------------

function getApprover() {
 
  var ls = SpreadsheetApp.openById(building_info).getSheetByName(LeaveType);
  var leave = ls.getDataRange().getValues();
  
  if (useSheet() == "Requests") {
    leave = getUserLeave();
    var building = leave[4];
    var status = leave[20];
   
    var schools = getSchools();
    var count = 1;
    
    var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
    var miscData = miscSheet.getDataRange().getValues();
    var HREmail = miscData[1][1];
    var Superint = miscData[2][1];
    
    if (status == "Stage1 Approved" && email == HREmail) {
     return 1; 
    }
    else if (status == "Stage2 Approved" && email == HREmail) {
     return 1;
    }
    else {
      while (count < schools.length && schools[count][0] != building) {
       count++;
      }
      
      if (status == "New Request" && email == schools[count][2]){
        return 1;
      }
      else if (status == "New Spec Ed Request" && email == schools[count][4]) {
        return 1;
      }
      else return 0;
      
    }
   
    
  }
  else return 0;
  
  
  
}

//-----------------------------------------------------------------------
function getUserLeaves() {
  var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
  var miscData = miscSheet.getDataRange().getValues();
  var schoolSheet = SpreadsheetApp.openById(building_info).getSheetByName(buildingstab);
  var schools = schoolSheet.getDataRange().getValues();
  var ltypeSheet = SpreadsheetApp.openById(building_info).getSheetByName(LeaveType);
  var ltype = ltypeSheet.getDataRange().getValues();
  
  var requestDB = miscData[3][1];
  var requestTab = miscData[4][1];
  var hrSec = miscData[1][1];
  var superint = miscData[2][1];
  var office1 = miscData[8][1];
  var office2 = miscData[9][1];
  
  var ssDB = SpreadsheetApp.openById(requestDB).getSheetByName(Approve);
  var rLeave = ssDB.getDataRange().getValues();
  var lastRow = ssDB.getDataRange().getValues().length ;
  //lastRow++;
  var dsubmit;
  var dsumbit;
  var allLeave = [[],[]];
  var count = 1;
  var i = 1;
  allLeave[0] = rLeave[0];
  
    if (hrSec == email || superint == email || office1 == email || office2 == email) {
     for (var i = 1; i < lastRow; i++) {    
      var userfilter = emp + '@carrabec.org';
      var filters = filter.toString().split("-");
       //add conditions for filters
      if (filters != "") {
       if ((emp == '' || userfilter == rLeave[i][3]) && (filters[1] == 0 || rLeave[i][11] == ltype[filters[1]][0]) && (filters[0] == 0 || rLeave[i][4] == schools[filters[0]][0])) {
       //fix the date formatting
      var dleave = rLeave[i][6].toString().split(" "); 
      rLeave[i][6] = dleave[0] + " " + dleave[1] + " " + dleave[2] + " " + dleave[3];
      
      allLeave[count]=rLeave[i];
      count++;
       }//end if filters
      }// end if there are filters
      else {
      var dleave = rLeave[i][6].toString().split(" "); 
      rLeave[i][6] = dleave[0] + " " + dleave[1] + " " + dleave[2] + " " + dleave[3];
      
      allLeave[count]=rLeave[i];
      count++;
     } 
     }
    } //end if admin office people
    else {
      Logger.log(schools.length);
      for (var j=1; j < schools.length; j++) {
      if (email == schools[j][2]) {
       for (i = 1; i < rLeave.length; i++) {
         var userfilter = emp + '@carrabec.org';
         var filters = filter.toString().split("-");
         Logger.log(filters);
       if (filters != "") {
        if (rLeave[i][4] == schools[j][0] && (emp == '' || userfilter == rLeave[i][3]) && (filters[1] == 0 || rLeave[i][11] == ltype[filters[1]][0]) && (filters[0] == 0 || rLeave[i][4] == schools[filters[0]][0])) { //add conditions for filters
       
          allLeave[count] = rLeave[i];
          
          var datearray2 = allLeave[count][6].toString().split(" ");
          allLeave[count][6] = datearray2[0] + " " + datearray2[1] + " " + datearray2[2] + " " + datearray2[3];
          
          count++;
        }//end if leave stage approved
       } // end if there are filters
       else {
        if (rLeave[i][4] == schools[j][0]) { 
       
          allLeave[count] = rLeave[i];
          
          var datearray2 = allLeave[count][6].toString().split(" ");
          allLeave[count][6] = datearray2[0] + " " + datearray2[1] + " " + datearray2[2] + " " + datearray2[3];
          
          count++;
        }//end if leave stage approved
       } // else there are no filters
         
      }//end for loop
        
     }//end if princ
      
      else if (email == schools[j][4]) {
       for (i = 1; i < rLeave.length; i++) {
         var userfilter = emp + '@carrabec.org';
         var filters = filter.toString().split("-");
        if (filters != "") {
       if (rLeave[i][5] == "Yes" && rLeave[i][4] == schools[j][0] && (emp == '' || userfilter == rLeave[i][3]) && (filters[1] == 0 || rLeave[i][11] == ltype[filters[1]][0]) && (filters[0] == 0 || rLeave[i][4] == schools[filters[0]][0])) { //add conditions for filters
          allLeave[count] = rLeave[i];
          
          datearray2 = allLeave[count][6].toString().split(" ");
          allLeave[count][6] = datearray2[0] + " " + datearray2[1] + " " + datearray2[2] + " " + datearray2[3];
          
          count++;
        }//end if leave stage approved
       } // end if no filters
       else {
        allLeave[count] = rLeave[i];
          
          datearray2 = allLeave[count][6].toString().split(" ");
          allLeave[count][6] = datearray2[0] + " " + datearray2[1] + " " + datearray2[2] + " " + datearray2[3];
          
          count++;
       }
      }//end for loop
      
      
    } //end else
  } //end for loop
 }
 
  if (count == 0) { 
    return 0;
  }
  else {
  return allLeave;
  }

}//end function

//-----------------------------------------------------------------------


function getLeave() {
  var ls = SpreadsheetApp.openById(building_info).getSheetByName(LeaveType);
  var leave = ls.getDataRange().getValues();
  
  return leave;
}

//--------------------------------------------------------------------

function returnParameter() {
  if (filter != "") {
    return filter;
  }
  else {
    return 0;
  }
}

//----------------------------------------------------------------------

function createLeaveSpread() {
    email = Session.getActiveUser().getEmail();
  
   var title = email + ' Department Leave Requests Export- ' + Utilities.formatDate(new Date(), "GMT-04:00", "E MM/dd/yyyy");
   var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
   var miscData = miscSheet.getDataRange().getValues();
  
   //var RequestDB = miscData[3][1];
   var approvetab = miscData[5][1];
   var denytab = miscData[6][1];
   var RequestTab = miscData[4][1];
  
    Approve = RequestTab;
    var requestData = getUserLeaves();
  
    Approve = approvetab;
    var approveData = getUserLeaves();
  
    Approve = denytab;
    var denyData = getUserLeaves();
    
    var ssNew = SpreadsheetApp.create(title);
    var url = ssNew.getUrl();
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
      hrange = sslocal.getRange(sslocal.getLastRow()+1, 1, 1, 26);
      store = [[requestData[rcount][0],requestData[rcount][1], requestData[rcount][2], requestData[rcount][3], requestData[rcount][4], requestData[rcount][5], requestData[rcount][6], requestData[rcount][7], requestData[rcount][8], requestData[rcount][9], requestData[rcount][10], requestData[rcount][11], requestData[rcount][12], requestData[rcount][13], requestData[rcount][14], requestData[rcount][15], requestData[rcount][16], requestData[rcount][17], requestData[rcount][18], requestData[rcount][19], requestData[rcount][20], requestData[rcount][21], requestData[rcount][22], requestData[rcount][23], requestData[rcount][24], requestData[rcount][25]]]
      hrange.setValues(store);
    rcount++;
  }//end while loop
  
  rcount =1;
  while (rcount < approveData.length) {
      arange = ssApprove.getRange(ssApprove.getLastRow()+1, 1, 1, 26);
      store = [[approveData[rcount][0],approveData[rcount][1], approveData[rcount][2], approveData[rcount][3], approveData[rcount][4], approveData[rcount][5], approveData[rcount][6], approveData[rcount][7], approveData[rcount][8], approveData[rcount][9], approveData[rcount][10], approveData[rcount][11], approveData[rcount][12], approveData[rcount][13], approveData[rcount][14], approveData[rcount][15], approveData[rcount][16], approveData[rcount][17], approveData[rcount][18], approveData[rcount][19], approveData[rcount][20], approveData[rcount][21], approveData[rcount][22], approveData[rcount][23], approveData[rcount][24], approveData[rcount][25]]]
      arange.setValues(store);
     rcount++;
  }//end while loop
  
  rcount =1;
  while (rcount < denyData.length) {
      drange = ssDeny.getRange(ssDeny.getLastRow()+1, 1, 1, 26);
      store = [[denyData[rcount][0],denyData[rcount][1], denyData[rcount][2], denyData[rcount][3], denyData[rcount][4], denyData[rcount][5], denyData[rcount][6], denyData[rcount][7], denyData[rcount][8], denyData[rcount][9], denyData[rcount][10], denyData[rcount][11], denyData[rcount][12], denyData[rcount][13], denyData[rcount][14], denyData[rcount][15], denyData[rcount][16], denyData[rcount][17], denyData[rcount][18], denyData[rcount][19], denyData[rcount][20], denyData[rcount][21], denyData[rcount][22], denyData[rcount][23], denyData[rcount][24], denyData[rcount][25]]]
      drange.setValues(store);
     rcount++;
  }//end while loop

  var message = "Your leave requests have been exported to this <a href='" + url + "'>spreadsheet</a>";
  MailApp.sendEmail(email, title, message, {noReply: true, htmlBody: message});
   //return ssNew.getId();
}


//------------------------------------------------------------------------------

function denyLeave(dID, reason) {
 id = dID;
 leaveTab = 'Requests';
 
  var schools = getSchools();
   var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
  var miscData = miscSheet.getDataRange().getValues();
  
  var requestDB = miscData[3][1];
  var requestTab = miscData[4][1];
  var denyTab = miscData[6][1];
  var hrsec = miscData[1][1];
  var superint = miscData[2][1];
  
  var denyDB =SpreadsheetApp.openById(requestDB).getSheetByName(denyTab);
  var ssDB = SpreadsheetApp.openById(requestDB).getSheetByName(requestTab);
  
  var rLeave = ssDB.getDataRange().getValues();
  var lastRow = ssDB.getDataRange().getValues().length ;
  
  var i = 1;
  
  while (i < lastRow && rLeave[i][0] != id) {
      i++;
  }//end while
  
  if (i < lastRow && rLeave[i][0] == id) {
    ssDB.deleteRow(++i);
    i--;
    
    //put status into a variable
    var status = rLeave[i][20];
    
    //change status to deny + reason
    rLeave[i][20] = 'Denied - ' + reason;
    var date = new Date();
    rLeave[i][21] = rLeave[i][21] + " " + email + " Denied request on: " + date;
    
    //put in deny tab
    var values = [[rLeave[i][0], rLeave[i][1], rLeave[i][2], rLeave[i][3], rLeave[i][4], rLeave[i][5], rLeave[i][6], rLeave[i][7], rLeave[i][8], rLeave[i][9], rLeave[i][10], rLeave[i][11], rLeave[i][12], rLeave[i][13], rLeave[i][14], rLeave[i][15], rLeave[i][16], rLeave[i][17], rLeave[i][18], rLeave[i][19], rLeave[i][20], rLeave[i][21], rLeave[i][22], rLeave[i][23], rLeave[i][24], rLeave[i][25], rLeave[i][26]]];
    var range = denyDB.getRange(denyDB.getLastRow()+1, 1, 1, 27);
    range.setValues(values);
    
    //fix time/date formatting
    var dsubmit = rLeave[i][1];
    var dsumbit = dsubmit.toString().split(" ");
    rLeave[i][1] = dsumbit[0] + " " + dsumbit[1] + " " + dsumbit[2] + " " + dsumbit[3];
    
    var dleave = rLeave[i][6].toString().split(" "); 
    rLeave[i][6] = dleave[0] + " " + dleave[1] + " " + dleave[2] + " " + dleave[3];
    
    if (rLeave[i][8]) {
     var lleave = rLeave[i][8].toString().split(" ");
     rLeave[i][8] = lleave[0] + " " + lleave[1] + " " + lleave[2] + " " + lleave[3];
    }
    
    
    //get email info ready
    var message = "<div>" + rLeave[i][0] + " was denied by " + getName() + " on " + date + "<br>&nbsp;</div>";
    for (var count = 0; count < 27; count++) {
      if (rLeave[i][count] != "") {
        message = message + " <div>" + rLeave[0][count] + ": " + rLeave[i][count] + "<br>&nbsp;</div> ";  
      }
    }
    
    var j = 1;
    while (j < schools.length && schools[j][0] != rLeave[i][4])
      j++;
    
    if (j <= schools.length) {
     var princ = schools[j][2];
     var sec = schools[j][1];
     var specdir = schools[j][4];
     var specsec = schools[j][3];
    }
        
    var subject = "Leave Denied " + rLeave[i][0];
    MailApp.sendEmail(rLeave[i][3], subject, message, {noReply: true, htmlBody: message});
    
    //email appropriate people
    if (status == "Stage1 Approved" && email == miscData[1][1]) {
      //Email HR Sec
       MailApp.sendEmail(miscData[1][1], subject, message, {noReply: true, htmlBody: message});
    }
    
    //Email Princ & Sec
    MailApp.sendEmail(princ, subject, message, {noReply: true, htmlBody: message});
    MailApp.sendEmail(sec, subject, message, {noReply: true, htmlBody: message});
    
    //if spec ed email dir & sp sec
    if (rLeave[i][5] == "Yes") {
     MailApp.sendEmail(specdir, subject, message, {noReply: true, htmlBody: message});
     MailApp.sendEmail(specsec, subject, message, {noReply: true, htmlBody: message});
    }
    
    //if other contacts email them
    if (rLeave[i][25]){
     var otherc = rLeave[i][25];
     var contact = otherc.split(",");
      
      if (contact.length > 1) {
        j=1;
        
        while (j < schools.length && schools[j][0] != contact[1])
          j++;
        
        if (j <= schools.length) {
          var princ = schools[j][2];
          var sec = schools[j][1];
        }
        MailApp.sendEmail(princ, subject, message, {noReply: true, htmlBody: message});
        MailApp.sendEmail(sec, subject, message, {noReply: true, htmlBody: message});
        
      }
       j=1;
        
        while (j < schools.length && schools[j][0] != contact[0])
          j++;
        
        if (j <= schools.length) {
          var princ = schools[j][2];
          var sec = schools[j][1];
        }
        MailApp.sendEmail(princ, subject, message, {noReply: true, htmlBody: message});
        MailApp.sendEmail(sec, subject, message, {noReply: true, htmlBody: message});
    }
    
  }
}

//------------------------------------------------------------------------------

function approveLeave(dID, reason) {
 id = dID;
 leaveTab = 'Requests';
 var URL = 'https://script.google.com/a/carrabec.org/macros/s/AKfycbwabyy0F218Jhgn5vhtPlB3U61YqghnalDYmHG7X54sCdzD8wg6/exec?ID='; //for the employee
  var URL2 = 'https://script.google.com/a/carrabec.org/macros/s/AKfycbwlNVj76AW-KaDT2DnQNAExjr2vpfYfPDYOM6cGPoNLn_9pzHM/exec?ID='; //for the admin
 
  var schools = getSchools();
  var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
  var miscData = miscSheet.getDataRange().getValues();
  
  var requestDB = miscData[3][1];
  var requestTab = miscData[4][1];
  var approveTab = miscData[5][1];
  var hrsec = miscData[1][1];
  var superint = miscData[2][1];
  
  var approveDB =SpreadsheetApp.openById(requestDB).getSheetByName(approveTab);
  var ssDB = SpreadsheetApp.openById(requestDB).getSheetByName(requestTab);
  
  var rLeave = ssDB.getDataRange().getValues();
  var lastRow = ssDB.getDataRange().getValues().length ;
  
  var i = 1;
  
  while (i < lastRow && rLeave[i][0] != id) {
      i++;
  }//end while 
  
  //delete row
  if (i < lastRow && rLeave[i][0] == id) {
    ssDB.deleteRow(++i);
    i--;
    
  //fix time/date formatting
    var dsubmit = rLeave[i][1];
    var dsumbit = dsubmit.toString().split(" ");
    rLeave[i][1] = dsumbit[0] + " " + dsumbit[1] + " " + dsumbit[2] + " " + dsumbit[3];
    
    var dleave = rLeave[i][6].toString().split(" "); 
    rLeave[i][6] = dleave[0] + " " + dleave[1] + " " + dleave[2] + " " + dleave[3];
    
    if (rLeave[i][8]) {
     var lleave = rLeave[i][8].toString().split(" ");
     rLeave[i][8] = lleave[0] + " " + lleave[1] + " " + lleave[2] + " " + lleave[3];
    }
    
  //Get email info ready
   var nowDate = new Date(); 
   var today = nowDate.getMonth()+1+'/'+(nowDate.getDate())+'/'+nowDate.getFullYear();
    
   var message = "<div>" + rLeave[i][0] + " was approved on " + today + "<br>&nbsp;</div>";
    for (var count = 0; count < 27; count++) {
      if (rLeave[i][count] != "") {
        message = message + " <div>" + rLeave[0][count] + ": " + rLeave[i][count] + "<br>&nbsp;</div> ";  
      }
    }
    
    var j = 1;
    while (j < schools.length && schools[j][0] != rLeave[i][4])
      j++;
    
    if (j <= schools.length) {
     var princ = schools[j][2];
     var sec = schools[j][1];
     var specdir = schools[j][4];
     var specsec = schools[j][3];
    }
    
  
    var subject = rLeave[i][2] + " Leave " + rLeave[i][0] + " Approved by ";
    var approveSubject = rLeave[i][2] + " Leave Request " + rLeave[i][0];
    
  var approveMessage = message;
  approveMessage = approveMessage + "<div>Click <a href='" + URL2 + rLeave[i][0] + "'>here</a> to add notes and approve/deny the leave.<br></div>";
  approveMessage = approveMessage + "<div>Or you can directly approve the leave without notes: <a href='" + URL2 + rLeave[i][0] + "&Approve=True'>Approve Leave</a><br></div>";
 
  var userMessage = message;
  userMessage += "<div>Click <a href='" + URL + rLeave[i][0] + "'>here</a> to view or cancel your leave.</div>";
 
    //what workflow to use for approval
  if (rLeave[i][20] == "New Request") {
    
       subject = subject + " the Supervisor";
       rLeave[i][20] = 'Stage1 Approved';
       //var date = new Date();
       rLeave[i][21] = email + " Approved request on: " + today + "\n\n";
    
    //put in approve tab
       var values = [[rLeave[i][0], rLeave[i][1], rLeave[i][2], rLeave[i][3], rLeave[i][4], rLeave[i][5], rLeave[i][6], rLeave[i][7], rLeave[i][8], rLeave[i][9], rLeave[i][10], rLeave[i][11], rLeave[i][12], rLeave[i][13], rLeave[i][14], rLeave[i][15], rLeave[i][16], rLeave[i][17], rLeave[i][18], rLeave[i][19], rLeave[i][20], rLeave[i][21], rLeave[i][22], rLeave[i][23], rLeave[i][24], rLeave[i][25], rLeave[i][26]]];
       var range = ssDB.getRange(ssDB.getLastRow()+1, 1, 1, 27);
       range.setValues(values); 
       rLeave[i][23] = reason;
    
    //email user(with cancel link) & princ
    MailApp.sendEmail(rLeave[i][3], subject, userMessage, {noReply: true, htmlBody: userMessage});
    
    //email hr sec approval
    MailApp.sendEmail(hrsec, approveSubject, approveMessage, {noReply: true, htmlBody: approveMessage});
    MailApp.sendEmail(email, subject, message, {noReply: true, htmlBody: message});
    
  }
  else if (rLeave[i][20] == "New Spec Ed Request") {
        
       subject = subject + " the Supervisor";
       rLeave[i][20] = 'Stage1 Approved';
       //var date = new Date();
       rLeave[i][21] = email + " Approved request on: " + today; + "\n\n"
       rLeave[i][23] = reason;
   
       var values = [[rLeave[i][0], rLeave[i][1], rLeave[i][2], rLeave[i][3], rLeave[i][4], rLeave[i][5], rLeave[i][6], rLeave[i][7], rLeave[i][8], rLeave[i][9], rLeave[i][10], rLeave[i][11], rLeave[i][12], rLeave[i][13], rLeave[i][14], rLeave[i][15], rLeave[i][16], rLeave[i][17], rLeave[i][18], rLeave[i][19], rLeave[i][20], rLeave[i][21], rLeave[i][22], rLeave[i][23], rLeave[i][24], rLeave[i][25], rLeave[i][26]]];
       var range = ssDB.getRange(ssDB.getLastRow()+1, 1, 1, 27);
       range.setValues(values); 
    
    //email user(with cancel link) & princ
    MailApp.sendEmail(rLeave[i][3], subject, userMessage, {noReply: true, htmlBody: userMessage});
    
    //email hr sec approval
    MailApp.sendEmail(hrsec, approveSubject, approveMessage, {noReply: true, htmlBody: approveMessage});
    MailApp.sendEmail(email, subject, message, {noReply: true, htmlBody: message});
    
  }
  
  else if (rLeave[i][20] == "Stage1 Approved" || rLeave[i][20] == "Stage2 Approved") {
    
       subject = subject + " the Superintendent's Office";
       rLeave[i][20] = 'Approved';
       //var date = new Date();
       rLeave[i][21] = "\n\n" + rLeave[i][21] + " " + email + " Approved request on: " + today;
       rLeave[i][24] = reason;
    
       var values = [[rLeave[i][0], rLeave[i][1], rLeave[i][2], rLeave[i][3], rLeave[i][4], rLeave[i][5], rLeave[i][6], rLeave[i][7], rLeave[i][8], rLeave[i][9], rLeave[i][10], rLeave[i][11], rLeave[i][12], rLeave[i][13], rLeave[i][14], rLeave[i][15], rLeave[i][16], rLeave[i][17], rLeave[i][18], rLeave[i][19], rLeave[i][20], rLeave[i][21], rLeave[i][22], rLeave[i][23], rLeave[i][24], rLeave[i][25], rLeave[i][26]]];
       var range = approveDB.getRange(approveDB.getLastRow()+1, 1, 1, 27);
       range.setValues(values); 
    
        
     //email user(with cancel link) & princ
     MailApp.sendEmail(rLeave[i][3], subject, userMessage, {noReply: true, htmlBody: userMessage});
    
    //email hr sec & princ approval
    MailApp.sendEmail(princ, subject, message, {noReply: true, htmlBody: message});
    MailApp.sendEmail(sec, subject, message, {noReply: true, htmlBody: message});
    MailApp.sendEmail(email, subject, message, {noReply: true, htmlBody: message});
    
       //if special ed email specdir and specsec
    if (rLeave[i][5] == "Yes") {
       MailApp.sendEmail(specdir, subject, message, {noReply: true, htmlBody: message});
       MailApp.sendEmail(specsec, subject, message, {noReply: true, htmlBody: message});
    }
       
    //put in calendar if there   
    if (schools[j][8]) {
     var scal = schools[j][8];
     var cal = CalendarApp.getCalendarById(scal);
      
     if (rLeave[i][8]) {
 
     cal.createAllDayEvent(rLeave[i][2] + ' is out', new Date(rLeave[i][6]), new Date(new Date(rLeave[i][8]).getTime() + 86400000));
  
     }
    else {
    if (rLeave[i][10] == '1 Day' || rLeave[i][10] == '1.0 Day')
     cal.createAllDayEvent(rLeave[i][2] + ' is out', new Date(rLeave[i][6]));
      else {
        cal.createAllDayEvent(rLeave[i][2] + ' is out for ' + rLeave[i][10], new Date(rLeave[i][6]));
      }
     }
  }
    
     //if other contacts email them
    if (rLeave[i][25]){
     var otherc = rLeave[i][25];
     var contact = otherc.split(",");
      
      if (contact.length > 1) {
        j=1;
        
        while (j < schools.length && schools[j][0] != contact[1])
          j++;
        
        if (j <= schools.length) {
          var princ = schools[j][2];
          var sec = schools[j][1];
        }
        MailApp.sendEmail(princ, subject, message, {noReply: true, htmlBody: message});
        MailApp.sendEmail(sec, subject, message, {noReply: true, htmlBody: message});
        
      }
       j=1;
        
        while (j < schools.length && schools[j][0] != contact[0])
          j++;
        
        if (j <= schools.length) {
          var princ = schools[j][2];
          var sec = schools[j][1];
        }
        MailApp.sendEmail(princ, subject, message, {noReply: true, htmlBody: message});
        MailApp.sendEmail(sec, subject, message, {noReply: true, htmlBody: message});
    }
    
  }
 }
}

//--------------------------------------------------------------------------------------

function getQueue(){

  var schoolSheet = SpreadsheetApp.openById(building_info).getSheetByName(buildingstab);
  var schools = schoolSheet.getDataRange().getValues();
  var miscSheet = SpreadsheetApp.openById(building_info).getSheetByName(misc);
  var miscData = miscSheet.getDataRange().getValues();
  
  //get info for copying Requests Spreadsheet
  var requestDB = miscData[3][1];
  var requestTab = miscData[4][1];
  var hrsec = miscData[1][1];
  
  var ssDB = SpreadsheetApp.openById(requestDB).getSheetByName(requestTab);
  
  //Make into an array
  var rLeave = ssDB.getDataRange().getValues();
  var datearray1;
  var datearray2;
  var queue = [[],[]];
  var count = 1;
  var i = 1;
  queue[0] = rLeave[0];
  
  if (email == hrsec) {
    for (i = 1; i < rLeave.length; i++) {
      if (rLeave[i][20] == "Stage1 Approved" || rLeave[i][20] == "Stage2 Approved") {
        
        queue[count] = rLeave[i];
                  
          datearray1 = queue[count][6];
          datearray2 = datearray1.toString().split(" ");
          queue[count][6] = datearray2[0] + " " + datearray2[1] + " " + datearray2[2] + " " + datearray2[3];
        
        count++;
        
      }//end if leave stage approved
      
    }//end for loop
    
  }//if email = hrsec
  
  else { //may be a principal or special ed
    for (var j=1; j < schools.length; j++) {
      if (email == schools[j][2]) {
       for (i = 1; i < rLeave.length; i++) {
        if (rLeave[i][20] == "New Request" && rLeave[i][4] == schools[j][0]) {
          queue[count] = rLeave[i];
          
          datearray1 = queue[count][6];
          datearray2 = datearray1.toString().split(" ");
          queue[count][6] = datearray2[0] + " " + datearray2[1] + " " + datearray2[2] + " " + datearray2[3];
          
          count++;
        }//end if leave stage approved
         
      }//end for loop
        
     }//end if princ
      
      else if (email == schools[j][4]) {
       for (i = 1; i < rLeave.length; i++) {
        if (rLeave[i][20] == "New Spec Ed Request" && rLeave[i][4] == schools[j][0]) {
          queue[count] = rLeave[i];
          
          
          datearray1 = queue[count][6];
          datearray2 = datearray1.toString().split(" ");
          queue[count][6] = datearray2[0] + " " + datearray2[1] + " " + datearray2[2] + " " + datearray2[3];
          
          count++;
        }//end if leave stage approved
         
      }//end for loop
        
      }//end if specdir
      
    }//end for schools
  }//else
  
  
  return queue;
}//end getQueue()

//--------------------------------------------------------------------------------------



