//===================================================================================
// [title]: Main script
// [Description]: WFL Processing that operates on Google Apps (server side)
// [file]: code.gs
// [Development language]: GoogleApps Sctipt
// [Update history]: 2019/07/12 Thawatchai R. (ITC Dept.)
//===================================================================================

//#IMPORTANT CONFIGURATION WorkFlow LUANCHER ===============================================================================

// 1) WEB PATH FOR SEND LINK EMAIL TO USER
//var WEB_PATH = "https://script.google.com/a/ngkntk-asia.com/macros/s/AKfycbw5h6d66C8TIR2MEJzg-tvfhHzp7Tiw63o-U69lzI_67gtThsYM/exec"; //PRD
var WEB_PATH = "https://script.google.com/a/ngkntk-asia.com/macros/s/AKfycbxaUcQBD9wJcBelB-pfTHmByYs9fOJ6YtqIChDhHA/exec"; //NEW PRD FOR (SNGK)
//var WEB_PATH = "https://script.google.com/a/ngkntk-asia.com/macros/s/AKfycbwsySkfqMpVS578qMyXM5Ugcu67p4gk9dL_H9ADycdN1-kjAg/exec"; //UAT
//var WEB_PATH = "https://script.google.com/a/ngkntk-asia.com/macros/s/AKfycbz8Wg13Z88G9Y7eYd6o84ity3bCLGX-7x7jBrtnxg/dev"; //dev

// 2) SPREAD SHEET DATABASE URL AND SHEET NAME FOR READ/WRITE DATA
//var DATA_SPREADSHEET = "https://docs.google.com/spreadsheets/d/1KziDVN-GMaLDWqAE0VRBOdbyV2Vhm5vslbKSxseMZs8/edit#gid=0"; //PRD ALL
//var DATA_SPREADSHEET = "https://docs.google.com/spreadsheets/d/1fNPwKdAJj9xRl5IlZigAAeqJvFNNJEV2HzVT57dODdg/"; //PRD AHQ
var DATA_SPREADSHEET = "https://docs.google.com/spreadsheets/d/1me8rCcLkEY4o5LAyfnIlpkcefYngXeSsWPPhHP2rQxU/"; //PRD SNGK
//var DATA_SPREADSHEET = "https://docs.google.com/spreadsheets/d/1FG9GxJMEekt6I-MPeYPHEU8CMQcgBCQ6oZec_OXHC1w/edit#gid=0"; //UAT


var INPUT_DATA = "INPUT"; //DATA WFL UPLOAD
var ROUTE_DATA = "Route"; //DATA MASTER ROUTE
var LOG_DATA = "Logs"; //DATA LOG FOR USER

// 3) SPREAD SHEET ID DATABASE URL FOR READ/WRITE LOG <USE FOR UAT>
var Sheet_Test = "1RuGoQQiY-pHzyn01Iz3Rk3kMQU79qvZApyDRZO5hEPI"
var Sheet_Test_Name = "RawData"

// 4) LINK FOR DOWNLOAD EXCEL FORM
var url_form = "https://drive.google.com/drive/folders/1bJbheMmFqY6RrCpQ3IJAUyrc2jhe0uz1" //PRD 
//var url_form = "https://drive.google.com/drive/folders/1-dALZpWsT2Rj-ibDoC_kSX-jDfswuSW5" //UAT

// 5) GOOGLE DRIVE ID FOR GET FILE FORM AND ATTACHMENT
var ATTACH_FOLDER = "1xhEJFZmNaVY2QsFwErTRLFv-wxook-jb"; //PRD
//var  ATTACH_FOLDER = "1l1J-Stol38oePR7DhHAt13k_9UFn644r"; //UAT

// 6) TEMP FOLDER FOR UPLOAD AND CHECK FILE
var TEMP_FOLDER = "1Hb0cG2DP9oaT__awnykqDk9KfyXik9n9"; //PRD
var TEMPLATE_ID;
//var TEMP_FOLDER = "17rzhxKPwJyaUtraZt6FyvrIDfeaB9SmP"; //UAT

// 7) TEMP FOLDER FOR UPLOAD AND CHECK FILE UNCONVERT
var ATTACH_FOLDER_UNCONVERT = "1sRWwjubRVediWfgaIcz4DquEjbS_pWRy";//PRD
//var ATTACH_FOLDER_UNCONVERT = "11XpoYhiwS5ZA88Pu1-eg8VT0MdqSlWms"; //UAT

// 8) OTHER DESCRIPTION
var a = "Routing";
var version_wfl = "Release 03.0220 DEV (Last on : 17-Feb-20)"; // Version to show on web
var txt_rt = ""; // route Html for render
var admin_viewer = 0; // admin mode 0= disable,1= enable
var companys; // array company data
var next_global;
var status_wfl = true;
var user_email;

// 9) TIME FOR DEBUG
var Tracking_status = 1; // 0= Off:1=On
var T1_doGet, T2_doGet, T1_inputform, T2_inputform, T1_getHistoryLog, T2_getHistoryLog;
var T1_genHead, T2_genHead;


//#END OF IMPORTANT CONFIGURATION WorkFlow LUANCHER ===============================================================================

/**
 * initial for WFL Luancher
 * @event {[*]} Url Parameter
 * @return {String} String HTML Code
 **/
function doGet(event) {
  
   
  
//  Function Code for call page Count down when deploy WFL production
//  if (new Date() < new Date('December 2, 2019 08:00:00 +0700')) {
  
//    // Count Down Page  
//    var output = HtmlService.createHtmlOutputFromFile('CountDown');
//    return output;
  
//  } else {
  
      // WFL Page
    var output,gen_html,viewer=0; // Set Variable
    user_email = Session.getActiveUser().getEmail(); // Get Email Session
    // ----- Check parameter from URL (Get) --------
    if (typeof(event.parameter.admin) != "undefined") { admin_viewer = 1; }
    if (typeof(event.parameter.view) != "undefined") { viewer = 1; }
    if (typeof(event.parameter.id) == "undefined"   ) { 
        gen_html = createInputForm("", "1", "0",viewer);  // Call Issue Page
    } else { 
        var p = findRow(event.parameter.id); // change RecID 220818
        //Logger.log(p);
        gen_html = createInputForm(p, event.parameter.rt, event.parameter.fr,viewer);  // Call Reviewer Page // change RecID 220818
    }
    // ----- End of Check parameter from URL (Get) --------
    if (admin_viewer == 1 && Tracking_status==1) {
        gen_html += "<hr>"
    }
  
    output = HtmlService.createHtmlOutput(gen_html)
    output.setTitle("WorkFlowLauncher " + version_wfl);
    return output;
//  } // END IF
}
function calculateRender(T1, T2) {    return (T2 - T1) / 1000; }

function putInputTag(pType, pId, pOpt) { return "<input type=\"" + pType + "\""+ " id=\"" + pId + "\""+ " name=\"" + pId + "\""+ " " + pOpt + ">";}

function createRadio(pData, pId, pOpt) {
    "use strict";
    //------ Define variable  ------
    var i,rg,rows,ss,strHTML,values;
  
  
    ss = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
    rg = ss.getSheetByName(pData).getRange("A:A");
    rows = rg.getLastRow();
    companys = ss.getSheetByName(pData).getRange("A1:C" + (rows)).getValues().filter(function (dataRow) {return dataRow[0] !="";});
    strHTML = "";
    for (i = 1; i <= companys.length - 1; i++) {
            strHTML = strHTML + "<label class=\"cls_" + pId + "\"><input type=\"radio\" name=\"" + pId + "\" title=\""+companys[i][1]+"\" class=\"required\" value=\"" + companys[i][0] + "\" " + pOpt + "   >" + companys[i][0] + "</label>";
    }
    return strHTML
}

function createInputForm(pId, pRt, pFr,viewer) {
var response = "";
    var strHTML,strRoute,title,FolderattId;
    var ss1 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
    var strformTitle = "<div class=\"inp_form\"><div class=\"form_title\">@title</div>";
    user_email = Session.getActiveUser().getEmail();
  //Logger.log(pId);
    docPID = pId; //PID CONTROLLED BY SYSTEM
    docPidExcel = Number(pId) - 1; //PIC PLUS ONE FOR HEADER FOR DATABASE ACCESSS
    
   // -----------Read Parameter-------------
    var para_config = ss1.getSheetByName("Config").getRange("B:B");
    var rw_para_config = para_config.getLastRow(); // Get Last Rows
    var data_config = ss1.getSheetByName("Config").getRange("A1:D" + (rw_para_config)).getValues().filter(function (dataRow) {return dataRow[1] !="";}); // [A]-Application Name,[B]-Company
    // Start of Set config
    var p_Qty_Atttach_File = data_config.filter(function(dt) { return dt[1] == 'Qty_Atttach_File';})[0][2]; // Define Maximum for Attach File.
    version_wfl = data_config.filter(function(dt) { return dt[1] == 'version_wfl';})[0][2]; // Define Maximum for Attach File.
    
    // End of set config
   //-------------------------
  
    strRoute = "";
    if (pRt != "1") { //HIDE FROM SUBMIT FROM (INITIAL)
        strRoute = getHistoryLog((docPidExcel + 1)); //docPidExcel to get Route Status
    }
  
  strHTML = "<ul id=\"ul_message\">Now loading...</ul><ul id=\"ul_message2\"></ul>"  // MODIFY BY WP 20210916
         //+ "<script>window.onload = function() {google.script.run.withSuccessHandler(onSuccess1).withFailureHandler(onFailure).test();}</script>"
         + HtmlService.createHtmlOutputFromFile('JavaScript').getContent() // Add Javascript from JavaScript.Html
         + HtmlService.createHtmlOutputFromFile('StyleSheet').getContent() // Add StyleSheet from StyleSheet.Html
         + "<link rel=\"stylesheet\" href=\"//ssl.gstatic.com/docs/script/css/add-ons.css\">" // add Css same bootrap
         + "<link rel=\"stylesheet\" href=\"https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css\">"
         + "<script src=\"https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.js\"></script>" // add ajax code
         + "<script src=\"https://apis.google.com/js/platform.js\" async defer></script>"
         + "<link href=\"https://fonts.googleapis.com/icon?family=Material+Icons\" rel=\"stylesheet\">"
         + "<link rel=\"stylesheet\" href=\"https://cdnjs.cloudflare.com/ajax/libs/animate.css/3.7.2/animate.min.css\">"
         //+ "<div id=\"preloader\"></div>"
         + "<div id=\"tab1\" class=\"tab animated  fadeInLeft faster \" >"
         + "<tbody>"
         + "<form id=\"inp\">"; // start form
    /*-------------- Create Header By Route -----------------------  */
         // Check Route from URL for select render form
         if (pRt != null) {
             // Has Route
             var form_typ = pRt.substring(0, 1);var form_ln = pRt.length;
         } else {
             // Hasn't Route
             var form_typ = "";var form_ln = "";
         }
    // GENERATE HEADER WEB PAGE
    switch ("" + form_typ + form_ln) {
    case "11":
        // Issue Form
        strHTML = strHTML + "<div><div id=\"charging\" class=\"material-icons\" style=\"color: #357ae8;\"></div><font size=\"5\" class=\"animated  fadeInLeft delay-1s\" >WorkFlow Launcher - Issue Application</font></div>";
        break;
    case "12":
    case "21":
    case "22":
        // Review Form
        strHTML = strHTML + "<div id=\"charging\" class=\"material-icons\" style=\"color: #357ae8;\"></div><font size=\"5\" class=\"animated  fadeInLeft delay-1s\">WorkFlow Launcher</font><div><font color=green size=\"3\" class=\"animated  fadeInLeft delay-1s\">" + ss1.getSheetByName(INPUT_DATA).getRange("D" + docPID).getValue() + " Request - Review Form </font></div>";
        break;
    case "31":
        // Approve Form
        strHTML = strHTML + "<div id=\"charging\" class=\"material-icons\" style=\"color: #357ae8;\"></div><font size=\"5\" class=\"animated  fadeInLeft delay-1s\">WorkFlow Launcher</font><div><font color=green size=\"3\" class=\"animated  fadeInLeft delay-1s\">" + ss1.getSheetByName(INPUT_DATA).getRange("D" + docPID).getValue() + " Request - Final Approve Form </font></div>";
        break;
    case "41":
        // Register Form
        strHTML = strHTML + "<div id=\"charging\" class=\"material-icons\" style=\"color: #357ae8;\"></div><font size=\"5\" class=\"animated  fadeInLeft delay-1s\">WorkFlow Launcher</font><div><font color=green size=\"3\" class=\"animated  fadeInLeft delay-1s\">" + ss1.getSheetByName(INPUT_DATA).getRange("D" + docPID).getValue() + " Request - Application Admin Form</font></div>";
        break;
    default: 
        // Other Form
        strHTML = strHTML + "<div id=\"charging\" class=\"material-icons\" style=\"color: #357ae8;\"></div><div><font size=\"5\" class=\"animated  fadeInLeft delay-1s\">WorkFlow Launcher - Issue Application</font></div>";
        break;
    }
    // GENERATE VERSION WEB PAGE
    strHTML = strHTML + "<font color=\"#80aaDD\" class=\"animated  fadeInUp delay-3s\" >" + version_wfl + "</font>"; // version WFL
    /*-------------- END Create Header By Route -----------------------  */


    /*-------------- Create Form By Route -----------------------  */
    // Issue Application
    if ("" + form_typ + form_ln == "11") {

        // Issue Date - Text Input
        strHTML += strformTitle.replace(/@title/g, "Issue Date")
        strHTML += "<div class=\"form_inp\">" + putInputTag("text", "txtIssueDate", "value=\"" + Utilities.formatDate(new Date(), "GMT+7", "yyyy/MM/dd") + "\"  readonly=\"readonly\"") + "</div>"
        strHTML += "</div>"

        // Issue By - Text Input
        strHTML += strformTitle.replace(/@title/g, "Issue By")
        strHTML += "<div class=\"form_inp\">" + putInputTag("text", "txtIssueBy", "value=\"" + user_email + "\"  readonly=\"readonly\" style=\"width:557px;color:#431AE6;\" ") + "</div>"
        strHTML += "</div>"
        
        // Urgent - Text Input
        strHTML += strformTitle.replace(/@title/g, "Priority")
        strHTML += "<div class=\"form_inp\">" +" <label class=\"cls_chkCompany\" ><input type=\"radio\" name=\"priority\" title=\"Priority : Normal \" class=\"required\" value=\"1\" onchange=\"changepriority(this.value)\" checked>Normal</label> <label class=\"cls_chkCompany\"><input type=\"radio\" name=\"priority\" title=\"Priority  : Urgent \"  class=\"required\" value=\"0\" onchange=\"changepriority(this.value)\" > <span class=\"attributecolor\" style=\"color:red;font-weight: bold;\">Urgent </span></label>" 
        //strHTML +=  putInputTag("text", "txtRef3_", "value=\"" + "\" style=\"width: 395px;\"   disabled") + "</div>"
        strHTML += "</div>"
        
         strHTML += "<div id='prireason' style=\"visibility: hidden; height: 0px;\" >" +strformTitle.replace(/@title/g, "Priority Reason")
        //<textarea id=\"txtComment\" name=\"txtComment\" rows=\"5\" cols=\"80\" style=\"width: 558px;\" " + (status_wfl == true ? "" : "disabled=\"disabled\"") + "></textarea>
        //strHTML +=  "<div class=\"form_inp\">"+putInputTag("textarea", "txtRef3_", "value=\""  + "\" style=\"width: 407px;\" "+reason_chk) + "</div>"
        strHTML +=  "<div class=\"form_inp\"><textarea id=\"txtRef3_\" name=\"txtRef3_\" rows=\"5\" cols=\"80\" style=\"width: 558px; \"  ></textarea></div>"
        strHTML += "</div></div>"
        
        // Issue Company  - Radio Check box
        strHTML += strformTitle.replace(/@title/g, "Company")
        strHTML += "<div class=\"form_inp\"><div class=\"chk_block\">" + createRadio("Company", "chkCompany", "onChange=\"getCompanyDept(this.value)\"") + "</div></div>"
        strHTML += "</div>"
         //Logger.log('Step 1'); //Satitpong 202312131655
         //response = UrlFetchApp.fetch("https://hook.eu2.make.com/vyw2nt9qg8hy3aherdxpm7eh9i4sy37o?ukey="+doc_name+"&doc_master="+doc_idi+"&user_on_machine="+user_email+"&machine_process=nn&biz_process=Issue Application and form type form ln is 11 and render company radio&scripting=https://script.google.com/a/ngkntk-asia.com/macros/s/AKfycbxaUcQBD9wJcBelB-pfTHmByYs9fOJ6YtqIChDhHA/exec&module_name=code.gs&function_name=createInputForm&step_process=1point0");
         //Logger.log(response.getContentText());
        // Issue Form  - Dropdown
        strHTML += strformTitle.replace(/@title/g, "Application Sheet");
        /* Form WFL List with Company  */
        //get Data master Form from Sheet [Master]
        var rg_req = ss1.getSheetByName("Routing").getRange("D:D");
        var rows_req = rg_req.getLastRow(); // Get Last Rows
        var data_req = ss1.getSheetByName("Routing").getRange("B1:D" + (rows_req)).getValues().filter(function (dataRow) {return dataRow[0] !="";}); // [A]-Application Name,[B]-Company
        var newData = new Array();
      var alen=data_req.length
      var adjlen=0 //alen-15
      for(i=adjlen;i<alen;i++){
        var row = data_req[i];
        var duplicate = false;
        for(j in newData){
          if(row[0] == newData[j][0]  && row[2] == newData[j][2]){ //changed to compare col A&B
           duplicate = true;
          }
        }
        if(!duplicate){
          newData.push(row);
        }
      }
      data_req = newData.sort();
        //Generate Form Name By Company
        for (i = 1; i < companys.length ; i++) {
                strHTML += "<div name=\"chk" + companys[i][0] + "MasterDiv\" style=\"display: none;\" class=\"form_inp\"><div class=\"chk_block\">" + getMasterListDropdown(companys[i][0], "Application Sheet", "chk" + companys[i][0] + "Master", data_req) + "  [ <a target=\"_blank\" href=\""+ url_form +"\">Download Latest Forms</a> ]</div></div>"
        }
        strHTML += "</div>";
        /* End Form WFL List with Company  */

        /* Department List with Company  */
        var data_dep = ss1.getSheetByName("Department").getRange("A:C").getValues().filter(function (dataRow) {return dataRow[0] !="";});
        strHTML += strformTitle.replace(/@title/g, "Department")
        for (i = 1; i < companys.length ; i++) {
                strHTML += "<div name=\"chk" + companys[i][0] + "DepartmentDiv\" style=\"display: none;\" class=\"form_inp\"><div class=\"chk_block\">" + getDepartmentListDropdown(companys[i][0], "Entertain Expense", "chk" + companys[i][0] + "Department", data_dep) + "</div></div>"
        }
        //var t_1d_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([user_email, "2D", t_1d_s, t_1d_e, t_1d_e - t_1d_s, version_wfl, "render data Master Department", "Issue Application"]);
        
        //var t_1e_s = new Date();
        strHTML += "<div name=\"AllDepartmentDiv\" style=\"display: none;\" class=\"form_inp\"><div class=\"chk_block\">" + setRoutingNames("listAllRouting") + "</div></div>"
        //var t_1e_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([user_email, "2E", t_1e_s, t_1e_e, t_1e_e - t_1e_s, version_wfl, "render data Master Route", "Issue Application"]);

        strHTML += "</div>"
        /* End Department List with Company  */

        // Remark for Issue
        strHTML += strformTitle.replace(/@title/g, "Remark")
        strHTML += "<div class=\"form_inp\">" + "<textarea id=\"txtRemark\" name=\"txtRemark\" rows=\"5\" cols=\"80\" style=\"width: 558px;\"></textarea>" + "</div>"
        strHTML += "</div>"

        // File Excel Form
        strHTML += strformTitle.replace(/@title/g, "File Excel Form")
        strHTML += "<div class=\"form_inp\">" + putInputTag("file", "docFile", "onchange=\"CheckFileUpload(this)\" style=\"width:500px;\" ") + "<label id=\"docFile_label\" for=\"docFile\" style=\"color:#e6071a;\" >File Type *.xlsm Only</label></div>"
        strHTML += "</div>"

        // File Attach Form
        //strHTML += strformTitle.replace(/@title/g, "+Support Doc.")
        strHTML += "<div class=\"inp_form\"><div class=\"form_title\">+Support Doc.</div><div class=\"form_inp\">"
        strHTML += "<a  class=\"buttona addico\" onclick=\"inputBtn("+p_Qty_Atttach_File+");\">Add Attach File.</a><hr>"
       // <input type="file" id="attFile" name="attFile" onchange="CheckFileAttach(this)" style="width:500px;">
        strHTML += "<div id=\"target_div\">" 
        strHTML += "<label id=\"lbl1\">1)&nbsp;</label><input type=\"file\" id=\"attFile1\" name=\"attFile1\" style=\"width: 500px; margin: 0px; color: rgb(71, 135, 237);\" onchange=\"CheckFileAttach(this);\"  ><a id=\"rem_att1\" name=\"1\" onclick='RemoveFileAttach(this);' >Remove</a><hr id=\"hr1\">"
        //+ putInputTag("file", "attFile", "onchange=\"CheckFileAttach(this)\" style=\"width:500px;\" ") + "</div>
        strHTML += "</div>"
        strHTML += "<label id=\"attFile_label\" for=\"attFile\" style=\"color:#e6071a;\" >File Support (*PDF,EXCEL,WORD,POWERPOINT,IMAGE*) Not over 10MB</label>"
        strHTML += "</div></div>"

        // Reference #1,#2,#3
        strHTML += strformTitle.replace(/@title/g, "Reference Info.")
        strHTML += " <div class=\"form_inp\"><b>#1&nbsp;</b>" + putInputTag("text", "txtRef1", "value=\"" + "\" style=\"width: 344px;\" ") + "</div>"
        strHTML += " <div class=\"form_inp\"><b>#2&nbsp;</b>" + putInputTag("text", "txtRef2", "value=\"" + Utilities.formatDate(new Date(), "GMT+7", "yyyy/MM/dd") + "\"  ") + "</div>"

        strHTML += "</div>"
        // Button Generate Email List
        strHTML += putInputTag("button", "btnGetList", "class=\"green\" value=\"Get Approve List\" onclick=\"setRoutingName(document.getElementById('inp'));\"")

        /*--------------Dynamic email route------------------*/
        // Reviewer e-mail route
        var review_num = 7;
        strHTML += "<table>"
        for (i = 1; i < review_num + 1; i++) {
            strHTML += "<tr>" +
            "<td><div style=\"display: block;\">Reviewer #" + i + "</td>" +
            "<td><input size=45 readonly type=\"text\" id=\"reviewLevel" + i + "\" " +
            "name=\"reviewLevel" + i + "\" value=\"\">" +
            "</div><button id =\"BTreviewLevel" + i + "\" type=\"button\" " + "onClick=\"RemoveRt(" + i + ")\">Remove</button>" +
            // " Comment : " +
            "<input size=45 readonly type=\"text\" id=\"CommentreviewLevel" + i + "\" style=\"display: none;\"  " +
            "name=\"CommentreviewLevel" + i + "\" value=\"\">" +
            "<input type=\"hidden\" id=\"STreviewLevel" + i + "\" name=\"STreviewLevel" + i + "\" value=\"1\">" +
            "</td>" +
            "</tr>"
        }
        // Reviewer e-mail route
        strHTML += "<tr><td><div style=\"display: block;\">Final Approve </td><td><input size=45 readonly type=\"text\" id=\"finalApprove\" name=\"finalApprove\" value=\"\"></div></div></td></tr>"
        // Reviewer e-mail route
        strHTML += "<tr><td><div style=\"display: block;\">Application Admin</td><td><input size=45 readonly type=\"text\" id=\"applicationAdmin\" name=\"applicationAdmin\" value=\"\">"
        // Reviewer e-mail route
        strHTML += " CC <input size=35 readonly type=\"text\" id=\"applicationAdminCC\" name=\"applicationAdminCC\" value=\"\"></div><!--<button type=\"button\" onClick=\"document.getElementById('applicationAdmin').value = ''; document.getElementById('applicationAdminCC').value = '';\">Remove</button>--></td></tr>"
        strHTML += "</table>"
        /*--------------End Dynamic email route------------------*/

        // DYNAMIC ROUNTING FOR EACH COMPANY --- hidden --- get value
        strHTML += "<div style=\"display: none;\" class=\"form_inp\">"
        strHTML += "<div style=\"display: none;\"><input size=35 readonly type=\"text\" id=\"txtRegister\"   name=\"txtRegister\"   value=\"\"></div>"
        strHTML += "<div style=\"display: none;\"><input size=35 readonly type=\"text\" id=\"txtCompany\"    name=\"txtCompany\"    value=\"\"></div>"
        strHTML += "<div style=\"display: none;\"><input size=35 readonly type=\"text\" id=\"txtDepartment\" name=\"txtDepartment\" value=\"\"></div>"
        strHTML += "<div style=\"display: none;\"><input size=35 readonly type=\"text\" id=\"txtMaster\"     name=\"txtMaster\"     value=\"\"></div>"
        strHTML += putInputTag("hidden", "txtPriority", "value=\"1\"")
        strHTML += "</div>"
        // DYNAMIC ROUNTING FOR EACH COMPANY

        //END ADDED FOR V2 -- Submit
         + putInputTag("button", "btnIns", "class=\"action\" onclick=\" return doAction(document.getElementById('inp'));\" value=\"Submit\"");
    }
  //Logger.log('Step 2'); 
    //END SUBMIT FORM

    //REVIEW AND APPROVE FORM
    if ("" + form_typ + form_ln > "11") {
    strHTML = strHTML;
    // Check Template for User
    var FolderTemplateID = "10BdH3hyLje-7yc3NRAl9W3jpfbCAKBCN";
    var FolderTempID = "1T0JMNMr4Bw94RMcu8TFkZl6K_B4HH37S";
    //var User_Email = user_email;
    var FolderTempUser = DriveApp.getFolderById(FolderTempID);
    var FolderTempUserEmail;
    if  (FolderTempUser.getFoldersByName(user_email).hasNext()==false) {
      FolderTempUserEmail = FolderTempUser.createFolder(user_email)
    } else {
      FolderTempUserEmail =  FolderTempUser.getFoldersByName(user_email).next();
    }
      
     // read data from Sheet
     var test = ss1.getSheetByName(INPUT_DATA).getRange("A" + docPID+":CK"+docPID).getValues();
     var doc_name = test[0][0];

     //------------- Preview PDF
     var masPID = findRowSheet(test[0][3], test[0][1], "Master");//Som 20240318 Preview PDF
     var master_sheet = ss1.getSheetByName("Master").getRange("A" + masPID + ":P" + masPID).getValues();//Som 20240318 Preview PDF
     //-------------- END Preview PDF

     var com_type = test[0][1];
     var doc_type = test[0][3];  
     var doc_idi  = test[0][8]; 
     var doc_urli = test[0][7];
      // Urgent Case
     var urgent_flg = test[0][88];
     var urgent_reason = test[0][59];
      
      urgent_chk =""
      normal_chk =""
      reason_chk =""
      if (urgent_flg==0) {
         urgent_chk="checked"
         normal_chk=""
      } else if (urgent_flg==1) {
         urgent_chk=""
         normal_chk="checked"
         reason_chk ="disabled"
         urgent_reason="";
      }
 //Logger.log('Step 3'); 
      if (FolderTempUserEmail.getFilesByName(doc_type).hasNext()==false) { //IF DOES NOT HAS TEMPLATE FILE, COPY FROM TEMPLATE MASTER FOLDER
        
        //var check1_ts =new Date();
        // get master template from id
        var master_form = ss1.getSheetByName("Master").getDataRange().getValues().filter(function (dataRow) {return dataRow[0] == doc_type && dataRow[1] == com_type;});
        
         //Browser.msgbox("Loading...  " + doc_type +  " " + com_type);
         //Logger.log("Loading...  " + doc_type + " " + com_type);
        
        var master_template_id = master_form[0][8];
        if (master_template_id==null) {return "Can't Found Master Template Please Check."}
        var master_template_file = DriveApp.getFileById(master_template_id);
        var master_template_name = master_template_file.getName();
        
        //return master_template_name;
        //var check2_ts =new Date();
        //strHTML += " Check Master Template : " + (check2_ts - check1_ts) + "MilliSec";
        
        //If not have template in folder user
        //var check3_ts =new Date();
        var file_temp,sht_temp;
        
        if (FolderTempUserEmail.getFilesByName(master_template_name).hasNext()==false) {
          
          //Browser.msgBox("Now Copy Form");
          //Browser.inputBox("Test")
          //strHTML += "<script>window.onload = function() {google.script.run.withSuccessHandler(onSuccess1).withFailureHandler(onFailure).test();}</script>"
          file_temp = master_template_file.makeCopy(master_template_name, FolderTempUserEmail).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.NONE);
          sht_temp = SpreadsheetApp.open(file_temp).insertSheet("Parameter")
          sht_temp.protect();
          sht_temp.hideSheet();
          SpreadsheetApp.open(file_temp).addEditors([user_email,"admin@ngkntk-asia.com"]);
          
        } else { 
          
          file_temp =FolderTempUserEmail.getFilesByName(master_template_name).next();
          file_temp.setName(doc_type);
          sht_temp =SpreadsheetApp.open(file_temp).getSheetByName("Parameter")
          
        } 
        
      } else { //IF TEMPLATE ALREADY EXISTS IN USER FOLDER
            
            file_temp = FolderTempUserEmail.getFilesByName(doc_type).next();
            sht_temp = SpreadsheetApp.open(file_temp).getSheetByName("Parameter")
            
      }

      // Set Parameter in Sheet Parameter
      //var check5_ts =new Date();
      var doc_i = SpreadsheetApp.open(file_temp);
      var sheets_c = doc_i.getSheets();
                for (var i = 0; i < sheets_c.length ; i++ ) {
                      sheets_c[i].clearContents();
                 }
      
      if (sht_temp != null) { //ADDED BY WNP 20210524 FIX BUG FOUND BY KUNITAKE-SAN
         sht_temp.clear() //REPLATED BY WNP
         sht_temp.getRange(1,1,6,2).setValues([["Status",""],["ID",docPID],["SheetID" ,doc_idi],["SheetURL",doc_urli],["SheetName",doc_name],["TemplateID",file_temp.getId()]]);

      }
//Logger.log('Step 4');
      TEMPLATE_ID = sht_temp.getRange('B6').getValue(); //PJA 20210820
      
         //var check6_ts =new Date();
         //strHTML += "| Copy Data : " + (check6_ts - check5_ts) + "MilliSec";

            // Company
         strHTML += strformTitle.replace(/@title/g, "Company") + "<div class=\"form_txt\">"
         strHTML += putInputTag("text", "txtCompany", "value=\"" + com_type + "\"  readonly=\"readonly\"") + "</div>"

         ///////////////////////////////DEBUG TEMPLATE ID
         /////////////////////////////////////////////////
         //strHTML += "<font color=#FFFFFF>" + doc_type + " " + com_type + " " + master_template_id + "</font></div>" 
         ///////////////////////////////END OF DEBUG TEMPLATE ID
         
         strHTML += "</div>"
            // Form Name
         strHTML += strformTitle.replace(/@title/g, "Application Sheet") + "<div class=\"form_txt\">"
         strHTML += putInputTag("text", "txtMaster", "value=\"" + doc_type + "\"    readonly=\"readonly\" style=\"width:577px;\" ") + "</div>"
         strHTML += "</div>"
                     // Urgent - Text Input
//        strHTML += strformTitle.replace(/@title/g, "Priority")
//        strHTML += "<div class=\"form_inp\">" +"<label class=\"cls_chkCompany\"><input type=\"radio\" name=\"priority\" title=\"Priority  : Urgent \"  class=\"required\" value=\"0\" " + urgent_chk + "> <span class=\"attributecolor\" style=\"color:red;font-weight: bold;\">Urgent </span></label> <label class=\"cls_chkCompany\" ><input type=\"radio\" name=\"priority\" title=\"Priority : Normal \" class=\"required\" value=\"1\" "+ normal_chk +">Normal</label>" + "</div>"
//        strHTML += "</div>"
         visi ="";
         if (pRt == "4") {
           //visi = "style=\"visibility:hidden\"";
//           if (urgent_flg==0) {
//                        normal_chk = "disabled";
//                        //urgent_chk = "disabled";
//                        reason_chk = "disabled";
//           }  else if (urgent_flg==1) {
//                        //normal_chk = "disabled";
//                        urgent_chk = "disabled";
//                        reason_chk = "disabled";
//                        urgent_reason="";
//           }
         }
      reason_chk = "disabled";
      // if (typeof(urgent_reason) == "undefined"   ) { urgent_reason=""}
        strHTML += strformTitle.replace(/@title/g, "Priority")
        strHTML += "<div class=\"form_inp\" "+ visi +" >" +" <label class=\"cls_chkCompany\" ><input type=\"radio\" name=\"priority\" title=\"Priority : Normal \" class=\"required\" value=\"1\" onchange=\"changepriority(this.value)\" "+ normal_chk +">Normal</label> <label class=\"cls_chkCompany\"><input type=\"radio\" name=\"priority\" title=\"Priority  : Urgent \"  class=\"required\" value=\"0\" onchange=\"changepriority(this.value)\" " + urgent_chk + " > <span class=\"attributecolor\" style=\"color:red;font-weight: bold;\">Urgent </span></label>" 
        //strHTML +=  putInputTag("text", "txtRef3_", "value=\"" + urgent_reason + "\" style=\"width: 407px;\" "+reason_chk) + "</div>"
        strHTML += "</div>"
        strHTML += "<div id='prireason' style=\"visibility: hidden; height: 0px;\" >" +strformTitle.replace(/@title/g, "Priority Reason")
        //<textarea id=\"txtComment\" name=\"txtComment\" rows=\"5\" cols=\"80\" style=\"width: 558px;\" " + (status_wfl == true ? "" : "disabled=\"disabled\"") + "></textarea>
        //strHTML +=  "<div class=\"form_inp\">"+putInputTag("textarea", "txtRef3_", "value=\""  + "\" style=\"width: 407px;\" "+reason_chk) + "</div>"
        strHTML +=  "<div class=\"form_inp\"><textarea id=\"txtRef3_\" name=\"txtRef3_\" rows=\"5\" cols=\"80\" style=\"width: 558px;\" " + reason_chk + "></textarea></div>"
        strHTML += "</div></div>"
          // }
            /* ------------ Excel Form --------------- */
        strHTML += strformTitle.replace(/@title/g, "Application sheet");
        // Check Autorize for Excel Form
        //var t_3b_s = new Date();
//Logger.log('Step 5');
        // Link Form Excel
        var link_url = doc_i.getUrl();
       
       //------------------ Preview PDF SOM20240318
    var sheetCode = doc_i.getSheetByName(master_sheet[0][2]).getSheetId().toString();
    var docRegister = test[0][21].toString().trim();
    var pv = test[0][16].toString().trim();//ประธาน
    var userLogin = Session.getEffectiveUser().getEmail().trim();
    var matchRegister = docRegister.search(userLogin) >= 0 ? true : false;
    var matchPv = pv.search(userLogin) >= 0 ? true : false;
    var masterStatus = master_sheet[0][12].toString().trim().toUpperCase();
    var masterScalePage = master_sheet[0][13].toString().trim().toUpperCase();
    masterScalePage = masterScalePage != "1" || masterScalePage != "2" || masterScalePage != "3" || masterScalePage != "4" ? "2" : masterScalePage;
    var docStatus = test[0][2].toString().toUpperCase();
    var docSetting = "export?gid=" + sheetCode + "&format=pdf&portrait=true&size=a4&gridlines=false&printnotes=false&scale=" + masterScalePage;
    var logTypeExport = "G-Sheet"
    var docName = master_sheet[0][0].toString() + "_" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMddHHmmss");

    if (docStatus == "Final Approved".toUpperCase()) 
    {
      if (((masterStatus == "PAU" && !matchRegister) || (masterStatus == "PUP") || (masterStatus == "PAG" && !matchRegister))) {
        link_url = convertPDF(doc_i.getId(), docName, link_url, docSetting);
        logTypeExport = "PDF";
      }
    }
    else if (docStatus == "Registerd".toUpperCase()) {
      if (((masterStatus == "PAU") || (masterStatus == "PUP") || (masterStatus == "PAG" && !matchRegister) || (masterStatus == "AAU" && !matchRegister) || (masterStatus == "AUP"))){
        link_url = convertPDF(doc_i.getId(), docName, link_url, docSetting);
        logTypeExport = "PDF";
      }
    } 
       //------------------ END Preview PDF SOM20240318


        strHTML += "<div class=\"form_inp\"><a id =\"link_form\" onclick=\"enablebutton('" + link_url + "');\">" + doc_name + "</a></div></div>";
       //window.open($('#target_link').attr('href'), '_blank');
      //enablebutton();setTimeout(function()    {        window.location = $('link_form').get('href');    },1800);
        if ((doc_urli.indexOf("#gid") == -1)) {
            //strHTML +="OK"
            // File attach upload
            if (status_wfl == true) {
                strHTML += strformTitle.replace(/@title/g, "+E-sign Doc.*")
                 + "<div class=\"form_inp\">" + putInputTag("file", "IesigFile", "onchange=\"CheckFileAttach(this)\" style=\"width:500px;\" ") + "<label id=\"attFile_label\" for=\"IesigFile\" style=\"color:#e6071a;\" > Please..download from file form [XLSM] and Please upload again.</label></div>"
                 + "</div>"
            }
        }
        // } else {
        //   strHTML = strHTML + "<font color=\"red\">This document is already singed OR you may not authorized to view attachement in this step</font></div>";
        // }
        /* ------------  End Excel Form --------------- */
        //var t_3b_e = new Date();
      /*  
      if ((pRt == "1A") || (pRt == "1B") || (pRt == "1C") || (pRt == "1D") || (pRt == "2") || (pRt == "2A") || (pRt == "2B")) {
            SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([user_email, "3B", t_3b_s, t_3b_e, t_3b_e - t_3b_s, version_wfl, "render url link form", "[" + docPID + "][Review][" + link_url + "]"]);
        }
        if ((pRt == "3")) {
            SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([user_email, "4B", t_3b_s, t_3b_e, t_3b_e - t_3b_s, version_wfl, "render url link form", "[" + docPID + "][Approved][" + link_url + "]"]);
        }
        if ((pRt == "4")) {
            SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([user_email, "5B", t_3b_s, t_3b_e, t_3b_e - t_3b_s, version_wfl, "render url link form", "[" + docPID + "][Register][" + link_url + "]"]);
        }
*/
        //var t_3d_s = new Date();
        /* ------------ Attach File --------------- */
        strHTML += strformTitle.replace(/@title/g, "Attachment");
        // Check Autorize for Excel Form
        //if ((getCurrentAuthorize(docPID,pRt) == Session.getActiveUser().getEmail()) || (admin_viewer==1)) {
        /*----------  Show Attach file--------------*/
        strHTML += "<div class=\"form_inp\">";

        var fo1 = DriveApp.getFolderById(ATTACH_FOLDER);
        var fo_att = fo1.getFoldersByName(doc_name);
        if (fo_att.hasNext()) {
            var files = fo_att.next().getFiles();
            var i_ = 1;
            if (files.hasNext()) { // has Found file attach
                strHTML += "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" >";
                while (files.hasNext()) {
                    var file = files.next();
                    // Link attach file
                  if (file.getName().indexOf("Delete")==-1) {
                    strHTML += "<tr><td style=\"width:470px;padding: 0px 0;\" >" + i_ + ") <a href=\"" + file.getUrl() + "\" target=\"_blank\" >" + file.getName() + "</a></td>";
                    if (viewer ==1)  {
                        status_wfl = false
                    }
                    if (status_wfl == true) {
                        strHTML += "<td style=\"width: 97px;padding: 0px 0;text-align: right;\"><a id =\"rem_" + i_ + "\"   onclick=\"removeattach('" + file.getId() + "')\">Remove</a></td></tr>";
                    }
                    i_++;
                  }
                }
            } 
            if (i_ == 1){
                // If not found attach file
                strHTML += " <tr><td><label style=\"color:red;\">-.none attach file.- </label><br></td></tr>";
            }
            strHTML += "</table>";
        } else {
            strHTML += "<label style=\"color:red;\">-.none fould folder attach file.- </label><br>";
        }
      
         
//Logger.log('Step 6');  
        strHTML += "</div>";
        strHTML += "</div>";
        /*---------- End Show Attach file--------------*/
        //var t_3d_e = new Date();
      /*  
      if ((pRt == "1A") || (pRt == "1B") || (pRt == "1C") || (pRt == "1D") || (pRt == "2") || (pRt == "2A") || (pRt == "2B")) {
            SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([user_email, "3D", t_3d_s, t_3d_e, t_3d_e - t_3d_s, version_wfl, "render url link attach file"]);
        }
        if ((pRt == "3")) {
            SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([user_email, "4D", t_3d_s, t_3d_e, t_3d_e - t_3d_s, version_wfl, "render url link attach file"]);
        }
        if ((pRt == "4")) {
            SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([user_email, "5D", t_3d_s, t_3d_e, t_3d_e - t_3d_s, version_wfl, "render url link attach file"]);
        }
        */
        // Get id for increase file attach upload

        //FolderattId = fo_att.getId();
      try {
           FolderattId = fo1.getFoldersByName(doc_name).next().getId()
      } catch (e) {
             strHTML = HtmlService.createHtmlOutputFromFile('LoadingPage').getContent();
            //strHTML.getContent().toString().replace("llink", "xx"); // WEB_PATH +"?id="+pId+"&rt="+pRt +"&fr="+ pFr+"&viewer="+viewer);  
            //return strHTML;
        
//           strHTML = "<ul id=\"ul_message\">Please wait, WorkFlow Luancher is now processing <br>Please try again in Later "
//           
//          //  parent.parent.window.location.replace(s_url); 
//           //pId, pRt, pFr,viewer
//           strHTML += "<br><a id=\"ll\" href=\"#\" onClick=\"parent.parent.window.location.replace('"+ WEB_PATH +"?id="+pId+"&rt="+pRt +"&fr="+ pFr+"&viewer="+viewer+"');\" >Refresh</a> </ul>"
           strHTML += "<script>setTimeout(function() {  parent.parent.window.location.replace('"+ WEB_PATH +"?id="+pId+"&rt="+pRt +"&fr="+ pFr+"&viewer="+viewer+"');;}, 7000);"
           strHTML += "x=10;setInterval(function() { document.getElementById(\"ll\").innerHTML='Refresh in ..'+(x)+' Sec' ;x-=1;}, 1000);"
           strHTML += "function refresh_libk(){parent.parent.window.location.replace('"+ WEB_PATH +"?id="+pId+"&rt="+pRt +"&fr="+ pFr+"&viewer="+viewer+"');}"
           strHTML += "</script>"
//Logger.log(strHTML);
           return strHTML;
           
      }
//Logger.log("Step 6-2");
        // File attach upload
        if ((status_wfl == true && viewer != 1 && (getCurrentAuthorize(docPID, pRt) == user_email)) || (admin_viewer == 1)) {
              // File Attach Form
        //strHTML += strformTitle.replace(/@title/g, "+Support Doc.")
        strHTML += "<div class=\"inp_form\"><div class=\"form_title\">+Support Doc.</div><div class=\"form_inp\">"
        strHTML += "<a  class=\"buttona addico\" onclick=\"inputBtn("+ p_Qty_Atttach_File + ");\">Add attach file.</a><hr>"
       // <input type="file" id="attFile" name="attFile" onchange="CheckFileAttach(this)" style="width:500px;">
        strHTML += "<div id=\"target_div\">" 
        strHTML += "<label id=\"lbl1\">1)&nbsp;</label><input type=\"file\" id=\"attFile1\" name=\"attFile[]\" style=\"width: 500px; margin: 0px; color: rgb(71, 135, 237);\" onchange=\"CheckFileAttach(this);\"     ><a id=\"rem_att1\" name=\"1\" onclick='RemoveFileAttach(this);' >Remove</a><hr id=\"hr1\">"
        //+ putInputTag("file", "attFile", "onchange=\"CheckFileAttach(this)\" style=\"width:500px;\" ") + "</div>
        strHTML += "</div>"
        strHTML += "<label id=\"attFile_label\" for=\"attFile\" style=\"color:#e6071a;\" >File Support (*PDF,EXCEL,WORD,POWERPOINT,IMAGE*) Not over 10MB</label>"
        strHTML += "</div></div>"
       
        strHTML += "<div class=\"form_inp\">" + strformTitle.replace(/@title/g, "Comment")
        strHTML += "<div class=\"form_inp\">" + "<textarea id=\"txtComment\" name=\"txtComment\" rows=\"2\" cols=\"80\" style=\"width: 558px;\" " + (status_wfl == true ? "" : "disabled=\"disabled\"") + "></textarea>" + "</div>"
        strHTML += "</div>"
        }
        if ((status_wfl == true) && (pRt!= "4") && viewer != 1 && (getCurrentAuthorize(docPID, pRt) == user_email)) {
            // Re-route dropdown
            strHTML += strformTitle.replace(/@title/g, "Re-Routing")
            strHTML += "<div class=\"form_inp\"><input type=\"checkbox\" id=\"chk_rt\" onclick=\"enableRe_Routing(this.checked)\">"
            strHTML += "<select  id=\"selreroute\" name=\"selreroute\" style=\"width:559px;\" disabled>"
            strHTML += " <option value=\"\">-</option>"
            strHTML += txt_rt // Gen email re-route in function get HistoryLog
            strHTML += "</select></div></div>";
        }
      
//Logger.log('Step 7');
//Satitpong 202312131704 
         //response = UrlFetchApp.fetch("https://hook.eu2.make.com/vyw2nt9qg8hy3aherdxpm7eh9i4sy37o?ukey="+doc_name+"&doc_master="+doc_idi+"&user_on_machine="+user_email+"&machine_process=nn&biz_process=Reroute button&scripting=https://script.google.com/a/ngkntk-asia.com/macros/s/AKfycbxaUcQBD9wJcBelB-pfTHmByYs9fOJ6YtqIChDhHA/exec&module_name=code.gs&function_name=createInputForm&step_process=7point0");
         //Logger.log(response.getContentText());
       ////////////////////////////////////////////////////////START APPROVE-REJECT-REGISTER BUTTONS
       if (viewer == 1)  {
                      
        } else {
        if ((getCurrentAuthorize(docPID, pRt) == user_email) || (admin_viewer == 1)) {
            // [u]--> Re-route,[a]-->approve,[r]-->Reject
            if ((pRt == "1A") || (pRt == "1B") || (pRt == "1C") || (pRt == "1D") || (pRt == "2") || (pRt == "2A") || (pRt == "2B")) { //AND REVIEW BUTTON
                strHTML = strHTML
                     + putInputTag("button", "btnReroute", "class=\"action\" onclick=\"return doAction2(document.getElementById('inp'),'u','Re-Route','" + pRt + "','" + docPID + "');\" value=\"Re-Route\"  hidden=\"hidden\" disabled ")
                     + " "
                     + putInputTag("button", "btnSign", "class=\"action\" onclick=\"return doAction2(document.getElementById('inp'),'a','Review Sign','" + pRt + "','" + docPID + "');\" value=\"Review Sign\" disabled")
                     + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "
                     + putInputTag("button", "btnReject", "class=\"action\" onclick=\"return doAction2(document.getElementById('inp'),'r','Reject','" + pRt + "','" + docPID + "');\" value=\"Reject\"  disabled");
            }
            if ((pRt == "3")) { //and FINAL REVIEW BUTTON
                strHTML = strHTML
                     + putInputTag("button", "btnReroute", "class=\"action\" onclick=\"return doAction2(document.getElementById('inp'),'u','Re-Route','" + pRt + "','" + docPID + "');\" value=\"Re-Route\"  hidden=\"hidden\" disabled")
                     + " "
                     + putInputTag("button", "btnSign", "class=\"action\" onclick=\"return doAction2(document.getElementById('inp'),'a','Approve','" + pRt + "','" + docPID + "');\" value=\"Approve\" disabled")
                     + " "
                     + putInputTag("button", "btnReject", "class=\"action\" onclick=\"return doAction2(document.getElementById('inp'),'r','Reject','" + pRt + "','" + docPID + "');\" value=\"Reject\"  disabled")
              
///////////////////////////////////////////////////////////////////////////////////////////////
////////////////////ADDED BY WP 202109 FOR CANCEL APPROVAL BY PDT//////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////    

                     + " "
                     + putInputTag("button", "btnCancelApprove", "class=\"action\" onclick=\"if (confirm('This will reset the document to previous state and your signature will be removed, Continue?')) { return doAction2(document.getElementById('inp'),'c','CancelApprove','" + pRt + "','" + docPID + "');}\" value=\"Cancel Action\" disabled")
              
                     + " "
                     + putInputTag("button", "btnReloadPage", "class=\"action\" onclick=\"if (confirm('This will reload the page, Continue?')) { parent.parent.window.location.replace(document.getElementById(\'txtURL\').value);}\" value=\"Reload\" disabled");
             
///////////////////////////////////////////////////////////////////////////////////////////////
////////////////////END OF ADDED BY WP 202109 FOR CANCEL APPROVAL BY PDT///////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////       

            }
            if ((pRt == "4")) { // REGISTER ACTION BUTTON
                strHTML = strHTML
                     + putInputTag("button", "btnReroute", "class=\"action\" onclick=\"return doAction2(document.getElementById('inp'),'u','Re-Route','" + pRt + "','" + docPID + "');\" value=\"Re-Route\"  hidden=\"hidden\" disabled")
                     + " "
                     + putInputTag("button", "btnSign", "class=\"action\" onclick=\"return doAction2(document.getElementById('inp'),'a','Register','" + pRt + "','" + docPID + "');\" value=\"Register\" disabled");
            }

        } else {
            if (status_wfl == true) {
                strHTML = strHTML + "<br><br><font color=\"red\">This document is already Signed OR you may not authorized to perform action in this step.</font>";
            } else {
                strHTML = strHTML + "<br><br><font color=\"red\">This document is already Rejected. </font>";
                     //+ "<H4>You can cancel the previous action and reset document stage back to before Final Approve/Reject.</H4><H5>(For Final Approver Only. System will keep this action to Log.)</H5>"
                     //+ putInputTag("button", "btnCancelApprove", "class=\"action\" onclick=\"return doAction2(document.getElementById('inp'),'c','CancelApprove','" + pRt + "','" + docPID + "');\" value=\"Cancel Action\"");
            }
        }
        }
    }
    //END OF START APPROVE-REJECT-REGISTER BUTTONS
    // Input Hidden fied for use
    strHTML = strHTML
         + putInputTag("hidden", "txtId", "value=\"" + docPID + "\"")
         + putInputTag("hidden", "txtRt", "value=\"" + pRt + "\"")
         + putInputTag("hidden", "txtJudge", "value=\"\"")
         + putInputTag("hidden", "txtFolderattId", "value=\"" + FolderattId + "\"")
         + putInputTag("hidden", "txtURL", "value=\"" + WEB_PATH + "?id=" + pId + "&rt=" + pRt + "&fr=" + pFr + "\"")
         + putInputTag("hidden", "txtDate", "value=\"\"")
         + putInputTag("hidden", "txtTempFileId", "value=\""+(doc_i != null?doc_i.getId():"")+"\"")
         + putInputTag("hidden", "txtPriority", "value=\""+urgent_flg+"\"")
         + "</div>";

    //END HISTORY LOG
    　strHTML = strHTML + "</form>" + strRoute + "</form>"
         + "</tbody>"
         + "<script>"
         + "document.querySelector('#ul_message').textContent = '';"
         + "document.querySelector('#ul_message2').textContent = '';" //ADDED BY WP 20210916 TO SEPARATE MSG OF CANCEL AND APPROVE/REJECT/REROUTE
         + "function animeicon(id,p1,p2,p3,p4) {"
         + "var a;"
         + "a = document.getElementById(id);"
         + "a.innerHTML = p1; "
         + "setTimeout(function () {"
         + "  a.innerHTML = p1;"
         + "}, 1000); "
         + "setTimeout(function () { "
         + "  a.innerHTML = p2;"
         + "}, 2000); "
         + "setTimeout(function () { "
         + "  a.innerHTML = p3; "
         + " }, 3000); "
         + "setTimeout(function () { "
         + "  a.innerHTML = p4; "
         + "}, 4000); "
         + "} "
         
         + "animeicon('charging','&#xe5dd;','&#xe5cc;','&#xe5dd;','&#xe5cc;'); "
         + "setInterval(animeicon, 5000,'charging','&#xe5dd;','&#xe5cc;','&#xe5dd;','&#xe5cc;'); "
         + "</script>";
  
    //T2_inputform = new Date();
    return strHTML;

}

function getMasterListDropdown(pCompany, pSheetname, pObjId, values) {
    "use strict";
    // return values;
    var i,rg,rows,ss,strHTML,values;
    strHTML = "<select id=\"" + pObjId + "\" name=\"" + pObjId + "\" onChange=\"setRegisterName(this.value,document.getElementsByName('txtCompany')[0].value)\" style=\"width: 327px;\" >";
    strHTML = strHTML + "<option value=\"\">-- Please Select --</option>";
    for (i = 0; i < values.length; i++) {
        if ((values[i][2] != "") && (values[i][0] == pCompany) && (values[i][2].substring(1, 1) != "x")) { // if A
            strHTML = strHTML + "<option value=\"" + values[i][2] + "\">" + values[i][2] + "</option>";
        } //end if A
    } //end for i
    strHTML = strHTML + "</select>";
    return strHTML;
}

function getDepartmentListDropdown(pCompany, pApplication, pObjId, values) {
    "use strict";
    var i,rg,rows,cols,ss,strHTML,values,txtselect,txtvalue;
    strHTML = "<select id=\"" + pObjId + "\" name=\"" + pObjId + "\" onChange=\"setValueToText(this.value,'txtDepartment');\" style=\"width: 327px;\" >";
    strHTML = strHTML + "<option value=\"\">-- Please Select --</option>"
    for (i = 1; i < values.length; i++) {
        if (values[i][0] == pCompany) { // if A
            strHTML = strHTML + "<option value=\"" + values[i][0] + "|" + values[i][2] + "|" + values[i][1] + "\"><b>[" + values[i][1] + "] - " + values[i][2] + "</b></option>";
        } //end if A
    } //end for i
    strHTML = strHTML + "</select>";
    return strHTML;
}

function getRegisterNameDropdown(pSheetname, pObjId) {

    "use strict";
    var i,    rg,    rows,    ss,    strHTML,    values;
    　 ss = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
    rg = ss.getSheetByName("POSITION").getDataRange();
    rows = rg.getLastRow();
    values = rg.getValues();

    strHTML = "<select id=\"" + pObjId + "\" name=\"" + pObjId + "\" style=\"width: 327px;\">";
    strHTML = strHTML + "<option value=\"\">-- Please Select --</option>";

    for (i = 1; i < rows; i++) {

        if (values[i][7] != "") { // if A

            strHTML = strHTML + "<option value=\"" + values[i][7] + "|" + values[i][8] + "\">" + values[i][10] + "</option>";

        } //end if A

    } //end for i

    strHTML = strHTML + "</select>";
    return strHTML;
}

function setRoutingNames(pObjId) {
    // var t1 = new Date();
    var i,    rg,    rows,    cols,    ss,    strHTML,    values,    selVal,    selTxt;
    ss = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
    rg = ss.getSheetByName("Routing").getDataRange();
    rows = rg.getLastRow();
    cols = rg.getLastColumn();
    values = rg.getValues();
    values = ss.getSheetByName("Routing").getRange(1, 1, rows + 1, cols + 1).getValues() //getRange("A1:B"+(rows+1)).getValues();

        strHTML = "<select id=\"" + pObjId + "\" name=\"" + pObjId + "\">";
    strHTML = strHTML + "<option value=\"\">-- Please Select --</option>";

    for (i = 1; i < rows; i++) {
            selTxt =   values[i][4].replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'\"') + "|"
                     + values[i][5].replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'\"') + "|"
                     + values[i][6].replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'\"') + "|"
                     + values[i][7].replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'\"') + "|" 
                     + values[i][8].replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'\"') + "|" 
                     + values[i][9].replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'\"') + "|" 
                     + values[i][10].replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'\"')+ "|" 
                     + values[i][11].replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'\"')+ "|" 
                     + values[i][12].replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'\"')+ "|" 
                     + values[i][13].replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'\"');
      //return selTxt
            selVal = values[i][1] + "|" + values[i][2] + "|" + values[i][3];
        if (selVal == "ANGK|IT Center(Applications)|DB Sheet - Items (Form : v1.22)") {
            //Logger.log("<option value=\"" + selVal + "\">" + selTxt.replace('"','\"').replace('<','&lt;').replace('>','&gt;')  + "</option>")
        }
        strHTML = strHTML + "<option value=\"" + selVal + "\">" + selTxt.replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;') + "</option>";

    }

    strHTML = strHTML + "</select>";
    // var t2 = new Date();
    // var diff = Math.floor((t2-t1)/(1000.0));//seconds
    return strHTML; // + "time : " +diff;

}

function rerouting(e) {
   // var t_3d_s = new Date();
    var id = e.txtId;
    var rt = e.txtRt;
    var user_ = Session.getActiveUser().getEmail().replace(/\s/g, "");
    var comment = e.txtComment;
    var priority = e.priority;
  var UniqueStr = GetRow2UniID(e.txtId);
    var rt_pth = ["1A", "1B", "1C", "1D", "2", "2A", "2B", "3", "4"];
    var rt_pos = [27, 32, 37, 42, 11, 48, 53, 16, 21];
    var to = e.selreroute;
    if (to != "") {
        var lst_ind = rt_pth.length;
        var frm_ind = rt_pth.indexOf(rt);
        var to_ind = rt_pth.indexOf(to);

        var doc_t = SpreadsheetApp.openByUrl(DATA_SPREADSHEET) // Active WFL
            //var DATA_SPREADSHEET = "https://docs.google.com/spreadsheets/d/1k0jInJm37XdD9oT2Ocn8NMpK1cvjXteESaBQpJ0eitY/edit#gid=0";
            var sht = doc_t.getSheetByName(INPUT_DATA);
        var rt_prv,
        email_re;
        for (var i = to_ind; i < lst_ind; i++) {
          if( sht.getRange(id, rt_pos[i]).getValue().indexOf("Remove-") == -1){

            sht.getRange(id, rt_pos[i]).setValue("");
            sht.getRange(id, (rt_pos[i] + 2)).setValue("");
            sht.getRange(id, (rt_pos[i] + 3)).setValue("");
            sht.getRange(id, (rt_pos[i] + 4)).setValue("");
          }
        }
        for (var i = to_ind; i >= 0; i--) {
            if (sht.getRange(id, (rt_pos[i] + 1)).getValue() != "") {
                rt_prv = rt_pth[i];
                email_re = sht.getRange(id, (rt_pos[i] + 1)).getValue();
                break;
            }
        }

        var doc_e_id = sht.getRange(id, 9).getValue();
        var doc_e = SpreadsheetApp.openById(doc_e_id); // Active WFL
        var sht1 = doc_e.getSheetByName("History");
        var last_row = sht1.getDataRange().getLastRow() + 1;
        sht1.getRange("A" + last_row).setValue(last_row - 1);
        sht1.getRange("B" + last_row).setValue(user_);
        sht1.getRange("C" + last_row).setValue("Re-route");
        sht1.getRange("D" + last_row).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
        sht1.getRange("E" + last_row).setValue(comment);

        sht1.autoResizeColumn(1);
        sht1.autoResizeColumn(2);
        sht1.autoResizeColumn(3);
        sht1.autoResizeColumn(4);
        sht1.autoResizeColumn(5);

        //var next = getNextApproveStep(id,"1A");
        // nextProcApprove = next.split("|");
        var txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=" + rt_prv + "&rt=" + to;
        var mlsubj;
        if (to == "3") {
            mlsubj = "Approver";
        } else {
            mlsubj = "Reviewer";
        }
      // Remove Signature
      
        var search_row,check_template_sheet; //,file_for_check_esig;
        //file_for_check_esig = SpreadsheetApp.openById(doc_id);

          // var sheet =doc_e.getSheets()[0];
           var master_form = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName("Master").getDataRange().getValues().filter(function (dataRow) {return dataRow[0] === e.txtMaster && dataRow[1] ===e.txtCompany;});
           var master_sheet_name = master_form[0][2];

            sheet = doc_e.getSheetByName(master_sheet_name);
            var search_string = Session.getActiveUser().getEmail();
            var textFinder = sheet.createTextFinder(search_string);
            search_row = textFinder.findNext();

                if (search_row != null) {
                do {
                      search_row.clearContent();
                      search_row = textFinder.findNext(); 
                     } while (search_row!= null); 
                } 
      

        var ml = email_re; //nextProcApprove[1];
        var msg1 = "You have received Application from WorkFlow Launcher.\r\n"
             + "Please Check & Confirm the Re-route following link.\r\n\r\n＜Link＞\r\n"
             + txtUrl + "\r\n\r\n"
             + "comment : " + comment + " \r\n";
        sht.getRange("C" + id).setValue('Reviewed');
        //act = "Signed By Reviewer";
      urgent_chk=""
      if (priority==0) {urgent_chk="[!!!URGENT!!!]"}

        MailApp.sendEmail(ml, " WFL <" + urgent_chk + "Re-Route  " + mlsubj + "> " + sht.getRange(id, 1).getValue(), msg1);
        //var t_3d_e = new Date();
        var pRt = rt;
      /*
        if ((pRt == "1A") || (pRt == "1B") || (pRt == "1C") || (pRt == "1D") || (pRt == "2") || (pRt == "2A") || (pRt == "2B")) {
            SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([user_email, "3G", t_3d_s, t_3d_e, t_3d_e - t_3d_s, version_wfl, "Reroute"]);
        }
        if ((pRt == "3")) {
            SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([user_email, "4G", t_3d_s, t_3d_e, t_3d_e - t_3d_s, version_wfl, "Reroute"]);
        }
        if ((pRt == "4")) {
            SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([user_email, "5G", t_3d_s, t_3d_e, t_3d_e - t_3d_s, version_wfl, "Reroute"]);
        }
*/
        return 2;
    } else {
        return " !!! Please select user to Re-route !!! "
    }
}

function preSubmitForm(e) {
 
    return 1;
  
}

// Generate RecID WFL
function getUniqueStr(myStrong){
 var strong = 1000;
 if (myStrong) strong = myStrong;
 return new Date().getTime().toString(16)  + Math.floor(strong*Math.random()).toString(16);
}

function findRow(pId){

  var ss,val,wk,p,z;

  p = pId.length--;
  ss = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
  val = ss.getSheetByName(INPUT_DATA).getRange("A:A").getValues();;
  
  for(var i=0;i<val.length;i++){
    wk = val[i][0];
    
    if  (wk.length > p) {
   　　　z = wk.split('_');
　　　　　if (z[0] == pId) {
           return i+1;
        }
    }
  }
  return '';
}

function GetRow2UniID(pId){

  var ss,val,z;

  ss = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
  val = ss.getSheetByName(INPUT_DATA).getRange('A' + pId).getValue();
  //Logger.log("val :" + val);
  z = val.split('_');
  return z[0];
}

//------------------ Preview PDF SOM20240318
function findRowSheet(pId1, pId2, sheetNames) {
  var ss, val1, val2, wk1, wk2;

  ss = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
  val1 = ss.getSheetByName(sheetNames).getRange("A:A").getValues();
  val2 = ss.getSheetByName(sheetNames).getRange("B:B").getValues();
  for (var i = 0; i < val1.length; i++) {
    wk1 = val1[i][0];
    wk2 = val2[i][0];
    if (wk1 == pId1 && wk2 == pId2) {
      return i + 1;
    }
  }
  return '';
}

//------------------ END Preview PDF SOM20240318

function submitForm(e) {
    
    var ss1,ss2,ss3;
    var i,rg,rows,values;
    var rt,ml,mlsubj,sender;
    var pk;
    
    ss1 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName(INPUT_DATA);
    rg = ss1.getRange("A:A");
  


    if (e.docFile.length != 0) {

        //Logger.log("Check File OK");
      
        var fileBlob = e.docFile;

        var folder = DriveApp.getFolderById(ATTACH_FOLDER);
        var tmp_folder = DriveApp.getFolderById(TEMP_FOLDER);
        var f_name = fileBlob.getName();

        var t_name = "TMP_" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss") + "-" + e.txtIssueBy;
        var App_name = e.txtMaster;
        var Comp_ = e.txtCompany;

        var Master_Convert = [];
        Master_Convert = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName("Master").getRange("A:H").getValues().filter(
                function (e) {
                  
                if (e[0].toString() != "") {
                    return (App_name.indexOf(e[0].toString()) > -1) && (Comp_.indexOf(e[1].toString())>-1);
                } else {
                    return false;
                }
            });
        Master_Convert.push(["", ""]);
      
        //Logger.log("Get Master Data - Master_Convert[0][6].toString().toLowerCase() = " + Master_Convert[0][6].toString().toLowerCase());
 
        var Convert_flg = (Master_Convert[0][0] == "" ? 'true' : Master_Convert[0][6].toString().toLowerCase());
        var file_des,doc;
       
  
        if (Convert_flg == "true") {

            var doc_t = GASConvertExcelToSpreadsheetActive.convertExcel2Sheets(fileBlob, t_name, TEMP_FOLDER);
            var rows_ = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName("Master").getDataRange().getLastRow();
            var Master_ver = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName("Master").getRange(1, 1, rows_, 12).getValues();
            //return "Joe-ok"
          
            var template_sheet,
            form_Sheet;
            //Column : [Application Name]	[COMPANY]	[Sheet Name]	[Pos_Ver]	[VER.DOC]	[Check_Sign]	[Conv_Google]	[Price_Autorize]	[Template_ID]   [Sheet Form]
              //Logger.log("Check version & signature  Start : ")
            for (var k = 1; k < rows_; k++) {
                if (Master_ver[k][0] == "") {}
                else {
                    if ((Master_ver[k][0] == App_name) && (Master_ver[k][1] == Comp_) && (Master_ver[k][2] != "")) {

                        var chk_sheet = doc_t.getSheetByName(Master_ver[k][2]);
                        template_sheet = Master_ver[k][8];
                        form_Sheet = Master_ver[k][2];
                        if (chk_sheet != null) {
                               // Check company
                           if (Master_ver[k][9] !== null && Master_ver[k][9] !== '') {
                             if (chk_sheet.getRange(Master_ver[k][9]).getValue().trim()!= Comp_.trim()) {
                               //return chk_sheet.getRange(Master_ver[k][9]).getValue().trim()+ " :: " + Comp_.trim()
                               return 31;
                             }
                           } 
                           
                           if (Master_ver[k][11] == 'PR amount greater than or equal to 50,000 THB or other currency'){
                              if (Master_ver[k][10] !== null && Master_ver[k][10] !== '') {
                                if (chk_sheet.getRange(Master_ver[k][10]).getValue() !== Master_ver[k][11]) {
                                  //return chk_sheet.getRange(Master_ver[k][9]).getValue().trim()+ " :: " + Comp_.trim()
                                  return 311;
                                }
                              }
                           }
                           if (Master_ver[k][11] == 'PR amount less than 50,000 THB'){
                              if (Master_ver[k][10] !== null && Master_ver[k][10] !== '') {
                                if (chk_sheet.getRange(Master_ver[k][10]).getValue() !== Master_ver[k][11]) {
                                  //return chk_sheet.getRange(Master_ver[k][9]).getValue().trim()+ " :: " + Comp_.trim()
                                  return 312;
                                }
                              }
                           }


                          //return "CHK";
                            var ver_ = "" + chk_sheet.getRange(Master_ver[k][3]).getValue();
                            //return ver_+ ":"+Master_ver[k][4];
                            if (ver_ != "") {
                                // Compare version with master
                                if (ver_ != Master_ver[k][4]) {
                                  //  ss1.deleteRow(rows);
                                    return 3;
                                } 

                                //var t_2h_s = new Date();
                                // Check sign if Master set range
                                if (((Master_ver[k][5] == null) || (Master_ver[k][5].toString().trim() == "")) == false) {
                                    // return "Error";
                                    //return Master_ver[k][5]=="" ;
                                  //Logger.log(chk_sheet.getRange(Master_ver[k][2]+"!"+Master_ver[k][5]).Name)
                                    if ((chk_sheet.getRange(Master_ver[k][2]+"!"+Master_ver[k][5]).getValue() == null) || (chk_sheet.getRange(Master_ver[k][2]+"!"+Master_ver[k][5]).getValue() == "")) {
                                      //  ss1.deleteRow(rows);
                                        return 4;
                                    }
                                }
                                // var t_2h_e = new Date();
                                //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(),"2H",t_2h_s,t_2h_e,t_2h_e-t_2h_s,version_wfl,"Check File Signature"]);

                            } else {
                             // ss1.deleteRow(rows);//appendRow([rows+"_"]);
                              return 3;
                            }
                        } else {
                            //ss1.deleteRow(rows);
                            return 3;
                        }
                    }
                }
            }
 
            //Logger.log("Check version & signature Stop : ")
            //var t_2g_e = new Date();
           // SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "2G", t_2g_s, t_2g_e, t_2g_e - t_2g_s, version_wfl, "Check File Version"]);
            //var t10 = new Date();
             //Logger.log("Check Move File temp to keep Form Folder  Start : ")
            //var t_2h_s = new Date();
            /*doc_t.getName()
            var file_src = tmp_folder.getFilesByName(doc_t.getName()).next();*/ // SOM20240122 Change from move by file name to Google ID
            var file_src = DriveApp.getFileById(doc_t.getId()); // SOM20240122

            file_des = file_src.makeCopy(folder);
            
            file_des.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.NONE);
            file_src.setTrashed(true);
                    
           var t_copy=0;
             //Logger.log("Check Move File temp to keep Form Folder  Stop : ")
          doc = SpreadsheetApp.open(file_des);
          doc.setSpreadsheetTimeZone("Asia/Bangkok");
            //--------------------------------------------------------

            var sht_input_id = 0;

            var sht_itm = doc.getSheetByName("INPUT sheet");
            if (sht_itm != null) {
                sht_input_id = "" + doc.getSheetByName("INPUT sheet").getSheetId();
            }

            var doc_id = doc.getId();
            var doc_url = doc.getUrl() + "#gid=" + sht_input_id;

        } else {

            //Uncovert

            // ATTACH_FOLDER_UNCONVERT
            var folder_uncon = DriveApp.getFolderById(ATTACH_FOLDER_UNCONVERT);
            var file_form = folder_uncon.createFile(fileBlob);
            var ff = folder_uncon.getFilesByName(file_form.getName()).next()
            var folder_con = DriveApp.getFolderById(ATTACH_FOLDER);

            var file_sprtsht = SpreadsheetApp.create('Creatsheet-' + Session.getActiveUser().getEmail(), 100, 10)
                file_sprtsht.getSheetByName("Sheet1").getRange("A1").setValue("This Sheet will use Download to Stamp Route.")
            var file_src = DriveApp.getFilesByName(file_sprtsht.getName()).next();
              
            file_des = file_src.makeCopy(folder_con);
            file_des.setSpreadsheetTimeZone("Asia/Bangkok");
            file_des.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.NONE);
            file_src.setTrashed(true);
            doc = SpreadsheetApp.open(file_des);
            var doc_id = file_des.getId();
            var doc_url = ff.getUrl(); //+"#gid="+sht_input_id;

        }
        
        //var t_2h_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "2H", t_2h_s, t_2h_e, (t_2h_e - t_2h_s)-t_copy, version_wfl, "Move File Form After Check temp to Folder"]);

        //var t9 = new Date();
        //var t_2m_s = new Date();

        //var doc = convertExcel2Sheets(fileBlob, f_name, folder.getId());
        //Logger.log("Check Write Data History  Start : ")
        var sht = doc.insertSheet("History", doc.getSheets().length).setTabColor('GREEN'); // Change from >> "History" to "History"
        
        var txtRemark = ""
        if ((typeof(e.txtRemark)!= "undefined") && e.txtRemark != "") {
          txtRemark +="[Comment]\r\n "+e.txtRemark;
        }
        if ((typeof(e.txtRef3_)!= "undefined") && e.txtRef3_ != "") {
          txtRemark +="\r\n[Priority Reason :: Start priority from URGENT]\r\n "+ e.txtRef3_;
        }

        var col_h = ["Step", "E-Mail", "Action", "Date/Time", "Comment"];

        sht.getRange("A1").setValue("Step");
        sht.getRange("B1").setValue("E-Mail");
        sht.getRange("C1").setValue("Action");
        sht.getRange("D1").setValue("Date/Time");
        sht.getRange("E1").setValue("Comment");

        sht.getRange("A2").setValue("1");
        sht.getRange("B2").setValue(e.txtIssueBy);
        sht.getRange("C2").setValue("Issued");
        sht.getRange("D2").setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
        sht.getRange("E2").setValue(txtRemark);
   
        sht.autoResizeColumn(1);
        sht.autoResizeColumn(2);
        sht.autoResizeColumn(3);
        sht.autoResizeColumn(4);
        sht.autoResizeColumn(5);


        var file_set = DriveApp.getFileById(doc.getId());
        var file_id_doc =file_set.getId();
        file_set.setOwner("admin@ngkntk-asia.com")
        file_set.setShareableByEditors(true);
        //file_set.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT) // Original
        file_set.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT)   // New From 14-07-2020 Change from domain to private
        
        //file_set.addEditor("s-tatcha@ngkntk-asia.com")                           // New From 14-07-2020 Add permission for person in route only
      //RV1 
      var output = [];
      //RV1
      var optionalArgs = {    sendNotificationEmails: false};  
      e.reviewLevel1.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
      e.reviewLevel2.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
        //return "OK-"+output
      e.reviewLevel3.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
      e.reviewLevel4.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
      e.reviewLevel5.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
      e.reviewLevel6.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
      e.reviewLevel7.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
      e.finalApprove.split(',').forEach(function(q){if (q.toString()!= "")       {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
      e.applicationAdmin.split(',').forEach(function(q){if (q.toString()!= "")   {try{Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false})}catch (e) {output.push(q + "-"+ e);};};})
     // return "OK-"+JSON.stringify(e.applicationAdminCC.split(','));
      e.applicationAdminCC.split(',').forEach(function(q){if (q.toString()!= "") {try{
        Drive.Permissions.insert({value: ''+q.toString()+'',type: 'user',role: 'writer'}, file_id_doc, {    sendNotificationEmails: false}) }catch (e) {output.push(q + "-"+ e);};}
                                                                                ;})

        
        
        //file_set.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT)
      // return "Joe-ok1"
        //-------------------------------------
        //var t7 = new Date();
        var folder2 = folder.createFolder(doc_id);

       //folder2.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDI)
        var user1 = [];
        user1.push("admin@ngkntk-asia.com");

    for(var a in e){
      if (a.indexOf("attFile")>-1 && e[a].length !=0) {
            var zipdoc = folder2.createFile(e[a]);

       }
    }

    } else {
        var doc_id = "";
        var doc_url = "";
    }
       //Logger.log("Upload attach File  Stop : ")
       
       //Logger.log("Check Write Data in Database Sheet Start: ")
       
    var txtId = e.txtId;
    var txtDept = e.txtDepartment;
    //return txtDept;
    var txtDeptCode = txtDept.split("|")[2];
    var txtDepartment = txtDept.split("|");
    var txtRef1 = e.txtRef1;
    var txtRef2 = e.txtRef2;
    var txtRef3 = e.txtRef3;
    var txtIssueDate = e.txtIssueDate;
    var txtIssueBy = e.txtIssueBy;
    if (txtIssueBy != "") {
        user1.push(txtIssueBy);
    }
    var chkCompany = e.chkCompany;
    var chkMaster = e.txtMaster; //e.chkMaster;
   
    var txtPriority = e.priority;
    // change folder name ----------------
    var f_code_name =  Comp_ + "_" + txtDeptCode + "_" + Utilities.formatDate(new Date(), "GMT+7", "yyMMddHHmmssSSS") + "_" + chkMaster
    //return f_code_name
    //t1 = new Date();
     
    var UniqueStr = getUniqueStr();
    
    //ss1.appendRow(["=Row()&\"_\"&\""+f_code_name+"\""]);
     ss1.appendRow([UniqueStr +"_"+f_code_name]); //test 220818
  var namepaths = UniqueStr +"_"+f_code_name;
    //t2 = new Date();
    var qvizQuery = "SELECT * WHERE A like '%" + f_code_name + "%'";
    var qvizURL = 'https://docs.google.com/spreadsheets/d/' + ss1.getParent().getId()
    + '/gviz/tq?tqx=out:json&headers=1&sheet=' + ss1.getName() + '&range=A:A&tq=' + encodeURIComponent(qvizQuery);
    var qvizret = UrlFetchApp.fetch(qvizURL, {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}}).getContentText();
    var qvizjson = JSON.parse(qvizret.replace("/*O_o*/", "").replace("google.visualization.Query.setResponse(", "").slice(0, -2))
     //return qvizjson
    if ((qvizjson.table.rows[0].c[0].v != "undefined") || (qvizjson.table.rows[0].c[0].v !="")) {
      //f_code_name = qvizjson.table.rows[0].c[0].v
        f_code_name =  UniqueStr; //test 220818
    } else {
      return 9;
    }
    //f_code_name =  UniqueStr; //test 220818
   // t3 = new Date();
    //Logger.log("Write Data : " + (t2 - t1) )
   // Logger.log("Read ID Data : " + (t3 - t2) )
    //  Logger.log(f_code_name)
    //ss1.appendRow([f_code_name]);
    //rows = rg.getLastRow() + 1;
     rows =  f_code_name.split("_")[0]
    //var f_code_name = rows + "_" +f_code_name
     
     
    doc.setName(namepaths);
    file_des.setName(namepaths)
    folder2.setName(namepaths);
    if (Convert_flg != "true") {
        file_form.setName(namepaths);
    }

    //var t_2j_e = new Date();
    //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "2J", t_2j_s, t_2j_e, t_2j_e - t_2j_s, version_wfl, "Upload File Support  & Unzip"]);

    //var t_2k_s = new Date();

    //ADDED FOR V2
    var chkReview1 = e.reviewLevel1;
    var STReview1 = e.STreviewLevel1;
    var RMReview1 = (STReview1 == "1" ? "" : "Remove-" + e.CommentreviewLevel1);
    var PTHReview1 = (STReview1 == "1" ? "" : doc_url);
    var IDReview1 = (STReview1 == "1" ? "" : doc_id);
    var PCReview1 = (STReview1 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
    //if (chkReview1    != "") {user1.push(chkReview1);}

    var chkReview2 = e.reviewLevel2;
    var STReview2 = e.STreviewLevel2;
    var RMReview2 = (STReview2 == "1" ? "" : "Remove-" + e.CommentreviewLevel2);
    var PTHReview2 = (STReview2 == "1" ? "" : doc_url);
    var IDReview2 = (STReview2 == "1" ? "" : doc_id);
    var PCReview2 = (STReview2 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
    //if (chkReview2    != "") {user1.push(chkReview2);}

    var chkReview3 = e.reviewLevel3;
    var STReview3 = e.STreviewLevel3;
    var RMReview3 = (STReview3 == "1" ? "" : "Remove-" + e.CommentreviewLevel3);
    var PTHReview3 = (STReview3 == "1" ? "" : doc_url);
    var IDReview3 = (STReview3 == "1" ? "" : doc_id);
    var PCReview3 = (STReview3 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
    //if (chkReview3    != "") {user1.push(chkReview3);}

    var chkReview4 = e.reviewLevel4;
    var STReview4 = e.STreviewLevel4;
    var RMReview4 = (STReview4 == "1" ? "" : "Remove-" + e.CommentreviewLevel4);
    var PTHReview4 = (STReview4 == "1" ? "" : doc_url);
    var IDReview4 = (STReview4 == "1" ? "" : doc_id);
    var PCReview4 = (STReview4 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
    //if (chkReview4    != "") {user1.push(chkReview4);}

    var chkReview5 = e.reviewLevel5;
    var STReview5 = e.STreviewLevel5;
    var RMReview5 = (STReview5 == "1" ? "" : "Remove-" + e.CommentreviewLevel5);
    var PTHReview5 = (STReview5 == "1" ? "" : doc_url);
    var IDReview5 = (STReview5 == "1" ? "" : doc_id);
    var PCReview5 = (STReview5 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
    //if (chkApprove1   != "") {user1.push(chkApprove1);}

    var chkReview6 = e.reviewLevel6;
    var STReview6 = e.STreviewLevel6;
    var RMReview6 = (STReview6 == "1" ? "" : "Remove-" + e.CommentreviewLevel6);
    var PTHReview6 = (STReview6 == "1" ? "" : doc_url);
    var IDReview6 = (STReview6 == "1" ? "" : doc_id);
    var PCReview6 = (STReview6 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
    //if (chkReview6    != "") {user1.push(chkReview6);}

    var chkReview7 = e.reviewLevel7;
    var STReview7 = e.STreviewLevel7;
    var RMReview7 = (STReview7 == "1" ? "" : "Remove-" + e.CommentreviewLevel7);
    var PTHReview7 = (STReview7 == "1" ? "" : doc_url);
    var IDReview7 = (STReview7 == "1" ? "" : doc_id);
    var PCReview7 = (STReview7 == "1" ? "" : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
    //if (chkReview7    != "") {user1.push(chkReview7);}

    var chkApprove2 = e.finalApprove;
    //if (chkApprove2   != "") {user1.push(chkApprove2);}
    var txtRegister = e.applicationAdmin;
    //if (txtRegister   != "") {user1.push(txtRegister);}
    var txtRegisterCC = e.applicationAdminCC;
    //if (txtRegisterCC != "") {user1.push(txtRegisterCC);}

    //return   STReview1+":"+STReview2+":"+STReview3+":"+STReview4+":"+STReview5+":"+STReview6+":"+STReview7+":"


    var bi;
    //END ADDED FOR V2

    //////SUBMIT REQUEST STEP
    //var t6 = new Date();
    /*
    pk = txtRef1 + ":" + chkCompany + '_' + chkMaster + '_' + Utilities.formatDate(new Date(), "GMT+7","yyMMddHHmmss");

    if (txtRef1 == "-") {
    pk = chkCompany + '_' + txtDepartment[1] + '_' + chkMaster + '_' + Utilities.formatDate(new Date(), "GMT+7","yyMMddHHmmss");
    } else {
    pk = txtRef1 + ":" + chkCompany + '_' + txtDepartment[1] + '_' + chkMaster + '_' + Utilities.formatDate(new Date(), "GMT+7","yyMMddHHmmss");
    }
     */
    //ss2 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
  
    //pk = f_code_name; 
  pk = namepaths;
 /*
  ss1.appendRow([
            　　　 pk,
            chkCompany,
            'Issue',
            chkMaster,
            txtRemark,
            txtIssueBy,
            txtIssueDate,
            doc_url,
            doc_id,
            Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")
            //ADDED FOR V2
            //[comment by reviewer],]reviewer email],[Attach Path],[Attcch ID],[process date & time]
        , RMReview5, chkReview5, PTHReview5, IDReview5, PCReview5, '', chkApprove2, '', '', '', '',
      txtRegister, '', '', '', '', RMReview1, chkReview1, PTHReview1, IDReview1, PCReview1, RMReview2,
      chkReview2, PTHReview2, IDReview2, PCReview2, RMReview3, chkReview3, PTHReview3, IDReview3, PCReview3,
      RMReview4, chkReview4, PTHReview4, IDReview4, PCReview4, txtRegisterCC, RMReview6, chkReview6, PTHReview6,
      IDReview6, PCReview6, RMReview7, chkReview7, PTHReview7, IDReview7, PCReview7, txtRef1, txtRef2, txtRef3
            //ADDED FOR RELEASE 02.052018

        ]);
        */

 //return  "A"+pk+":BH"+pk;
   var data_input=[[
            　　　 pk,
            chkCompany,
            'Issue',
            chkMaster,
            txtRemark,
            txtIssueBy,
            txtIssueDate,
            doc_url,
            doc_id,
            Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")
            //ADDED FOR V2
            //[comment by reviewer],]reviewer email],[Attach Path],[Attcch ID],[process date & time]
        , RMReview5, chkReview5, PTHReview5, IDReview5, PCReview5, '', chkApprove2, '', '', '', '',
      txtRegister, '', '', '', '', RMReview1, chkReview1, PTHReview1, IDReview1, PCReview1, RMReview2,
      chkReview2, PTHReview2, IDReview2, PCReview2, RMReview3, chkReview3, PTHReview3, IDReview3, PCReview3,
      RMReview4, chkReview4, PTHReview4, IDReview4, PCReview4, txtRegisterCC, RMReview6, chkReview6, PTHReview6,
      IDReview6, PCReview6, RMReview7, chkReview7, PTHReview7, IDReview7, PCReview7, txtRef1, txtRef2, txtRef3,
      '', '', '', '','', '', '', '','', '', '', '','', '', '', '','', '', '', '','', '', '', '','', '', '', '',txtPriority
            //ADDED FOR RELEASE 02.052018

        ]];
  var temprow = findRow(rows); // test 202218
  
  rows = temprow; // test 202218
  ss1.getRange("A" +  rows +":CK"+rows).setValues(data_input)
  //ss1.getRange("A"+pk+":BH"+pk).setValues()

/// INPUT CULUMN CH
 var min = 1;
 var max = 100;
 const now =  Utilities.formatDate(new Date(), "GMT+7:00", "yyyyMMddHHmmss") + Math.random() * (max - min) + min;
 var result = parseInt(now).toString();
////
    //var t5 = new Date();
  var formulas = [
    ['=IF(IFERROR(FIND(\"Rejected\",C' + rows + '), -1) >= 0, \"\", CONCATENATE(CH' + rows + ',"|",BU' + rows + ',"|",CF' + rows + ',"|",CG' + rows + '))',
    '',
    '=IF(AE' + rows + '="",0,1)',
    '=IF(AJ' + rows + '="",0,1)',
    '=IF(AO' + rows + '="",0,1)',
    '=IF(AT' + rows + '="",0,1)',
    '=IF(AZ' + rows + '="",0,1)',
    '=IF(BE' + rows + '="",0,1)',
    '=IF(O' + rows + '="",0,1)',
    '=IF(T' + rows + '="",0,1)',
    '=IF(Y' + rows + '="",0,1)',
    '=SUMPRODUCT($BK$1:$BS$1,BK' + rows + ':BS' + rows + ')',
    '=IF(BT' + rows + '<1,"0",IF(BT' + rows + '<2,"1",IF(BT' + rows + '<4,"1A",IF(BT' + rows + '<8,"1B",IF(BT' + rows + '<16,"1C",IF(BT' + rows + '<32,"1D",IF(BT' + rows + '<64,"2",IF(BT' + rows + '<128,"2A",IF(BT' + rows + '<256,"2B",IF(BT' + rows + '<512,"3","9"))))))))))',
    '=IF(AND(AB' + rows + '<>"",AE' + rows + '=""),1,0)',
    '=IF(AND(AG' + rows + '<>"",AJ' + rows + '=""),1,0)',
    '=IF(AND(AL' + rows + '<>"",AO' + rows + '=""),1,0)',
    '=IF(AND(AQ' + rows + '<>"",AT' + rows + '=""),1,0)',
    '=IF(AND(AW' + rows + '<>"",AZ' + rows + '=""),1,0)',
    '=IF(AND(BB' + rows + '<>"",BE' + rows + '=""),1,0)',
    '=IF(AND(L' + rows + '<>"",O' + rows + '=""),1,0)',
    '=IF(AND(Q' + rows + '<>"",T' + rows + '=""),1,0)',
    '=IF(AND(V' + rows + '<>"",Y' + rows + '=""),1,0)',
    '=SUMPRODUCT($BV$1:$CD$1,BV' + rows + ':CD' + rows + ')',
    '=IF(CE' + rows + '>=512,"1A",IF(CE' + rows + '>=256,"1B",IF(CE' + rows + '>=128,"1C",IF(CE' + rows + '>=64,"1D",IF(CE' + rows + '>=32,"2",IF(CE' + rows + '>=16,"2A",IF(CE' + rows + '>=8,"2B",IF(CE' + rows + '>=4,"3",IF(CE' + rows + '>=2,"4","9")))))))))',
    '=IF(CF' + rows + '="1A",AB' + rows + ',IF(CF' + rows + '="1B",AG' + rows + ',IF(CF' + rows + '="1C",AL' + rows + ',IF(CF' + rows + '="1D",AQ' + rows + ',IF(CF' + rows + '="2",L' + rows + ',IF(CF' + rows + '="2A",AW' + rows
         + ',IF(CF' + rows + '="2B",BB' + rows + ',IF(CF' + rows + '="3",Q' + rows + ',IF(CF' + rows + '="4",V' + rows + ',"")))))))))',
    result]
  ];
    ss1.getRange("BI" + rows+":CH"+rows).setFormulas(formulas)
  
    //var t_2k_e = new Date();
    //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "2K", t_2k_s, t_2k_e, t_2k_e - t_2k_s, version_wfl, "Write to Master Sheet After Upload"]);
    //   //Logger.log("Check Write Data in Database Sheet Stop: ")
    //          //Logger.log("Check Send Mail Start: ")
    //ADDED FOR V2
    //var t_2l_s = new Date();
    
    //Check Urgent
    urgt = ""
    if (txtPriority==0) {urgt="[!!!URGENT!!!]"}
    
    mlsubj = "WFL <" + urgt + "You're Reviewer>"
        if (chkReview1 != "" && STReview1 == "1") {
            ml = chkReview1; //NEXT APPROVAL is Review #1
            rt = "1A";
        } else if (chkReview2 != "" && STReview2 == "1") {
            ml = chkReview2; //NEXT APPROVAL is Review #2
            rt = "1B";
        } else if (chkReview3 != "" && STReview3 == "1") {
            ml = chkReview3; //NEXT APPROVAL is Review #3
            rt = "1C";
        } else if (chkReview4 != "" && STReview4 == "1") {
            ml = chkReview4; //NEXT APPROVAL is Review #4
            rt = "1D";
        } else if (chkReview5 != "" && STReview5 == "1") {
            ml = chkReview5; //NEXT APPROVAL is Review #5
            rt = "2";
        } else if (chkReview6 != "" && STReview6 == "1") {
            ml = chkReview6; //NEXT APPROVAL is Review #5
            rt = "2A";
        } else if (chkReview7 != "" && STReview7 == "1") {
            ml = chkReview7; //NEXT APPROVAL is Review #5
            rt = "2B";
        }

        else if (chkApprove2 != "") {
            ml = chkApprove2; //NEXT APPROVAL is Approver.
            rt = "3";
            mlsubj = "WFL <" + urgt + "You're Approver>"
        }


       // var t4 = new Date();
    //START WRITE LOG

    ss3 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
    //ss3.getSheetByName("Logs").activate();
    ss3.getSheetByName("Logs").appendRow([
            pk,
            chkCompany,
            'Issue',
            chkMaster,
            txtRemark,
            txtIssueBy,
            txtIssueDate,
            doc_url,
            doc_id,
            Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"), WEB_PATH + "?id=" + UniqueStr + "&rt=" + rt
            //ADDED FOR RELEASE 02.052018
        , txtRef1, txtRef2, txtRef3
            //ADDED FOR RELEASE 02.052018
        ]);

    //END FOR WRITE LOG

//var t3 = new Date();

    //SEND EMAIL TO NEXT STEP (REVIEW)
    ml = "admin@ngkntk-asia.com,"+ml;
    MailApp.sendEmail(ml.replace(/\s/g, ""), "" + mlsubj + " - " + txtRef1 + ':' + pk, "You have received Application from WorkFlow Launcher.\r\n"
         + "Please confirm the following link.\r\n\r\n＜Link＞\r\n"
         + WEB_PATH + "?id=" + (UniqueStr) + "&rt=" + rt + "&fr=" + 0);

    //SEND EMAIL TO ISSUE PERSON (ORIGINATOR)

    MailApp.sendEmail(txtIssueBy, "" + "WFL <"+ urgt + txtRef1 + ":" + pk + ">", "You have received this email because you're submitted WorkFlow Launcher (" + chkMaster + ").\r\n"
         + "You can click on the link to view latest document status.\r\n\r\n＜Link＞\r\n"
         //+ WEB_PATH + "?id=" + (rows) + "&rt=" + 0 + "&fr=" + 0);
         + WEB_PATH + "?id=" + (UniqueStr) + "&rt=" + "x" + "&fr=" + 0);
    //var t_2l_e = new Date();
    //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "2L", t_2l_s, t_2l_e, t_2l_e - t_2l_s, version_wfl, "Send Mail to Next Process"]);
    //          //Logger.log("Check Send Mail Stop: ")
    //var t2 = new Date();
  /*  
  var diff = Math.floor((t2 - t1) / (1000.0)); //seconds
    var diff_mail = Math.floor((t2 - t3) / (1000.0)); //seconds
    var diff_Log = Math.floor((t3 - t4) / (1000.0)); //seconds
    var diff_formular = Math.floor((t4 - t5) / (1000.0)); //seconds
    var diff_wrt = Math.floor((t5 - t6) / (1000.0)); //seconds
    var diff_upform = Math.floor((t9 - t1) / (1000.0)); //seconds
 */
 //Logger.log("Stop : ")
    return 1;

}

//*******************************************************************************
// Added for V2 : Route Change
// 2017/12/26 P.Winyou (ANTK)
// FUNCTION TO CHECK CURRENT AUTHORIZE
//*******************************************************************************

function getCurrentAuthorize(pid, pRt) {

    var currAuthor,
    aUser,
    cUser,
    compareUser,
    ss,
    values;

    ss = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);

    values = ss.getSheetByName(INPUT_DATA).getRange(pid + ":" + pid).getValues();

    if (values[0][0] != "") {

        compareUser = Session.getActiveUser().getEmail().replace(/\s/g, ""); //To fixed case non-unique name of VNGK.
        cUser = Session.getActiveUser().getEmail().replace(/\s/g, "");

        switch (pRt) {
        case "1A":
            //if ((values[pid][27] == Session.getActiveUser().getEmail()) && (values[pid][30] == "")) { currAuthor = values[pid][27]; } else { currAuthor = "not authorized"; }
            aUser = "," + values[0][27].replace(/\s/g, "");

            //MailApp.sendEmail("p-winyou@ngkntk-asia.com", "cUser="+cUser+"     aUser=" + aUser,"Result=" + (aUser.indexOf(cUser) > -1))
            if ((aUser.indexOf(compareUser) > -1) && (values[0][30] == "")) {
                currAuthor = cUser;
            } else {
                currAuthor = "not authorized";
            }
            break;

        case "1B":
            //if ((values[pid][32] == Session.getActiveUser().getEmail()) && (values[pid][35] == "")) { currAuthor = values[pid][32]; } else { currAuthor = "not authorized"; }
            aUser = "," + values[0][32];
            if ((aUser.indexOf(compareUser) > -1) && (values[0][35] == "")) {
                currAuthor = cUser;
            } else {
                currAuthor = "not authorized";
            }
            break;

        case "1C":
            //if ((values[pid][37] == Session.getActiveUser().getEmail()) && (values[pid][40] == "")) { currAuthor = values[pid][37]; } else { currAuthor = "not authorized"; }
            aUser = "," + values[0][37];
            if ((aUser.indexOf(compareUser) > -1) && (values[0][40] == "")) {
                currAuthor = cUser;
            } else {
                currAuthor = "not authorized";
            }
            break;

        case "1D":
            //if ((values[pid][42] == Session.getActiveUser().getEmail()) && (values[pid][45] == "")) { currAuthor = values[pid][42]; } else { currAuthor = "not authorized"; }
            aUser = "," + values[0][42];
            if ((aUser.indexOf(compareUser) > -1) && (values[0][45] == "")) {
                currAuthor = cUser;
            } else {
                currAuthor = "not authorized";
            }
            break;

        case "2":
            //if ((values[pid][11] == Session.getActiveUser().getEmail()) && (values[pid][14] == "")) { currAuthor = values[pid][11]; } else { currAuthor = "not authorized"; }
            aUser = "," + values[0][11];
            if ((aUser.indexOf(compareUser) > -1) && (values[0][14] == "")) {
                currAuthor = cUser;
            } else {
                currAuthor = "not authorized";
            }
            break;

            /****************************************************************/

        case "2A":
            //      if ((values[pid][48] == Session.getActiveUser().getEmail()) && (values[pid][51] == "")) { currAuthor = values[pid][48]; } else { currAuthor = "not authorized"; }
            aUser = "," + values[0][48];
            if ((aUser.indexOf(compareUser) > -1) && (values[0][51] == "")) {
                currAuthor = cUser;
            } else {
                currAuthor = "not authorized";
            }
            break;

        case "2B":
            //      if ((values[pid][53] == Session.getActiveUser().getEmail()) && (values[pid][56] == "")) { currAuthor = values[pid][53]; } else { currAuthor = "not authorized"; }
            aUser = "," + values[0][53];
            if ((aUser.indexOf(compareUser) > -1) && (values[0][56] == "")) {
                currAuthor = cUser;
            } else {
                currAuthor = "not authorized";
            }
            break;

            /****************************************************************/

        case "3":
            //if ((values[pid][16] == Session.getActiveUser().getEmail()) && (values[pid][19] == "")) { currAuthor = values[pid][16]; } else { currAuthor = "not authorized"; }
            aUser = "," + values[0][16];
            ////Logger.log(aUser + ":" + compareUser + " = " + aUser.indexOf("r-thawatchai@ngkntk-asia.com"))
            if ((aUser.indexOf(compareUser) > -1) && (values[0][19] == "")) {
                currAuthor = cUser;
            } else {
                currAuthor = "not authorized";
            }
            break;

        case "4":
            //if ((values[pid][21] == Session.getActiveUser().getEmail()) && (values[pid][24] == "")) { currAuthor = values[pid][21]; } else { currAuthor = "not authorized"; }
            aUser = "," + values[0][21];
            if ((aUser.indexOf(compareUser) > -1) && (values[0][24] == "")) {
                currAuthor = cUser;
            } else {
                currAuthor = "not authorized";
            }
            break;

        }

    } else {

        currAuthor = Session.getActiveUser().getEmail();

    }

    //START WRITE LOG
/*
    ss3 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
    ss3.getSheetByName("Logs").activate();
    ss3.getSheetByName("Logs").appendRow([
    　　　 pid,
    pRt,
    values[pid][27],
    values[pid][30],
    Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")
    ]);
*/
     //END FOR WRITE LOG


    return currAuthor;

}

//*******************************************************************************
// Added for V2 : Route Change
// 2017/12/26 P.Winyou (ANTK)
// FUNCTION TO GET HISTORY FROM
//*******************************************************************************

function getHistoryLog(pid) {
    //console.log(pid);
    //event.parameter.id
    //event.parameter.rt
    var status,j,nxt,plus,pmsg,pstatus,
    paction,ss,rg,rg_pos,rg_dept,strHTML,rows,values;

    //pid = "DB_PDOKD18062301:TNGK_PU_DB Sheet - Items_180627130751"
    //JOE-----------
    strHTML = "";
    //var t1_check = new Date();
    //Joe-----------

    ss = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
    rg = ss.getSheetByName(INPUT_DATA).getRange(pid + ":" + pid);

    rows = rg.getLastRow();
    values = rg.getValues();
  
    rg_dept = ss.getSheetByName("Department");
    var dept = rg_dept.getRange("A:C").getValues().filter(function(row){
      if (row[0] === values[0][1] && row[1] === values[0][0].toString().split("_")[2] ) {      return row;    }
    });
    dept = typeof(dept)!="undefined"? dept:[];
  //return JSON.stringify(dept)
    rg_pos = ss.getSheetByName("Routing-Pos");
    var pos = rg_pos.getRange("A:N").getValues().filter(function(row){
      if (row[1] === values[0][1] && row[2] === dept[0][2] && row[3] === values[0][3]) {      return row;    }
    });
    pos = typeof(pos)!="undefined"? pos:[];
    //return JSON.stringify(pos)
  
  var timezone_ = ss.getSheetByName("Timezone").getRange("A:C").getValues().filter(function (dataRow) {return dataRow[0] ==values[0][1];});
  timezone_ = timezone_ ? timezone_ :ss.getSheetByName("Timezone").getRange("A:C").getValues().filter(function (dataRow) {return dataRow[0] =="DEFAULT";});
    strHTML += "<div  class=\"animated  fadeInLeft delay-1s \" ><I><B><br>Request ID# <span style=\"background-color:#f3fe0063\">[" + values[0][85] + "]</span> - " + values[0][0] + "</B></I><br><br><TABLE bgcolor=\"#e8f8f5\" width=\"100%\">";
    strHTML = strHTML + "<tr><th width='2%'>Step</th>"
         + "<th width='25%' >&nbsp;&nbsp;E-Mail</th>"
         + "<th width='15%' >Action</th>"
         + "<th width='7%'  >Date/Time</th>"
         + "<th>Comment/Urgent Reason</th></tr>";
    j = 0;
//(parseInt(pid)) for line 1785
    nxt = j;
    //j++;
    pmsg = "((Pending))";

    pstatus = "Issued";

    if (j < 1) {

        nxt++;
        strHTML = strHTML + "<tr><td>" + nxt + "</td>";
        strHTML = strHTML + "<td>" +values[0][5].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";
        ////Logger.log(values[0][5])
        if (values[0][6] != "") {
            strHTML = strHTML + "<td>" + pstatus + "</td>"
        } else {
            strHTML = strHTML + "<td> ((Pending)) </td>"
        };
        if (values[0][9] != "") {
            strHTML = strHTML + "<td>" + Utilities.formatDate(values[0][9], timezone_[0][2], "dd-MMM-yy HH:mm") + "</td>"
        } else {
            strHTML = strHTML + "<td> - </td>"
        };
        strHTML = strHTML + "<td>" + values[0][4].replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";
        strHTML = strHTML + "</tr>";
    }

    //pstatus = tvalues[pid][2];

    paction = values[0][2].indexOf("Rejected by Reviewer1");
    if (paction > -1)  {
        status_wfl = false
    }

    if (paction == -1) {
        pstatus = "Signed By Reviewer";
    } else {
        pstatus = "Rejected by Reviewer";
    }

    //if  ((paction > -1) && ((i+1) == j) ) { pstatus = values[pid][2] + nxt; } else {  pstatus = "Signed By Reviewer"; }

    if ((j < 2) && (values[0][27] != "")) {

        nxt++;
        strHTML = strHTML + "<tr><td>" + nxt + "</td>";
      strHTML = strHTML + "<td>" + (typeof(pos[0])!="undefined"? (pos[0][4]!=""?pos[0][4]+" - ":""):"")+values[0][27].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";

        if (status_wfl == true) {
            if (values[0][27] != "") {
                pmsg = "<td><font color=red> ((Pending Review)) </font></td>";
            } else {
                pmsg = "<td> - </td>";
            }
        } else {
            pmsg = "<td></td>";
        }
        if (values[0][26].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td>Skip By Issuer (" + values[0][26].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + ") </td>";
            //
        } else {
            if (values[0][30] != "") {
                strHTML = strHTML + "<td>" + pstatus + "1</td>"
            } else {
                strHTML = strHTML + pmsg
            };
        }
        if (values[0][30] != "") {
            strHTML = strHTML + "<td>" + Utilities.formatDate(values[0][30], timezone_[0][2], "dd-MMM-yy HH:mm") + "</td>"
        } else {
            strHTML = strHTML + pmsg
        };
        if (values[0][26].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td></td>";
        } else {
            strHTML = strHTML + "<td>" + values[0][26].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";
        }
        strHTML = strHTML + "</tr>";
        var dd = "";
        var dis_rt = "disabled";
        if (values[0][30] != "" && values[0][26].indexOf("Remove-") < 0) {
            ////Logger.log(values[0][26].indexOf("Remove-"))
            dd = Utilities.formatDate(values[0][30], timezone_[0][2], "dd/MM/yy");
            dis_rt = "";
            txt_rt += "<option value=\"1A\" " + dis_rt + " >[Reviewer1]  - " + values[0][27] + " - " + dd + "</option>";
        }
    }
    paction = values[0][2].indexOf("Rejected by Reviewer2");
    if (paction > -1) {
        status_wfl = false
    }
    
    if (paction == -1) {
        pstatus = "Signed By Reviewer";
    } else {
        pstatus = "Rejected by Reviewer";
    }
    if ((j < 3) && (values[0][32] != "")) {

        nxt++;
        strHTML = strHTML + "<tr><td>" + nxt + "</td>";
        strHTML = strHTML + "<td>" + (typeof(pos[0])!="undefined"? (pos[0][5]!=""?pos[0][5]+" - ":""):"")+values[0][32].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";

        if (status_wfl == true) {
            if (values[0][32] != "") {
                pmsg = "<td><font color=red> ((Pending Review))</font></td>";
            } else {
                pmsg = "<td> - </td>";
            }
        } else {
            pmsg = "<td></td>";
        }
        //strHTML = strHTML + "<td>Review</td>";
        if (values[0][31].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td>Skip By Issuer(" + values[0][31].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + ") </td>";

        } else {
            if (values[0][35] != "") {
                strHTML = strHTML + "<td>" + pstatus + "2</td>"
            } else {
                strHTML = strHTML + pmsg
            };
        }
        if (values[0][35] != "") {
            strHTML = strHTML + "<td>" + Utilities.formatDate(values[0][35], timezone_[0][2], "dd-MMM-yy HH:mm") + "</td>"
        } else {
            strHTML = strHTML + pmsg
        };
        if (values[0][31].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td></td>"
        } else {
            strHTML = strHTML + "<td>" + values[0][31].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";
        }
        strHTML = strHTML + "</tr>";

        var dd = "";
        var dis_rt = "disabled";
        if (values[0][35] != "" && values[0][31].indexOf("Remove-") < 0) {
            dd = Utilities.formatDate(values[0][35], timezone_[0][2], "dd/MM/yy");
            dis_rt = "";
            txt_rt += "<option value=\"1B\" " + dis_rt + " >[Reviewer2]  - " + values[0][32] + " - " + dd + "</option>";
        }
    }
      paction = values[0][2].indexOf("Rejected by Reviewer3");
    if (paction > -1) {
        status_wfl = false
    }

    if (paction == -1) {
        pstatus = "Signed By Reviewer";
    } else {
        pstatus = "Rejected by Reviewer";
    }
    if ((j < 4) && (values[0][37] != "")) {

        nxt++;
        strHTML = strHTML + "<tr><td>" + nxt + "</td>";
        strHTML = strHTML + "<td>" + (typeof(pos[0])!="undefined"? (pos[0][6]!=""?pos[0][6]+" - ":""):"")+values[0][37].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";

        if (status_wfl == true) {
            if (values[0][37] != "") {
                pmsg = "<td><font color=red> ((Pending Review)) </font></td>";
            } else {
                pmsg = "<td> - </td>";
            }
        } else {
            pmsg = "<td></td>";
        }
        //strHTML = strHTML + "<td>Review</td>";
        if (values[0][36].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td>Skip By Issuer (" + values[0][36].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + ") </td>";
        } else {
            if (values[0][40] != "") {
                strHTML = strHTML + "<td>" + pstatus + "3</td>"
            } else {
                strHTML = strHTML + pmsg
            };
        }
        if (values[0][40] != "") {
            strHTML = strHTML + "<td>" + Utilities.formatDate(values[0][40], timezone_[0][2], "dd-MMM-yy HH:mm") + "</td>"
        } else {
            strHTML = strHTML + pmsg
        };
        if (values[0][36].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td></td>"
        } else {
            strHTML = strHTML + "<td>" + values[0][36].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";
        }
       // strHTML = strHTML + "<td>" + values[0][36].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;') + "</td>";
        strHTML = strHTML + "</tr>";

        var dd = "";
        var dis_rt = "disabled";
        if (values[0][40] != "" && values[0][36].indexOf("Remove-") < 0) {
            dd = Utilities.formatDate(values[0][40],timezone_[0][2], "dd/MM/yy");
            dis_rt = "";
            txt_rt += "<option value=\"1C\" " + dis_rt + " >[Reviewer3]  - " + values[0][37] + " - " + dd + "</option>";
        }
    }
    paction = values[0][2].indexOf("Rejected by Reviewer4");
    if (paction > -1) {
        status_wfl = false
    }

    if (paction == -1) {
        pstatus = "Signed By Reviewer";
    } else {
        pstatus = "Rejected by Reviewer";
    }
    if ((j < 5) && (values[0][42] != "")) {

        nxt++;
        strHTML = strHTML + "<tr><td>" + nxt + "</td>";
        strHTML = strHTML + "<td>" + (typeof(pos[0])!="undefined"? (pos[0][7]!=""?pos[0][7]+" - ":""):"")+values[0][42].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";

        if (status_wfl == true) {
            if (values[0][42] != "") {
                pmsg = "<td><font color=red> ((Pending Review)) </font></td>";
            } else {
                pmsg = "<td> - </td>";
            }
        } else {
            pmsg = "<td></td>";
        }
        //strHTML = strHTML + "<td>Review</td>";
        if (values[0][41].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td>Skip By Issuer (" + values[0][41].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + ") </td>";
        } else {
            if (values[0][45] != "") {
                strHTML = strHTML + "<td>" + pstatus + "4</td>"
            } else {
                strHTML = strHTML + pmsg
            };
        }
        if (values[0][45] != "") {
            strHTML = strHTML + "<td>" + Utilities.formatDate(values[0][45], timezone_[0][2], "dd-MMM-yy HH:mm") + "</td>"
        } else {
            strHTML = strHTML + pmsg
        };
        if (values[0][41].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td></td>"
        } else {
            strHTML = strHTML + "<td>" + values[0][41].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";
        }
        //strHTML = strHTML + "<td>" + values[0][41].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;') + "</td>";
        strHTML = strHTML + "</tr>";

        var dd = "";
        var dis_rt = "disabled";
        if (values[0][45] != "" && values[0][41].indexOf("Remove-") < 0) {
            dd = Utilities.formatDate(values[0][45], timezone_[0][2], "dd/MM/yy");
            dis_rt = "";
            txt_rt += "<option value=\"1D\" " + dis_rt + " >[Reviewer4]  - " + values[0][42] + " - " + dd + "</option>";
        }
    }
    paction = values[0][2].indexOf("Rejected by Reviewer5");
    if (paction > -1) {
        status_wfl = false
    }

    if (paction == -1) {
        pstatus = "Signed By Reviewer";
    } else {
        pstatus = "Rejected by Reviewer";
    }
    if ((j < 6) && (values[0][11] != "")) {

        nxt++;
        strHTML = strHTML + "<tr><td>" + nxt + "</td>";
        strHTML = strHTML + "<td>" + (typeof(pos[0])!="undefined"? (pos[0][8]!=""?pos[0][8]+" - ":""):"")+ values[0][11].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";

        if (status_wfl == true) {
            if (values[0][11] != "") {
                pmsg = "<td><font color=red> ((Pending Review)) </font></td>";
            } else {
                pmsg = "<td> - </td>";
            }
        } else {
            pmsg = "<td></td>";
        }
        //strHTML = strHTML + "<td>Approve</td>";
        if (values[0][10].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td>Skip By Issuer (" + values[0][10].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + ") </td>";

        } else {
            if (values[0][14] != "") {
                strHTML = strHTML + "<td>" + pstatus + "5</td>"
            } else {
                strHTML = strHTML + pmsg
            };
        }
        if (values[0][14] != "") {
            strHTML = strHTML + "<td>" + Utilities.formatDate(values[0][14], timezone_[0][2], "dd-MMM-yy HH:mm") + "</td>"
        } else {
            strHTML = strHTML + pmsg
        };
        if (values[0][10].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td></td>"
        } else {
            strHTML = strHTML + "<td>" + values[0][10].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";
        }
       // strHTML = strHTML + "<td>" + values[0][10].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;') + "</td>";
        strHTML = strHTML + "</tr>";

        var dd = "";
        var dis_rt = "disabled";
        if (values[0][14] != "" && values[0][10].indexOf("Remove-") < 0) {
            dd = Utilities.formatDate(values[0][14], timezone_[0][2], "dd/MM/yy");
            dis_rt = "";
            txt_rt += "<option value=\"2\" " + dis_rt + " >[Reviewer5]  - " + values[0][11] + " - " + dd + "</option>";
        }
    }

    /******************************************************************/
    paction = values[0][2].indexOf("Rejected by Reviewer6");
    if (paction > -1) {
        status_wfl = false
    }

    if (paction == -1) {
        pstatus = "Signed By Reviewer";
    } else {
        pstatus = "Rejected by Reviewer";
    }
    if ((j < 7) && (values[0][48] != "")) {

        nxt++;
        strHTML = strHTML + "<tr><td>" + nxt + "</td>";
        strHTML = strHTML + "<td>" + (typeof(pos[0])!="undefined"? (pos[0][12]!=""?pos[0][12]+" - ":""):"")+ values[0][48].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";

        if (status_wfl == true) {
            if (values[0][48] != "") {
                pmsg = "<td><font color=red> ((Pending Review)) </font></td>";
            } else {
                pmsg = "<td> - </td>";
            }
        } else {
            pmsg = "<td></td>";
        }
        //strHTML = strHTML + "<td>Final Approve</td>";
        if (values[0][47].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td>Skip By Issuer (" + values[0][47].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + ") </td>";
        } else {
            if (values[0][51] != "") {
                strHTML = strHTML + "<td>" + pstatus + "6</td>"
            } else {
                strHTML = strHTML + pmsg
            };
        }
        if (values[0][51] != "") {
            strHTML = strHTML + "<td>" + Utilities.formatDate(values[0][51], timezone_[0][2], "dd-MMM-yy HH:mm") + "</td>"
        } else {
            strHTML = strHTML + pmsg
        };
        if (values[0][47].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td></td>"
        } else {
            strHTML = strHTML + "<td>" + values[0][47].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";
        }
        //strHTML = strHTML + "<td>" + values[0][47].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;') + "</td>";
        strHTML = strHTML + "</tr>";

        var dd = "";
        var dis_rt = "disabled";
        if (values[0][51] != "" && values[0][47].indexOf("Remove-") < 0) {
            dd = Utilities.formatDate(values[0][51], timezone_[0][2], "dd/MM/yy");
            dis_rt = "";
            txt_rt += "<option value=\"2A\" " + dis_rt + " >[Reviewer6]  - " + values[0][48] + " - " + dd + "</option>";
        }
    }
    paction = values[0][2].indexOf("Rejected by Reviewer7");
    if (paction > -1) {
        status_wfl = false
    }

    if (paction == -1) {
        pstatus = "Signed By Reviewer";
    } else {
        pstatus = "Rejected by Reviewer";
    }
    if ((j < 8) && (values[0][53] != "")) {
        nxt++;
        strHTML = strHTML + "<tr><td>" + nxt + "</td>";
        strHTML = strHTML + "<td>" + (typeof(pos[0])!="undefined"? (pos[0][13]!=""?pos[0][13]+" - ":""):"")+  values[0][53].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";

        if (status_wfl == true) {
            if (values[0][53] != "") {
                pmsg = "<td><font color=red> ((Pending Review)) </font></td>";
            } else {
                pmsg = "<td> - </td>";
            }
        } else {
            pmsg = "<td></td>";
        }
        //strHTML = strHTML + "<td>Final Approve</td>";
        if (values[0][52].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td>Skip By Issuer (" + values[0][52].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + ") </td>";
        } else {
            if (values[0][56] != "") {
                strHTML = strHTML + "<td>" + pstatus + "7</td>"
            } else {
                strHTML = strHTML + pmsg
            };
        }
        if (values[0][56] != "") {
            strHTML = strHTML + "<td>" + Utilities.formatDate(values[0][56], timezone_[0][2], "dd-MMM-yy HH:mm") + "</td>"
        } else {
            strHTML = strHTML + pmsg
        };
        if (values[0][52].indexOf("Remove-") > -1) {
            strHTML = strHTML + "<td></td>"
        } else {
            strHTML = strHTML + "<td>" + values[0][52].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";
        }
        //strHTML = strHTML + "<td>" + values[0][52].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;') + "</td>";
        strHTML = strHTML + "</tr>";

        var dd = "";
        var dis_rt = "disabled";
        if (values[0][56] != "" && values[0][52].indexOf("Remove-") < 0) {
            dd = Utilities.formatDate(values[0][56], timezone_[0][2], "dd/MM/yy");
            dis_rt = "";
            txt_rt += "<option value=\"2B\" " + dis_rt + " >[Reviewer7]  - " + values[0][53] + " - " + dd + "</option>";
        }
    }

    /******************************************************************/
    paction = values[0][2].indexOf("Rejected by Final Approver");
    if (paction > -1) {
        status_wfl = false
    }
    if (paction == -1) {
        pstatus = "Final Approved";
    } else {
        pstatus = values[0][2];
    }

    if ((j < 9) && (values[0][16] != "")) {

        nxt++;
        strHTML = strHTML + "<tr><td>" + nxt + "</td>";
        strHTML = strHTML + "<td>" + (typeof(pos[0])!="undefined"? (pos[0][9]!=""?pos[0][9]+" - ":""):"")+  values[0][16].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";

        if (status_wfl == true) {
            if (values[0][16] != "") {
                pmsg = "<td><font color=red> ((Pending Final Approve)) </font></td>";
            } else {
                pmsg = "<td> - </td>";
            }
        } else {
            pmsg = "<td></td>";
        }
        //strHTML = strHTML + "<td>Final Approve</td>";
        if (values[0][19] != "") {
            strHTML = strHTML + "<td>" + pstatus + "</td>"
        } else {
            strHTML = strHTML + pmsg
        };
        if (values[0][19] != "") {
            strHTML = strHTML + "<td>" + Utilities.formatDate(values[0][19], timezone_[0][2], "dd-MMM-yy HH:mm") + "</td>"
        } else {
            strHTML = strHTML + pmsg
        };
        strHTML = strHTML + "<td>" + values[0][15].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";
        strHTML = strHTML + "</tr>";

        var dd = "";
        var dis_rt = "disabled";
        if (values[0][19] != "") {
            dd = Utilities.formatDate(values[0][19], timezone_[0][2], "dd/MM/yy");
            dis_rt = "";
            txt_rt += "<option value=\"3\" " + dis_rt + " >[Approver]  - " + values[0][16] + " - " + dd + "</option>";
        }

    }

    if ((j < 10) && (values[0][21] != "")) {

        nxt++;
        strHTML = strHTML + "<tr><td>" + nxt + "</td>";
        strHTML = strHTML + "<td>" + values[0][21].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";
        if (status_wfl == true) {
            if (values[0][21] != "") {
                pmsg = "<td><font color=red> ((Pending Register)) </font></td>";
            } else {
                pmsg = "<td> - </td>";
            }
        } else {
            pmsg = "<td></td>";
        }

        //strHTML = strHTML + "<td>Register</td>";
        if (values[0][24] != "") {
            strHTML = strHTML + "<td>Completed</td>"
        } else {
            strHTML = strHTML + pmsg
        };
        if (values[0][24] != "") {
            strHTML = strHTML + "<td>" + Utilities.formatDate(values[0][24], timezone_[0][2], "dd-MMM-yy HH:mm") + "</td>"
        } else {
            strHTML = strHTML + pmsg
        };
        strHTML = strHTML + "<td>" + values[0][20].replace(/\"/gim, '\"').replace(/\</gim, '&lt;').replace(/\>/gim, '&gt;').replace(/\r\n/gim, '<br>').replace(/\n/gim, '<br>') + "</td>";
        strHTML = strHTML + "</tr>";

        var dd = "";
        var dis_rt = "disabled";
        if (values[0][24] != "") {
            dd = Utilities.formatDate(values[0][24], timezone_[0][2], "dd/MM/yy");
            dis_rt = "";
            txt_rt += "<option value=\"4\" " + dis_rt + " >[Register]  - " + values[0][21] + " - " + dd + "</option>";
        }
    }

    //END OF  NEXT PROCESSES (IN CASE PENDING APPROVE)

    strHTML = strHTML + "</TABLE></div>";

    //strHTML.replace("@Reroute",reroute);
    return strHTML;

}

function getNextApproveStep(pk, fr) {
    var id = pk;
    var rt = fr;
    //var user= Session.getActiveUser().getEmail().replace(/\s/g, "");
    //var comment = e.txtComment;

    var rt_pth = ["1A", "1B", "1C", "1D", "2", "2A", "2B", "3", "4"];
    var rt_pos = [27, 32, 37, 42, 11, 48, 53, 16, 21];
    var to = "";

    var lst_ind = rt_pth.length;
    var frm_ind = rt_pth.indexOf(rt);
    //var to_ind = rt_pth.indexOf(to);

    var doc_t = SpreadsheetApp.openByUrl(DATA_SPREADSHEET) // Active WFL
        //var DATA_SPREADSHEET = "https://docs.google.com/spreadsheets/d/1k0jInJm37XdD9oT2Ocn8NMpK1cvjXteESaBQpJ0eitY/edit#gid=0";
        var sht = doc_t.getSheetByName(INPUT_DATA);
    var rt_prv,
    email_re;

    for (var i = frm_ind + 1; i < lst_ind; i++) {
        // //Logger.log(rt_pth[i]+"-"+sht.getRange(id, (rt_pos[i]+1)).getValue())
        if (sht.getRange(id, (rt_pos[i] + 1)).getValue() != "" && sht.getRange(id, (rt_pos[i] + 4)).getValue() == "") {
            ////Logger.log(sht.getRange(id, (rt_pos[i] + 1)).getValue())
            rt_prv = rt_pth[i];
            email_re = sht.getRange(id, (rt_pos[i] + 1)).getValue();
            break;
        }
    }
    return rt_prv + "|" + email_re;

    
}

//*******************************************************************************
// Added for V2 : Route Change
// 2017/12/26 P.Winyou (ANTK)
// FUNCTION TO SAVE DATA OF REVIEW/APPROVE/REVIEW TO SYTEM
//*******************************************************************************

function approveForm(e) {
    "use strict";
  
    var returnVal;

    //var t1 = new Date();
    var ss1,ss2,e_sign;
    var i,rg,rows,values,rw;
    var rt,ml,msg;
    var act,mlcc,mlsubj,mlregistcc;
    var pk,next,nextProcApprove;
    var docPID,docPIDPk;

    var txtId = e.txtId;
    var txtRt = e.txtRt; //ORIGINAL
    var txtCompany = e.txtCompany;
    var txtMaster = e.txtMaster;
    var txtComment = e.txtComment
    if (typeof(e.txtComment) == "undefined") { txtComment = ""; }
    var txtJudge = e.txtJudge;
    var txtTempFileId = e.txtTempFileId;
    var txtPriority = e.priority;
    var txtUrl = "";
    var txtIssueBy;
    var urgent_reason = e.txtRef3_ ;
    if (typeof(e.txtRef3_) == "undefined") { urgent_reason = ""; }
    //return e.txtJudge;
    //-----------------------New
    //    return DATA_SPREADSHEET;
      ss1 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName(INPUT_DATA);
      var doc_id= ss1.getRange("I" + txtId).getValue();
      var doc_url = ss1.getRange("H" + txtId).getValue();
      var priority = ss1.getRange("CK" + txtId).getValue();

    //------------------------
   
//Logger.log('e.txtId : ' +  e.txtId);
//Logger.log('docPID : ' + docPID);
//Logger.log('e.txtRt : ' + e.txtRt);
//Logger.log('UniqueStr : ' + UniqueStr);
var UniqueStr = GetRow2UniID(e.txtId);
    // for check stamp sheet
    var chk_sht_his = SpreadsheetApp.openById(doc_id).getSheetByName("History");
    if (chk_sht_his == null) { return 5; }

    // for check template sheet
    var search_row,check_template_sheet,file_for_check_esig;
  //return txtTempFileId
   /*
    if (txtTempFileId!="") {
        file_for_check_esig = SpreadsheetApp.openById(txtTempFileId);
    } else {
        file_for_check_esig = SpreadsheetApp.openById(doc_id);
    }*/
    file_for_check_esig = SpreadsheetApp.openById(doc_id);
    //check_template_sheet = file_for_check_esig.getSheetByName("Template");
    // if (check_template_sheet != null) {
    var sheet =file_for_check_esig.getSheets()[0];
    var master_form = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName("Master").getDataRange().getValues().filter(function (dataRow) {return dataRow[0] === e.txtMaster && dataRow[1] ===e.txtCompany;});
    //Logger.log(master_form)
    var master_sheet_name = master_form[0][2];
    var master_sheet_id = master_form[0][8];
  
    //return master_sheet_name;
//            if (sheet.getName()=="History") {
//                sheet = file_for_check_esig.getSheets()[1];
//            }

  
  ///////////////////////////////////////////////////
  ///////////////////////MODIFIED BY WP 20210916, IN CASE SIGNATURE NOT FOUND
     sheet = file_for_check_esig.getSheetByName(master_sheet_name);
     var search_string = Session.getActiveUser().getEmail();
  
     if (search_string != "admin@ngkntk-asia.com") {
       var textFinder = sheet.createTextFinder(search_string);
       search_row = textFinder.findNext();
       
       if (search_row != null) {
         // Clear signature if reject.
         // if (txtJudge=='r'){ 
         if ((txtJudge=='r')||(txtJudge=='c')) {  //ADDED BY WP 20210916 TO REMOVE THE SIGNATURE IF CANCEL ACTION
           do {
             search_row.clearContent();
             ////Logger.log(""+search_row.getA1Notation()+":"+search_row.getValue());
             search_row = textFinder.findNext(); 
           } while (search_row!= null); 
           ////Logger.log("'" + sheet.getName() + "'!" + search_row.getA1Notation());
         }
       } else {
         if(txtJudge=='a'){ return 6;}
       }
     }
       // }
        //return doc_id
  
  /////////////////////////////////////////////////////////
  ////////////////////////////////////////////////////////
  
  
    // Start to Upload...
    var folder = DriveApp.getFolderById(e.txtFolderattId);
    for(var a in e){
      if (a.indexOf("attFile")>-1 && e[a].length !=0) {
        var zipdoc = folder.createFile(e[a]);
        
      }
    }
    var pRt = e.txtRt;
          
       /*     
        if (e.IattFile.length != 0) {
           // var t_3i_s = new Date();
            var fileBlob2 = e.IattFile;
            var fileBlobType = fileBlob2.getContentType();
            var folder = DriveApp.getFolderById(e.txtFolderattId);
            var zipdoc = folder.createFile(fileBlob2);
          
          
          /* UnzIp
            if (fileBlobType.indexOf("zip") == -1) {
                var zipdoc = folder.createFile(fileBlob2);
            } else {
                fileBlob2.setContentType("application/zip");
                var unZippedfile = Utilities.unzip(fileBlob2);
               // for each(var file_ in unZippedfile) {
                    var zipdoc = folder.createFile(file_);
                    //zipdoc.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.NONE);
               // }
            }
            
           // var t_3i_e = new Date();
            var pRt = e.txtRt;
          
            if ((pRt == "1A") || (pRt == "1B") || (pRt == "1C") || (pRt == "1D") || (pRt == "2") || (pRt == "2A") || (pRt == "2B")) {
                SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3I", t_3i_s, t_3i_e, t_3i_e - t_3i_s, version_wfl, "Upload File Attach"]);
                SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3H", t_3i_s, t_3i_s, t_3i_s - t_3i_s, version_wfl, "Upload File Form"]);
            }
            if ((pRt == "3")) {
                SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "4I", t_3i_s, t_3i_e, t_3i_e - t_3i_s, version_wfl, "Upload File Attach"]);
                SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "4H", t_3i_s, t_3i_s, t_3i_s - t_3i_s, version_wfl, "Upload File Form"]);
            }
            if ((pRt == "4")) {
                SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5I", t_3i_s, t_3i_e, t_3i_e - t_3i_s, version_wfl, "Upload File Attach"]);
                SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5H", t_3i_s, t_3i_s, t_3i_s - t_3i_s, version_wfl, "Upload File Form"]);
            }

        }
  */

        //return "OK1"
        if (typeof(e.IesigFile) != "undefined") {
            if (e.IesigFile.length != 0) {
              //  var t_3i_s = new Date();
                var fileBlob2 = e.IesigFile;
                var fileBlobType = fileBlob2.getContentType();
                var folder = DriveApp.getFolderById(ATTACH_FOLDER_UNCONVERT);
                var zipdoc = folder.createFile(fileBlob2);
                //return "OK"
                var re_name = zipdoc.getName().replace(".xlsm", "[" + txtRt + "]_.xlsm");

                zipdoc.setName(re_name);
                doc_url = zipdoc.getUrl();
              //  var t_3i_e = new Date();
                var pRt = e.txtRt;
              /*
                if ((pRt == "1A") || (pRt == "1B") || (pRt == "1C") || (pRt == "1D") || (pRt == "2") || (pRt == "2A") || (pRt == "2B")) {
                    SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3I", t_3i_s, t_3i_e, t_3i_e - t_3i_s, version_wfl, "Upload File Attach"]);
                    SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3H", t_3i_s, t_3i_s, t_3i_s - t_3i_s, version_wfl, "Upload File Form"]);
                }
                if ((pRt == "3")) {
                    SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "4I", t_3i_s, t_3i_e, t_3i_e - t_3i_s, version_wfl, "Upload File Attach"]);
                    SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "4H", t_3i_s, t_3i_s, t_3i_s - t_3i_s, version_wfl, "Upload File Form"]);
                }
                if ((pRt == "4")) {
                    SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5I", t_3i_s, t_3i_e, t_3i_e - t_3i_s, version_wfl, "Upload File Attach"]);
                    SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5H", t_3i_s, t_3i_s, t_3i_s - t_3i_s, version_wfl, "Upload File Form"]);
                }
*/
            }
        }

        //ss1.getSheetByName(INPUT_DATA).activate();

    docPID = txtId; //PID CONTROLLED BY SYSTEM
    docPIDPk = docPID + 1;
    //txtId = Number(txtId) - 1; //PIC PLUS ONE FOR HEADER FOR DATABASE ACCESSS

    ////Logger.log(txtJudge + " txtCompany:" + txtCompany + " txtMaster:" + txtMaster + " txtRt:" + txtRt);
    var text_ref1;
    pk = ss1.getRange("A" + docPID).getValue(); // get Primary Key
    text_ref1 = ss1.getRange("BF" + docPID).getValue(); // get Primary Key
    if (text_ref1.length > 120) {text_ref1=text_ref1.substr(0, 120)+"...";}
    ////Logger.log(pk)
    //return ;
    //var t6 = new Date();
    //var t_3j_s = new Date();
    //var t_3m_s = new Date();
    mlcc = getCCMail(txtId); // send id to get CC Mail
    //return mlcc
    ////Logger.log(mlcc);
    //var t_3m_e = new Date();
    //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3M", t_3m_s, t_3m_e, t_3m_e - t_3m_s, version_wfl, "getCCMail"]);
    //var t_3n_s = new Date();
    //var t5 = new Date();
    mlregistcc = getRegistCCMail(txtId); // send id to get CC Mail
    //return mlregistcc;
    ////Logger.log(mlregistcc);
    //var t_3n_e = new Date();
   // SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3N", t_3n_s, t_3n_e, t_3n_e - t_3n_s, version_wfl, "getRegistCCMail"]);

    //var t4 = new Date();
    var end_rt=0;
    // Set Urgent Flag
    ss1.getRange("CK" + txtId).setValue(txtPriority);
    urgent_mail = ""
    if  (txtPriority==0) {
      urgent_mail = "[!!!URGENT!!!]"
    } else {
      urgent_mail = ""
    }
//   if (typeof(e.txtComment) == "undefined") { txtComment = ""; }
//   if (typeof(e.txtComment) == "undefined") { txtComment = ""; }
   if (txtComment!="") {
                         txtComment = "[Comment]\r\n "+txtComment; 
   }
   if (txtPriority!=priority) { 
     if (txtPriority == 0) { 

                if (urgent_reason!="") {
                         txtComment += "\r\n[Priority Reason :: Change priority from NORMAL to URGENT ]\r\n "+urgent_reason;
                }
     }
     if (txtPriority == 1) { 
               //txtComment = "[Comment]\r\n "+txtComment+"\r\n[Priority Reason :: Change priority from URGENT to NORMAL ]\r\n " 
                if (urgent_reason!="") {
                         txtComment += "\r\n[Priority Reason :: Change priority from URGENT to NORMAL ]\r\n "+urgent_reason;
                }
     }
     } else {
     if (txtPriority == 0) { 
               //txtComment = "[Comment]\r\n "+txtComment+"\r\n[Priority Reason :: ]\r\n " +urgent_reason 
                if (urgent_reason!="") {
                         txtComment += "\r\n[Priority Reason :: ]\r\n "+urgent_reason;
                }
     }
     if (txtPriority == 1) { 
               //txtComment = "[Comment]\r\n "+txtComment+"\r\n[Priority Reason :: ]\r\n " 
                if (urgent_reason!="") {
                         txtComment += "\r\n[Priority Reason :: ]\r\n "+urgent_reason;
                }
     }
     }
    
    if (txtRt == "1A") {
        //var t_3j_s = new Date();
        // Approve Review
        ss1.getRange("AA" + txtId).setValue(txtComment);
        ss1.getRange("H" + txtId).setValue(doc_url);
        ss1.getRange("AE" + txtId).setValue(doc_url);
        ss1.getRange("AD" + txtId).setValue(doc_id);
        ss1.getRange("AE" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
        ss1.getRange("BH" + txtId).setValue(urgent_reason)
        //------------------------------------------------
        e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
        rw = e_sign.getLastRow() + 1;

        /*-------------ADDED BY P-WINYOU FOR ENTERTAIN EXPENSE AND OVERSEA BUSINESS TRIP
        var input = SpreadsheetApp.openById(doc_id).getSheetByName("Ent.pre-approval");
        //if (input != null) {

        e_sign.getRange("U31").setValue("Signature Here");
        e_sign.getRange("U32").setValue("Date/Time Here");

        //}
        //-------------END OF ADDED BY P-WINYOU 20190723*/

        e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
        e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
        e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date-Time
        e_sign.getRange("E" + rw).setValue(txtComment); //Comment
        //var t_3j_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
        //var t_3f_s = new Date();
        next = getNextApproveStep(docPID, "1A");
        //return docPID+":"+next;
        nextProcApprove = next.split("|");
      
         txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=1A&rt=" + nextProcApprove[0];
        //txtUrl = WEB_PATH + "?id=" + docPID + "&fr=1A&rt=" + nextProcApprove[0];
       
      
        //var t_3f_e = new Date();
       // SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

        if (nextProcApprove[0] == "3") {
            mlsubj = "Approver";
        } else {
            mlsubj = "Reviewer";
        }

        if (txtJudge == "a") {
            ml = nextProcApprove[1];
            msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Reviewed');
            act = "Signed By Reviewer";
            e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer1"); //Action
        } else {
            ml = ss1.getRange("F" + txtId).getValue();
            //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
            txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
            msg = "Your master registration request has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Rejected by Reviewer1');
            act = "Reviewer Rejected";
            e_sign.getRange("C" + rw).setValue("Rejected By Reviewer1"); //Action
        }

        //var t_3l_s = new Date();
        if (txtJudge == "a") {
            MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're  " + mlsubj + ">  " + text_ref1 + ":" + pk, msg);
        } else {
            MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>  " + text_ref1 + ":" + pk, msg, {
                cc: mlcc
            });
        }
        //var t_3l_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

    }

    if (txtRt == "1B") {
        //var t_3j_s = new Date();
        ss1.getRange("AF" + txtId).setValue(txtComment);
        //ss1.getRange("AL" + txtId).setValue(values[txtId][3]);
        ss1.getRange("H" + txtId).setValue(doc_url);
        ss1.getRange("AH" + txtId).setValue(doc_url);
        ss1.getRange("AI" + txtId).setValue(doc_id);
        ss1.getRange("AJ" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
        ss1.getRange("BH" + txtId).setValue(urgent_reason)
        //------------------------------------------------
        e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
        rw = e_sign.getLastRow() + 1;

        e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
        e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
        //e_sign.getRange("C"+rw).setValue("Reviewer2");//Action
        e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
        e_sign.getRange("E" + rw).setValue(txtComment); //Comment
        //var t_3j_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
        //var t_3f_s = new Date();

        next = getNextApproveStep(docPID, "1B");

        nextProcApprove = next.split("|");
        //txtUrl = WEB_PATH + "?id=" + txtId + "&fr=1B&rt=" + nextProcApprove[0];
        txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=1B&rt=" + nextProcApprove[0];
        //var t_3f_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

        if (nextProcApprove[0] == "3") {
            mlsubj = "Approver";
        } else {
            mlsubj = "Reviewer";
        }

        if (txtJudge == "a") {
            ml = nextProcApprove[1];
            msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Reviewed');
            act = "Signed By Reviewer";
            e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer2"); //Action
        } else {
            ml = ss1.getRange("F" + txtId).getValue();
              // txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
                txtUrl = WEB_PATH + "?id=" +  UniqueStr + "&rt=0&view=1";
              
            msg = "Your application has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Rejected by Reviewer2');
            act = "Reviewer Rejected";
            e_sign.getRange("C" + rw).setValue("Rejected By Reviewer2"); //Action
        }
        //var t_3l_s = new Date();
        if (txtJudge == "a") {
            MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're  " + mlsubj + ">  " + text_ref1 + ":" + pk, msg);
        } else {
            MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>  " + text_ref1 + ":" + pk, msg, {
                cc: mlcc
            });
        }
        //var t_3l_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);
    }

    if (txtRt == "1C") {
        //var t_3j_s = new Date();
        ss1.getRange("AK" + txtId).setValue(txtComment);
        //ss1.getRange("AL" + txtId).setValue(values[txtId][3]);
        ss1.getRange("H" + txtId).setValue(doc_url);
        ss1.getRange("AM" + txtId).setValue(doc_url);
        ss1.getRange("AN" + txtId).setValue(doc_id);
        ss1.getRange("AO" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
        ss1.getRange("BH" + txtId).setValue(urgent_reason)
        //------------------------------------------------
        e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
        rw = e_sign.getLastRow() + 1;

        e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
        e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
        //e_sign.getRange("C"+rw).setValue("Reviewer3");//Action
        e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
        e_sign.getRange("E" + rw).setValue(txtComment); //Comment
        //var t_3j_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
        //var t_3f_s = new Date();
        next = getNextApproveStep(docPID, "1C");
        nextProcApprove = next.split("|");
       // txtUrl = WEB_PATH + "?id=" + txtId + "&fr=1C&rt=" + nextProcApprove[0];
        txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=1C&rt=" + nextProcApprove[0];
        
        //var t_3f_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

        if (nextProcApprove[0] == "3") {
            mlsubj = "Approver";
        } else {
            mlsubj = "Reviewer";
        }

        if (txtJudge == "a") {
            ml = nextProcApprove[1];
            msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Reviewed');
            act = "Signed By Reviewer";
            e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer3"); //Action
        } else {
            ml = ss1.getRange("F" + txtId).getValue();
            //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
            txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
            
            msg = "Your application has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Rejected by Reviewer3');
            act = "Reviewer Rejected";
            e_sign.getRange("C" + rw).setValue("Rejected By Reviewer3"); //Action
        }

        //var t_3l_s = new Date();
        if (txtJudge == "a") {
            MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're  " + mlsubj + ">  " + text_ref1 + ":" + pk, msg);
        } else {
            MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>  " + text_ref1 + ":" + pk, msg, {
                cc: mlcc
            });
        }
        //var t_3l_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

    }

    if (txtRt == "1D") {
        //var t_3j_s = new Date();
        ss1.getRange("AP" + txtId).setValue(txtComment);
        //ss1.getRange("AQ" + txtId).setValue(values[txtId][3]);
        ss1.getRange("H" + txtId).setValue(doc_url);
        ss1.getRange("AR" + txtId).setValue(doc_url);
        ss1.getRange("AS" + txtId).setValue(doc_id);
        ss1.getRange("AT" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
        ss1.getRange("BH" + txtId).setValue(urgent_reason);
        //------------------------------------------------
        e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
        rw = e_sign.getLastRow() + 1;

        e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
        e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
        //e_sign.getRange("C"+rw).setValue("Reviewer4");//Action
        e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
        e_sign.getRange("E" + rw).setValue(txtComment); //Comment
        //var t_3j_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
        //var t_3f_s = new Date();
        next = getNextApproveStep(docPID, "1D");
        nextProcApprove = next.split("|");
        //txtUrl = WEB_PATH + "?id=" + txtId + "&fr=1D&rt=" + nextProcApprove[0];
        txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=1D&rt=" + nextProcApprove[0];
        
        //var t_3f_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

        if (nextProcApprove[0] == "3") {
            mlsubj = "Approver";
        } else {
            mlsubj = "Reviewer";
        }

        if (txtJudge == "a") {
            ml = nextProcApprove[1];
            msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Reviewed');
            act = "Signed By Reviewer";
            e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer4"); //Action
        } else {
            ml = ss1.getRange("F" + txtId).getValue();
          //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
          txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
          
            msg = "Your application has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Rejected by Reviewer4');
            act = "Reviewer Rejected";
            e_sign.getRange("C" + rw).setValue("Rejected By Reviewer4"); //Action
        }

        //var t_3l_s = new Date();
        if (txtJudge == "a") {
            MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're  " + mlsubj + ">  " + text_ref1 + ":" + pk, msg);
        } else {
            MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>  " + text_ref1 + ":" + pk, msg, {
                cc: mlcc
            });
        }
        //var t_3l_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

    }

    if (txtRt == "2") {
        //var t_3j_s = new Date();
        ss1.getRange("K" + txtId).setValue(txtComment);
        //ss1.getRange("L" + txtId).setValue(values[txtId][3]);
        ss1.getRange("H" + txtId).setValue(doc_url);
        ss1.getRange("M" + txtId).setValue(doc_url);
        ss1.getRange("N" + txtId).setValue(doc_id);
        ss1.getRange("O" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
        ss1.getRange("BH" + txtId).setValue(urgent_reason);
        //------------------------------------------------
        e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
        rw = e_sign.getLastRow() + 1;

        e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
        e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
        e_sign.getRange("C" + rw).setValue("Reviewer5"); //Action
        e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
        e_sign.getRange("E" + rw).setValue(txtComment); //Comment
        //var t_3j_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
        //var t_3f_s = new Date();
        next = getNextApproveStep(docPID, "2");
        nextProcApprove = next.split("|");
       // txtUrl = WEB_PATH + "?id=" + txtId + "&fr=2&rt=" + nextProcApprove[0];
        txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=2&rt=" + nextProcApprove[0];
        
        //var t_3f_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

        if (nextProcApprove[0] == "3") {
            mlsubj = "Approver";
        } else {
            mlsubj = "Reviewer";
        }

        if (txtJudge == "a") {
            ml = nextProcApprove[1];
            msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Reviewed');
            act = "Signed By Reviewer";
            e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer5"); //Action
        } else {
            ml = ss1.getRange("F" + txtId).getValue();
          //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
          txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
          
            msg = "Your application has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Rejected by Reviewer5');
            act = "Reviewer Rejected";
            e_sign.getRange("C" + rw).setValue("Rejected By Reviewer5"); //Action
        }
        //var t_3l_s = new Date();
        if (txtJudge == "a") {
            MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're Final  " + text_ref1 + ":" + mlsubj + ">  " + pk, msg);
        } else {
            MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>" + text_ref1 + ":" + pk, msg, {
                cc: mlcc
            });
        }
        //var t_3l_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

    }

    if (txtRt == "2A") {
        //var t_3j_s = new Date();
        ss1.getRange("AV" + txtId).setValue(txtComment);
        //ss1.getRange("AW" + txtId).setValue(values[txtId][3]);
        ss1.getRange("H" + txtId).setValue(doc_url);
        ss1.getRange("AX" + txtId).setValue(doc_url);
        ss1.getRange("AY" + txtId).setValue(doc_id);
        ss1.getRange("AZ" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
        ss1.getRange("BH" + txtId).setValue(urgent_reason);
        //------------------------------------------------
        e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
        rw = e_sign.getLastRow() + 1;

        e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
        e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
        //e_sign.getRange("C"+rw).setValue("Reviewer6");//Action
        e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
        e_sign.getRange("E" + rw).setValue(txtComment); //Comment
        //var t_3j_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
        //var t_3f_s = new Date();
        next = getNextApproveStep(docPID, "2A");
        nextProcApprove = next.split("|");
        //txtUrl = WEB_PATH + "?id=" + txtId + "&fr=2A&rt=" + nextProcApprove[0];
        txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=2A&rt=" + nextProcApprove[0];
        
        //var t_3f_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

        if (nextProcApprove[0] == "3") {
            mlsubj = "Approver";
        } else {
            mlsubj = "Reviewer";
        }

        if (txtJudge == "a") {
            ml = nextProcApprove[1];
            msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Reviewed');
            act = "Signed By Reviewer";
            e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer6"); //Action
        } else {
            ml = ss1.getRange("F" + txtId).getValue();
         // txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
          txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
          
            msg = "Your application has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Rejected by Reviewer6');
            act = "Reviewer Rejected";
            e_sign.getRange("C" + rw).setValue("Rejected By Reviewer6"); //Action
        }

        //var t_3l_s = new Date();
        if (txtJudge == "a") {
            MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're  " + mlsubj + ">  " + text_ref1 + ":" + pk, msg);
        } else {
            MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>  " + text_ref1 + ":" + pk, msg, {
                cc: mlcc
            });
        }
        //var t_3l_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

    }

    if (txtRt == "2B") {
        //var t_3j_s = new Date();
        ss1.getRange("BA" + txtId).setValue(txtComment);
        //ss1.getRange("BB" + txtId).setValue(values[txtId][3]);
        ss1.getRange("H" + txtId).setValue(doc_url);
        ss1.getRange("BC" + txtId).setValue(doc_url);
        ss1.getRange("BD" + txtId).setValue(doc_id);
        ss1.getRange("BE" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
        ss1.getRange("BH" + txtId).setValue(urgent_reason);
        //------------------------------------------------
        e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
        rw = e_sign.getLastRow() + 1;

        e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
        e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
        e_sign.getRange("C" + rw).setValue("Reviewer7"); //Action
        e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
        e_sign.getRange("E" + rw).setValue(txtComment); //Comment
        //var t_3j_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
        //var t_3f_s = new Date();
        next = getNextApproveStep(docPID, "2B");
        nextProcApprove = next.split("|");
        //txtUrl = WEB_PATH + "?id=" + txtId + "&fr=2B&rt=" + nextProcApprove[0];
        txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=2B&rt=" + nextProcApprove[0];
        
        //var t_3f_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

        if (nextProcApprove[0] == "3") {
            mlsubj = "Approver";
        } else {
            mlsubj = "Reviewer";
        }

        if (txtJudge == "a") {
            ml = nextProcApprove[1];
            msg = "You have received Application from WorkFlow Launcher.\r\n" + "Please confirm the following link.\r\n\r\n＜Link＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Reviewed');
            act = "Signed By Reviewer";
            e_sign.getRange("C" + rw).setValue("Reviewed By Reviewer7"); //Action
        } else {
            ml = ss1.getRange("F" + txtId).getValue();
          //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
          txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
          
            msg = "Your application has been rejected\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n＜Application Sheet＞\r\n" + txtUrl;
            ss1.getRange("C" + txtId).setValue('Rejected by Reviewer7');
            act = "Reviewer Rejected";
            e_sign.getRange("C" + rw).setValue("Rejected By Reviewer7"); //Action
        }

        //var t_3l_s = new Date();
        if (txtJudge == "a") {
            MailApp.sendEmail(ml, "" + "WFL <" + urgent_mail + "You're  " + mlsubj + ">  " + text_ref1 + ":" + pk, msg);
        } else {
            MailApp.sendEmail(ml, "" + "WFL <Rejected by Reviewer>  " + text_ref1 + ":" + pk, msg, {
                cc: mlcc
            });
        }
        //var t_3l_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "3L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

    }
  
  /////////////////////////////////////////////////////////////
  ////////////////////////START OF APPROVAL ///////////////////
  /////////////////////////////////////////////////////////////

    if (txtRt == "3") {
      
        //////////////////////////////////////////SET INITIAL VALUES
        //var t_3j_s = new Date();
        ss1.getRange("P" + txtId).setValue(txtComment);
        //ss1.getRange("Q" + txtId).setValue(values[txtId][5]);
        ss1.getRange("H" + txtId).setValue(doc_url);
        ss1.getRange("R" + txtId).setValue(doc_url);
        ss1.getRange("S" + txtId).setValue(doc_id);
        ss1.getRange("T" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
        ss1.getRange("BH" + txtId).setValue(urgent_reason);
        //------------------------------------------------
      
        ///////////////////////////////////UPDATE DATA TO THE HISTORY SHEET IN DATASHEET ITSELF
        e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
        rw = e_sign.getLastRow() + 1;

        e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
        e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
        //e_sign.getRange("C"+rw).setValue("Approver");//Action
        e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
        e_sign.getRange("E" + rw).setValue(txtComment); //Comment
        //var t_3j_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "4J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
        //var t_3f_s = new Date();
        next = getNextApproveStep(docPID, "3");
      
        ///////////////////////////////////////GENTERATE URL FOR EMAIL SENDING      
        // return next.toString();
        if (next.split("|")[0]!="undefined") {
          //has register
          nextProcApprove = next.split("|");
          //txtUrl = WEB_PATH + "?id=" + txtId + "&fr=3&rt=" + nextProcApprove[0];
          txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=3&rt=" + nextProcApprove[0];
          
        } else {
          nextProcApprove=["",""];
          //txtUrl = WEB_PATH + "?id=" + txtId + "&fr=3&rt=9";
          txtUrl = WEB_PATH + "?id=" + UniqueStr + "&fr=3&rt=9";
          
          end_rt = 1;
        }
        //var t_3f_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "4F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);

      
        ///////////////////////////////////////DETERMINE EMAIL SUBJECT      
        if (nextProcApprove[0] == "3") {
            mlsubj = "Approver";
        } else {
            mlsubj = "Reviewer";
        }
      
      
        //////////////////////////////////////CHECK IF ACTION IS APPROVE OR REJECT OR CANCEL APPROVE
        /////////////////////////////////////////////////////////////////////////////////////////////
      
      
        if (txtJudge == "a") { /////////////APPROVE

            ml   = nextProcApprove[1];

            mlcc = "admin@ngkntk-asia.com,"+ss1.getRange("F" + txtId).getValue()+","+ss1.getRange("AU" + txtId).getValue(); //ORIGINAL
            mlcc = mlcc.toString().replace(" ",""); //ADDED BY WNP 20210707 TO FIX CASE SPACE WITH EMAIL ADDRESS - ERROR INVALID ADDRESS
          
          
            //return ml;
            msg = "WorkFlow Launcher have been approved. Please perform next process.\r\n"
                //+ "This Application is not completed until Application Admin processed the signed document.\r\n\r\n＜Link＞\r\n" + txtUrl
                 + "This Application is not completed until Application Admin processed the signed document.\r\n"
                 + "\r\n＜Application sheet＞\r\n" + txtUrl

                if (nextProcApprove[1] != "") {

                    msg = "WorkFlow Launcher have been approved. Please perform next process.\r\n"
                         + "This Application is not completed until Application Admin processed the signed document.\r\n\r\n＜Link＞\r\n" + txtUrl
                         + "\r\n";//＜Application sheet＞\r\n" + doc_url

                } else {

                    msg = "WorkFlow Launcher have been approved. Please perform next process.\r\n"
                         + "\r\n\r\n＜Application sheet＞\r\n" + txtUrl;

                }
          
            //////////////UPDATE STATUS TO DATABASE SHEET AND DATA SHEET///////////////////

            ss1.getRange("C" + txtId).setValue('Final Approved');
            act = "Final Approve - Signed";
            e_sign.getRange("C" + rw).setValue("Final Approved"); //ActionS
          
          
                ///////////////////////////////////SEND EMAIL

                //var t_3l_s = new Date();
                  m_s = ("|,"+ml+","+mlcc+",|").replace(",,,",",").replace(",,",",").replace("|,","").replace(",|","")
                  //return "WFL : <" + urgent_mail + "" + text_ref1 + ":" + pk + "> has been Approved.";
                  MailApp.sendEmail(m_s, "" + "WFL : <" + urgent_mail + "" + text_ref1 + ":" + pk + "> has been Approved.", msg);

          
      
          
        } else if (txtJudge == "r") { ////////////////////////////REJECT
          
            ml = ss1.getRange("F" + txtId).getValue();
          
            mlcc = "admin@ngkntk-asia.com,"+ss1.getRange("F" + txtId).getValue()+","+ss1.getRange("AU" + txtId).getValue(); //ORIGINAL
            mlcc = mlcc.toString().replace(" ",""); //ADDED BY WNP 20210707 TO FIX CASE SPACE WITH EMAIL ADDRESS - ERROR INVALID ADDRESS
          
            //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
            txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
            
            msg = "Your application has been rejected by approver.\r\n\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n\r\n"+ txtUrl;//＜Application Sheet＞\r\n" + doc_url;
            ss1.getRange("C" + txtId).setValue('Rejected by Final Approver');
            act = "Final Approve - Rejceted";
            e_sign.getRange("C" + rw).setValue("Final Approve Rejceted"); //Action
          
                ///////////////////////////////////SEND EMAIL

                //var t_3l_s = new Date();

                  //MailApp.sendEmail(mlcc + mlregistcc,"" + ss1.getSheetByName(INPUT_DATA).getRange("D" + txtId).getValue() + " Application <Rejected by Final Approver>", msg);
                  m_s = ("|,"+ml+","+mlcc+",|").replace(",,,",",").replace(",,",",").replace("|,","").replace(",|","") 
                  MailApp.sendEmail(m_s, "" + "WFL <Rejected by Final Approver>" + text_ref1 + ":" + pk, msg, { cc: mlcc});
          
          
            
            /////////////////////////////////ADDED BY WP 20210916
            /////////////////////////////////REMOVE SIGNATURE FROM THE SPREADSHEET
            /*
            /////////////////////////////////////////REMOVE SIGNATURE FROM TEMPLATE FILE
            //sheet = file_for_check_esig.getSheetByName(master_sheet_name);
            var search_string = Session.getActiveUser().getEmail();
  
            if (search_string != "admin@ngkntk-asia.com") {
              var textFinder = sheet.createTextFinder(search_string);
              search_row = textFinder.findNext();
              
              if (search_row != null) {

                  do {
                    search_row.clearContent();
                    search_row = textFinder.findNext(); 
                    SpreadsheetApp.flush(); //ADDED BY WP 20210917                    
                  } while (search_row!= null); 
                
              }
            }
          
          
            SpreadsheetApp.flush(); //ADDED BY WP 20210917  
            /////////////////////////////////END OF ADDED BY WP 20210916
            /////////////////////////////////REMOVE SIGNATURE FROM THE SPREADSHEET
            */
        }
      

          //////////////////////////////////////////////////////////////////////////
          //////////////////////////////////////////////////////////////////////////
          //////ADDED BY WP 20210914 - 'CANCEL APPROVE' BUTTON
          //////TO RESTORE STAGE BACK BEFORE FINAL APPROVE
          //////NO NEED TO SEND EMAIL
                   
          
        else if (txtJudge == "c") {   //ACTION CANCELL
          
          
            /////////////////////////////////ADDED BY WP 20210916
            /////////////////////////////////REMOVE SIGNATURE FROM THE SPREADSHEETS
          

          
 
            ////////////////////////////////////////REMOVE SIGNATRUE FROM TEMPLATE FILE
            //////////////////////////////////////////FIND TEMPLATE FILE
            /*
 
            var FolderTempUserEmail;
            var FolderTempID = "1T0JMNMr4Bw94RMcu8TFkZl6K_B4HH37S";
            var FolderTempUser = DriveApp.getFolderById(FolderTempID); //ACCESS TO FOLDER 'TEMPLATE_USERS' 
            var user_email = Session.getActiveUser().getEmail();
            FolderTempUserEmail =  FolderTempUser.getFoldersByName(user_email); //ACCESS TO FOLDER OF THE CURRENT USER
          
            var file_temp,sht_temp;
            var active_record = ss1.getRange("A" + txtId+":CK"+txtId).getValues();
            var doc_type = active_record[0][3]; //GET TEMPLATE NAME OF THIS DOCUMENT
          
            file_temp = FolderTempUserEmail.getFilesByName(doc_type); //GET TEMPLATE FILE FROM FOLDER OF THE CURRENT USER
            sht_temp = file_temp.getSheetByName(master_sheet_name);

            var search_string = user_email;
  
            if (search_string != "admin@ngkntk-asia.com") {
              var textFinder = sht_temp.createTextFinder(search_string);
              search_row = textFinder.findNext();
              
              if (search_row != null) {

                  do {
                    search_row.clearContent();
                    search_row = textFinder.findNext(); 
                    SpreadsheetApp.flush(); //ADDED BY WP 20210917                    
                  } while (search_row!= null); 
                
              }
            }
            */
            ////////////////////////////////////////END OF REMOVE SIGNATRUE FROM TEMPLATE FILE

            ////////////////////////////////////////REMOVE SIGNATRUE FROM DATA FILE          
            //sheet = file_for_check_esig.getSheetByName(master_sheet_name);
            var search_string = Session.getActiveUser().getEmail();
  
            if (search_string != "admin@ngkntk-asia.com") {
              var textFinder = sheet.createTextFinder(search_string);
              search_row = textFinder.findNext();
              
              if (search_row != null) {

                  do {
                    search_row.clearContent();  //MODIFIED BY WP 20210917
                    SpreadsheetApp.flush(); //ADDED BY WP 20210917
                    search_row = textFinder.findNext(); 
                  } while (search_row!= null);                
              }
            }
            /////////////////////////////////END OF ADDED BY WP 20210916
            /////////////////////////////////REMOVE SIGNATURE FROM THE SPREADSHEET

          ss1.getRange("C" + txtId).setValue('Reviewed'); //BACK STAGE
          ss1.getRange("H" + txtId).setValue(doc_url);
          act = "Cancel Action";
          e_sign.getRange("C" + rw).setValue("Cancel Action"); //KEEP ACTION LOG    
          
          ///////////////////CLEAR DATA - RESTORE TO BLANK (Approve Cells)
          ss1.getRange("P" + txtId).setValue("");
          ss1.getRange("R" + txtId).setValue("");
          ss1.getRange("S" + txtId).setValue("");
          ss1.getRange("T" + txtId).setValue("");
          ss1.getRange("BH" + txtId).setValue("");
          //returnVal = 99;
          
          ///////////////////CLEAR DATA - RESTORE TO BLANK (Register Cells)

         // if ( end_rt==1) { //20211001
            ss1.getRange("V" + txtId).setValue("");  
            ss1.getRange("U" + txtId).setValue("");
            ss1.getRange("W" + txtId).setValue("");
            ss1.getRange("X" + txtId).setValue("");
            ss1.getRange("Y" + txtId).setValue("");
          //------------------------------------------------
          
            ml = ss1.getRange("F" + txtId).getValue();
          
            mlcc = "admin@ngkntk-asia.com,"+ss1.getRange("F" + txtId).getValue()+","+ss1.getRange("AU" + txtId).getValue(); //ORIGINAL
            mlcc = mlcc.toString().replace(" ",""); //ADDED BY WNP 20210707 TO FIX CASE SPACE WITH EMAIL ADDRESS - ERROR INVALID ADDRESS
          
            //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=0&view=1";
            txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=0&view=1";
            
            msg = "Approver has cancelled the previous action, the document is now pending for Approver.\r\n\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n\r\n"+ txtUrl;//＜Application Sheet＞\r\n" + doc_url;
            ss1.getRange("C" + txtId).setValue('Reviewed');
            act = "Final Approve - Rejceted";
            e_sign.getRange("C" + rw).setValue("Action Cancelled"); //Action
          
                ///////////////////////////////////SEND EMAIL
                //MailApp.sendEmail(mlcc + mlregistcc,"" + ss1.getSheetByName(INPUT_DATA).getRange("D" + txtId).getValue() + " Application <Action Cancelled by Final Approver>", msg);
                m_s = ("|,"+ml+","+mlcc+",|").replace(",,,",",").replace(",,",",").replace("|,","").replace(",|","") 
                MailApp.sendEmail(m_s, "" + "WFL <Action Cancelled by Final Approver>" + text_ref1 + ":" + pk, msg, { cc: mlcc});
    
        }
      
     
          //////END OF ADDED BY WP 20210914 - 'CANCEL APPROVE' BUTTON
          ////////////////////////////////////////////////////////////////////////////          
          ////////////////////////////////////////////////////////////////////////////  

    }
  
  /////////////////////////////////////////////////////////////
  ////////////////////////END OF APPROVAL /////////////////////
  /////////////////////////////////////////////////////////////

    //if (txtRt == "4" || end_rt == 1) {
    if ( (txtRt == "4" || end_rt == 1) && (txtJudge != "c")) { //20211001
      
        //var t_3j_s = new Date();
        act = "Register Signed";

        //ss2 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET).getSheetByName(name);
        //ss2.getSheetByName(INPUT_DATA).activate();
        //txtUrl = WEB_PATH + "?id=" + txtId + "&rt=9&view=1";
        txtUrl = WEB_PATH + "?id=" + UniqueStr + "&rt=9&view=1";
        
        msg = "Process completed.\r\n\r\n" + "＜Comment＞\r\n" + txtComment + "\r\n\r\n"+ txtUrl;//＜Application Sheet＞\r\n" + doc_url;
        ss1.getRange("C" + txtId).setValue('Registerd');
        ss1.getRange("U" + txtId).setValue(txtComment);
        if ( end_rt==1) {
        ss1.getRange("V" + txtId).setValue("<Auto Register After Approved>");  
        } else {
        ss1.getRange("V" + txtId).setValue(Session.getActiveUser().getEmail());
        }
        ss1.getRange("H" + txtId).setValue(doc_url);
        ss1.getRange("W" + txtId).setValue(doc_url);
        ss1.getRange("X" + txtId).setValue(doc_id);
        ss1.getRange("Y" + txtId).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"));
      //return "track : " + ss1.getName()
        //------------------------------------------------
        e_sign = SpreadsheetApp.openById(doc_id).getSheetByName("History");
        rw = e_sign.getLastRow() + 1;

        e_sign.getRange("A" + rw).setValue((rw - 1)); //Step
        e_sign.getRange("B" + rw).setValue(Session.getActiveUser().getEmail()); //E-Mail
        e_sign.getRange("C" + rw).setValue("Registerd"); //Action
        e_sign.getRange("D" + rw).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")); //Date/Time
        e_sign.getRange("E" + rw).setValue(txtComment); //Comment
      
        //var t_3j_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5J", t_3j_s, t_3j_e, t_3j_e - t_3j_s, version_wfl, "Write to Master Sheet "]);
        //var t_3f_s = new Date();
        //next = getNextApproveStep(docPID,"2A");
        //var t_3f_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5F", t_3f_s, t_3f_e, t_3f_e - t_3f_s, version_wfl, "get Next action sign e-mail"]);
        //var t_3l_s = new Date();
        mlcc ="admin@ngkntk-asia.com,"+ss1.getRange("F" + txtId).getValue()+","+ss1.getRange("AU" + txtId).getValue();
        //MailApp.sendEmail(mlcc , "" + "WFL : <"+ urgent_mail  + text_ref1 + ":" + pk + "> has been Registered.", msg);
        
        if (pk.indexOf('ANGK') > -1) {
          if (pk.indexOf('Items') > -1) {
            //Logger.log("Found DBSheet Items of ANGK");
          } else {
            MailApp.sendEmail(mlcc , "" + "WFL : <"+ urgent_mail  + text_ref1 + ":" + pk + "> has been Registered.", msg);
          }
        } else {
          MailApp.sendEmail(mlcc , "" + "WFL : <"+ urgent_mail  + text_ref1 + ":" + pk + "> has been Registered.", msg);
        }
        
      //var t_3l_e = new Date();
        //SpreadsheetApp.openById(Sheet_Test).getSheetByName(Sheet_Test_Name).appendRow([Session.getActiveUser().getEmail(), "5L", t_3l_s, t_3l_e, t_3l_e - t_3l_s, version_wfl, "Send Mail to Next Process"]);

    }
/*
     if (txtTempFileId!="") { 
          // Copy Data To WorkSheet  
                 //.openByUrl(ss1.getSheetByName(INPUT_DATA).getRange("H" + docPID).getValue())
                 var sheets_c = SpreadsheetApp.openById(txtTempFileId).getSheets() ; //SpreadsheetApp.open(doc_id).getSheets();
                 var file_des = SpreadsheetApp.openByUrl(ss1.getRange("H" + docPID).getValue());
       //return "OK"          
       //return sheets_c
                for (var i = 0; i < sheets_c.length ; i++ ) {
                   var sheet_c = sheets_c[i];
                   if (sheet_c.getName()!= "Template") {
                     if (file_des.getSheetByName(sheet_c.getName())!=null) {
                       var rng_all = sheet_c.getDataRange().getValues();
                       //var doc_i = SpreadsheetApp.open(file_temp);
                       file_des.getSheetByName(sheet_c.getName()).getRange(1, 1, rng_all.length, rng_all[0].length).setValues(rng_all);
                     }  
                   }    
                 }
    }
    */
    //START WRITE LOG
    //var t3 = new Date();
    ss3 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
    //ss3.getSheetByName("Logs").activate();
    ss3.getSheetByName("Logs").appendRow([
            　　　 pk,
            txtCompany,
            act,
            txtMaster,
            txtComment, // + " - by - " + Session.getActiveUser().getEmail(),
            Session.getActiveUser().getEmail(),
            Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"),
            doc_url,
            doc_id,
            Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"), txtUrl
        ]);

    //END FOR WRITE LOG
    //var t2 = new Date();
  /* 
  var diff = Math.floor((t2 - t1) / (1000.0)); //seconds
    var diff_log = Math.floor((t2 - t3) / (1000.0)); //seconds
    var diff_wrt = Math.floor((t3 - t4) / (1000.0)); //seconds
    var diff_CCregmail = Math.floor((t4 - t5) / (1000.0)); //seconds
    var diff_regmail = Math.floor((t5 - t6) / (1000.0)); //seconds
*/
    ////Logger.log("Over All :"+diff+" Second");
    ////Logger.log("+ WRT LOG :"+diff_log+" Second");
    ////Logger.log("+ WRT SHEET :"+diff_wrt+" Second");
    ////Logger.log("+ GET CC REGIST MAIL :"+diff_CCregmail+" Second");
    ////Logger.log("+ GET  REGIST MAIL :"+diff_regmail+" Second");
  
    if (txtJudge == "c") {
    
      return 99; //IN CASE CANCEL ACTION
    
    } else if (txtJudge == "r") {
    
      return 98; //IN CASE REJECT
    
    } else if (txtJudge == "a") {
    
      return 100; //IN CASE APPROVE
    
    } else {
      
      return 10; //IN CASE REROUTE
    }
  
    ss1 = null;
    ss3 = null;
    sheet = null;
    e_sign = null;

}

function getCCMail(pk) {

    var mail9,
    ccMail,
    mail1,
    mail2,
    mail3,
    mail4,
    mail5,
    mail6,
    mail7,
    mail8;
    var signDt9,
    signDt1,
    signDt2,
    signDt3,
    signDt4,
    signDt5,
    signDt6,
    signDt7,
    signDt8;

    //ss = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
    //rg = ss.getSheetByName(INPUT_DATA).getDataRange();
    //rows = rg.getLastRow();
    //values = rg.getValues();

    ss1 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
    //ss1.getSheetByName(INPUT_DATA).activate();

    /****************************************** UAT

    ccMail = "";
    mail1 = ss1.getRange("F" + pk).getValue();

    signDt2 = ss1.getRange("AE" + pk).getValue();
    mail2 = ss1.getRange("L" + pk).getValue();

    signDt3 = ss1.getRange("AJ" + pk).getValue();
    mail3 = ss1.getRange("AG" + pk).getValue();

    signDt4 = ss1.getRange("AO" + pk).getValue();
    mail4 = ss1.getRange("AQ" + pk).getValue();

    signDt5 = ss1.getRange("AT" + pk).getValue();
    mail5 = ss1.getRange("AB" + pk).getValue();

    signDt6 = ss1.getRange("O" + pk).getValue();
    mail6 = ss1.getRange("AL" + pk).getValue();

    signDt7 = ss1.getRange("O" + pk).getValue();
    mail7 = ss1.getRange("AL" + pk).getValue();

    signDt8 = ss1.getRange("O" + pk).getValue();
    mail8 = ss1.getRange("AL" + pk).getValue();

    /********************************************/

    ccMail = "";
    mail1 = ss1.getRange("F" + pk).getValue();

    signDt9 = ss1.getRange("AE" + pk).getValue(); //REVIEWER #1
    mail9 = ss1.getRange("AB" + pk).getValue();

    signDt2 = ss1.getRange("AT" + pk).getValue();
    mail2 = ss1.getRange("AQ" + pk).getValue();

    signDt3 = ss1.getRange("AJ" + pk).getValue();
    mail3 = ss1.getRange("AG" + pk).getValue();

    signDt4 = ss1.getRange("AO" + pk).getValue();
    mail4 = ss1.getRange("AL" + pk).getValue();

    signDt5 = ss1.getRange("AT" + pk).getValue();
    mail5 = ss1.getRange("AQ" + pk).getValue();

    signDt6 = ss1.getRange("O" + pk).getValue();
    mail6 = ss1.getRange("L" + pk).getValue();

    /********************************************/

    signDt7 = ss1.getRange("AZ" + pk).getValue();
    mail7 = ss1.getRange("AW" + pk).getValue();

    signDt8 = ss1.getRange("BE" + pk).getValue();
    mail8 = ss1.getRange("BB" + pk).getValue();

    /********************************************/

    if (mail1 != "") {
        ccMail = ccMail + mail1 + ',';
    }
    if ((mail2 != "") && (signDt2 != "")) {
        ccMail = ccMail + mail2 + ',';
    }
    if ((mail3 != "") && (signDt3 != "")) {
        ccMail = ccMail + mail3 + ',';
    }
    if ((mail4 != "") && (signDt4 != "")) {
        ccMail = ccMail + mail4 + ',';
    }
    if ((mail5 != "") && (signDt5 != "")) {
        ccMail = ccMail + mail5 + ',';
    }
    if ((mail6 != "") && (signDt6 != "")) {
        ccMail = ccMail + mail6 + ',';
    }

    /********************************************/

    if ((mail7 != "") && (signDt7 != "")) {
        ccMail = ccMail + mail7 + ',';
    }
    if ((mail8 != "") && (signDt8 != "")) {
        ccMail = ccMail + mail8 + ',';
    }
    if ((mail9 != "") && (signDt9 != "")) {
        ccMail = ccMail + mail9 + ',';
    } //REVIEWER #1

    /********************************************/

    /*
    //Logger.log(mail1);
    //Logger.log(mail2);
    //Logger.log(mail3);
    //Logger.log(mail4);
    //Logger.log(mail5);
    //Logger.log(mail6);
    //Logger.log(ccMail);
     */

    return ccMail;

}

function getRegistCCMail(pk) {

    var ccMail,
    mail1;

    //ss = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
    //rg = ss.getSheetByName(INPUT_DATA).getDataRange();
    //rows = rg.getLastRow();
    //values = rg.getValues();

    ss1 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
    //ss1.getSheetByName(INPUT_DATA).activate();

    ccMail = "";
    mail1 = ss1.getRange("AU" + pk).getValue();

    if (mail1 != "") {
        ccMail = ccMail + "," + mail1;
    }

    return ccMail;

}



function removefile(id) {
    var file = DriveApp.getFileById(id);
    if (!file) {
        return "File not exist. Can't Remove.  "
    } else {
      //Drive.Files.remove(file.id);
      //DriveApp.removeFile(file)
      file.setName("Delete_"+file.getName())
        //file.setTrashed(true);
        return 1;
    }
}

function copyDataDbSheet(id_){
    //PJA-- var sl = SpreadsheetApp.openById("1RuGoQQiY-pHzyn01Iz3Rk3kMQU79qvZApyDRZO5hEPI");
    var ss1 = SpreadsheetApp.openByUrl(DATA_SPREADSHEET);
  
    var sht = SpreadsheetApp.openById(id_);
//Logger.log("id_ : " + id_);    
    var fr = sht.getSheetByName("Parameter").getRange("B3").getValue();
  
    if (fr =="") {return "WFL Copy Data Not Complete"}
  
    var comp = sht.getSheetByName("Parameter").getRange("B5").getValue().toString();
  
    comp = comp.split("_")[1];

    var timezone_ = ss1.getSheetByName("Timezone").getRange("A:C").getValues().filter(function (dataRow) {return dataRow[0] ==comp;});
    timezone_ = timezone_ ? timezone_:ss1.getSheetByName("Timezone").getRange("A:C").getValues().filter(function (dataRow) {return dataRow[0] =="DEFAULT";});
    sht.setSpreadsheetTimeZone(timezone_[0][1]);

    var to = sht.getId();
  
      //PJA-- var    sl = SpreadsheetApp.openById("1RuGoQQiY-pHzyn01Iz3Rk3kMQU79qvZApyDRZO5hEPI");
      //PJA-- var  sl_s = sl.getSheetByName("Tracking")
      user_email = Session.getActiveUser().getEmail();
  
    var sht_c = SpreadsheetApp.openById(fr);
      //PJA-- sl_s.appendRow([user_email,"Start Copy",new Date(),sht_c.getName()])    
                  var sheets_c = sht_c.getSheets();
                  var sheet_c,sheet_d;
                  for (var i = 0; i < sheets_c.length ; i++ ) {
                        sheet_c = sheets_c[i];                     
                        sheet_d = sht.getSheetByName(sheet_c.getName());
                       if ((sheet_c.getName()!="Parameter")&&(sheet_d!=null)) {                   
                           var rng_all = sheet_c.getDataRange();
                           var val_all = rng_all.getDisplayValues(); //.getValues();getDisplayValues
                           
//PJA++ 20220426 Change to API : For fix problem Service error: Spreadsheet (ERROR : サービス エラー: スプレッドシート)

                           //sheet_d.getRange(1,1,val_all.length, val_all[0].length).setValues(val_all);     //Old Coding
                           
                           var rng2 = sheet_d.getRange(1,1,val_all.length, val_all[0].length);
                           var request = {
                            'valueInputOption': 'USER_ENTERED',
                            'data': [
                              {
                                'range': sheet_d.getName() + '!' + rng2.getA1Notation(),
                                'majorDimension': 'ROWS',
                                'values': val_all
                              }
                            ]
                          };
                          Sheets.Spreadsheets.Values.batchUpdate(request, id_);
//PJA-- 20220426 Change to API : For fix problem Service error: Spreadsheet (ERROR : サービス エラー: スプレッドシート)
                     }    
                   }
      var backToDate = new Date()
  
      //PJA-- sl_s.appendRow([user_email,"Copy Finish",backToDate,sht_c.getName()])
       return 7;

}

function startlog(dd){
    var backToDate = new Date(dd)
   //return backToDate.toUTCString() + "-" + dd
    var    sl = SpreadsheetApp.openById("1RuGoQQiY-pHzyn01Iz3Rk3kMQU79qvZApyDRZO5hEPI");
    var  sl_s = sl.getSheetByName("Tracking")
    user_email = Session.getActiveUser().getEmail();
    sl_s.appendRow([user_email,"Click",backToDate])
    return 11
}

function getScriptURL(urlText) {
  urlText = "http://script.google.com/a/ngkntk-asia.com/macros/s/AKfycbz8Wg13Z88G9Y7eYd6o84ity3bCLGX-7x7jBrtnxg/dev?id=14678&fr=0&rt=3";
  return urlText; //ScriptApp.getService().getUrl();
}

//------------------ Preview PDF SOM20240318
function getScriptURL(urlText) {
  urlText = "http://script.google.com/a/ngkntk-asia.com/macros/s/AKfycbz8Wg13Z88G9Y7eYd6o84ity3bCLGX-7x7jBrtnxg/dev?id=14678&fr=0&rt=3";
  return urlText; //ScriptApp.getService().getUrl();
}

function getFileAsBlob(exportUrl) {
  var response = UrlFetchApp.fetch(exportUrl, {
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
  });
  return response.getBlob();
}
function convertPDF(doc_id, doc_name, doc_url, docSetting) {
  try{
  doc_url = doc_url.replace("edit", docSetting);
  copyDataDbSheet(doc_id);
  var blobPDF = getFileAsBlob(doc_url);
  blobPDF.setName(doc_name);
  var filePDF = DriveApp.createFile(blobPDF);
  var folderPDF = DriveApp.getFolderById("1fjNwzeYp9vXZX7RltyQBWDmv3sqeL6hT");
  filePDF.moveTo(folderPDF);

  filePDF.setOwner("admin@ngkntk-asia.com");
  filePDF.setShareableByEditors(true);

  return filePDF.getUrl();
  } catch (ex) {
    //LogTit
    console.error("Error At")
    console.error(catchToString(ex))
  }
}

function catchToString(err) {
  var errInfo = "Catched something:\n";
  for (var prop in err) {
    errInfo += "  property: " + prop + "\n    value: [" + err[prop] + "]\n";
  }
  errInfo += "  toString(): " + " value: [" + err.toString() + "]";
  return errInfo;
}
//------------------ END Preview PDF SOM20240318

function abc(){
 var min = 1;
 var max = 100;
 const now =  Utilities.formatDate(new Date(), "GMT+7:00", "yyyyMMddHHmmss") + Math.random() * (max - min) + min;
 var result = parseInt(now).toString();
 
;
}
