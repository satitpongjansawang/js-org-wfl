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