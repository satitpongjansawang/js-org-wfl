function findRow(pid){
    var source_of_row = sheet.getRange('A:A').getValues();
    var rowIndex = source_of_row.findIndex(row => row[0] === pid);
    if (rowIndex !== -1) {
        return rowIndex+1;
    }else{
        return "";
    }
}