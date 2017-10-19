function monthLZero(num) {
  Logger.log(num);
  if ( num < 10 ) {
    var str_month = "0"+num.toString();
  } else if ( num >= 10 && num < 100 ) {
    var str_month = num.toString();
  }
  return(str_month);
}

function Rename() {
  var mySS = SpreadsheetApp.getActive();
  Logger.log(mySS.getName());
  
  var asheet = mySS.getActiveSheet();
  Logger.log(asheet.getName());
  
  var rpdate_raw = asheet.getRange("A1").getValue();
  var rpdate = rpdate_raw.getFullYear().toString() + monthLZero(rpdate_raw.getMonth()+1) + rpdate_raw.getDate().toString();
  Logger.log(rpdate);
  
  mySS.getActiveSheet().setName("Query_"+rpdate);
}
