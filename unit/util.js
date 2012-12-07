function getColumnByDate(date) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRange = sheet.getRange(2, sheet.getLastColumn());
  var chartDates = sheet.getRange("E2:" + lastRange.getA1Notation()).getValues();
  if( date < chartDates[0][0] )
    return 1;
  for(var i=0; i<chartDates[0].length; i++){
    var tmp = chartDates[0][i];
    if (typeof tmp == 'object'){
      if(judgeDate(tmp, date) == 1)
        return (i+1);
    }
  }
  return (i);
}

function addProperty(key,obj){
  var json = ScriptProperties.getProperty(key);
  if(json == null) var newObj = obj;
  else{
    var spObj = Utilities.jsonParse(json);
    var newObj = merge(spObj,obj);
  }
  var newJson = Utilities.jsonStringify(newObj);
  ScriptProperties.setProperty(key, newJson);
}

