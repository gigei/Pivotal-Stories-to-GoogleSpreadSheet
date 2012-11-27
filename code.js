// Copyright (C) 2012 Nihon Gigei, Inc. <http://rakumo.gigei.jp/>
// Licensed under the terms of the MIT License

/////////////////////////////
// 初期設定
/////////////////////////////

// Global Params
var HEAD = 3;
var LEFT = 4;
var COLOR = {
  done    : 'dimgray',
  current : 'royalblue',
  backlog : 'mediumaquamarine',
  titlebg : 'steelblue',
  titlefg : 'white',
  create  : 'red',
  update  : 'orange',
  accept  : 'yellow',
  deadline: 'pink',
};
var WIDTH = [18,60,200,70,140];
var MSHEET = 'プロジェクト管理';
var PARANUM = 7;
var LIMIT  = 50;

// 各セルの色
var COLOR_HOLIDAY = "#DDD";
var COLOR_WEEKDAY = "#FFD";
var COLOR_TODAY_BACKGROUND = "#99F";
var COLOR_TODAY_TEXT = "#000";
var COLOR_DATE_BACKGROUND = "lightgray";
var COLOR_DATE_TEXT = "#000";

// シートオープン時にメニュー追加
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var menuEntries = [
  {name: "プロジェクト選択",    functionName : "selectProject"},
  {name: "データベース一括更新",    functionName : "reloadStories"},
  {name: "新規シート作成",    functionName : "newSheets"},
  {name: "シートの更新",    functionName : "rewrite"}
  ];
  ss.addMenu("PivotalTracker", menuEntries);
  var apiUrl = 'https://www.pivotaltracker.com/services/v3/';
  var prop = ScriptProperties.getProperties();
  if(!( 'API' in prop )) ScriptProperties.setProperty('API',apiUrl);
}


// 未作成の選択状態プロジェクトからシート作成
function newSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var db = ScriptDb.getMyDb();
  var mData = getManageData('select');
  
  // 新しいシートを作成
  for(var i in mData){
    var flag = false;
    if(mData[i].flag == true){
      for(var j=0; j<sheets.length; j++){
        if(sheets[j].getName() == mData[i].projectName){
          flag = true;
          ss.toast("["+sheets[j].getName()+"] already exists");
        }
      }
      if(!flag){
        var newSheet = createNewSheet(mData[i].projectName, i);
        ScriptProperties.setProperty(mData[i].projectName, i);
        newSheet.activate();
      }
    }
  }
}


// ユーザーデータの取得
function setUserData() {
  var apiUrl = 'https://www.pivotaltracker.com/services/v3/';
  var prop = ScriptProperties.getProperties();
  if(!( 'API' in prop )) ScriptProperties.setProperty('API',apiUrl);

  var user = Browser.inputBox('User Name');
  var pass = Browser.inputBox('Password');
  var token = getToken(user,pass);
  Browser.msgBox('token : ' + token);
  
  var userProp = {
    'pivotal_user'  : user,
    'pivotal_token' : token
  };
  UserProperties.setProperties(userProp);
  return userProp;
}



///////////////////////////////
// APIまわり
///////////////////////////////

// APIトークンの取得
function getToken(user,pass){
  var apiUrl = ScriptProperties.getProperty('API');

  // リクエストの中身
  var url = 'tokens/active';
  var auth_data = Utilities.base64Encode(user + ":" + pass);
  var headers = {"Authorization" : "Basic " + auth_data};
  var params = {"headers" : headers};
  
  // fetch
  var res = UrlFetchApp.fetch(apiUrl+url , params).getContentText();
  
  // XML parse
  var els = Xml.parse(res, true).getElement();
  var token = els.getElement('guid').getText();  //tokenを取得
  return token;
}

// API fetch & XML parse
function getElementFromApi(extUrl) {
  var apiUrl = ScriptProperties.getProperty('API');
  var token = UserProperties.getProperty('pivotal_token');
  var headers = {"X-TrackerToken"  : token,
                "X-Tracker-Use-UTC": true};
  var params ={method : 'get',
               headers: headers};
  // fetch
  var res = UrlFetchApp.fetch(apiUrl+extUrl , params);
  var txt = res.getContentText();
  // XML parse
  var xml = Xml.parse(txt , true);
  var els = xml.getElement();
  return els;
}


//////////////////////////////
// シートの整形
//////////////////////////////

//ガントチャートのシートを整形
function formGantt(pData,row_num){
  var sheet = SpreadsheetApp.getActiveSheet();
  var itNum = pData.end - pData.start + 1;
  var pid   = pData.ID;
  
  sheet.getRange(1, 1, sheet.getMaxRows(),sheet.getMaxColumns()).clear();
  var labelRange = sheet.getRange("A3:D3");
  labelRange.setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment("center").setBackgroundColor('darkgrey')
    .setValues([['Type','Story','state','owner']]);
  
  sheet.setColumnWidth(1, WIDTH[1]);
  sheet.setColumnWidth(2, WIDTH[2]);
  sheet.setColumnWidth(3, WIDTH[3]);
  sheet.setColumnWidth(4, WIDTH[4]);
 
  var startDate = getIterationDate(pData.start, 'start', pid);
  
  var startDateColumn = LEFT + 1;
  if(itNum > 24) itNum = 24;
  var diffDate = itNum * 7;
    
  diffDate = Math.floor(diffDate);
  Logger.log('diffDate: '+diffDate);
  var endDateColumn = startDateColumn + diffDate;
  
  var week = new Array("日", "月", "火", "水", "木", "金", "土");
  var month = new Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec");
  
  sheet.getRange(2, startDateColumn, row_num, diffDate).setBorder(true, true, true, true, true, true);
  sheet.getRange(2, startDateColumn, 1, diffDate).setNumberFormat("d");
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).breakApart();
  for (var i = 0, j = 0; i < diffDate; i++) {
    sheet.setColumnWidth(startDateColumn + i, WIDTH[0]);
    var newDate = new Date(startDate.getYear(), startDate.getMonth(), startDate.getDate() + i);
    var day = week[newDate.getDay()];
    if (i==0 || newDate.getDate() == 1) {
      sheet.getRange(1, startDateColumn + i).setValue(month[newDate.getMonth()]);
    } else {
      // 月のセルを結合
      var mergeRangeWidth = newDate.getDate(); // 結合するセルの幅
      var firstDayColumn = startDateColumn + i - (mergeRangeWidth - 1); // 月の初めのセルの列
      if(firstDayColumn<startDateColumn){
        firstDayColumn = startDateColumn;
        mergeRangeWidth = mergeRangeWidth - startDate.getDate()+1;
      }
      sheet.getRange(1, firstDayColumn, 1, mergeRangeWidth).mergeAcross();
    }
    if (day == "土" || day == "日") {
      sheet.getRange(2, startDateColumn + i, row_num, 1).setBackgroundColor(COLOR_HOLIDAY);
    }
    else {
      sheet.getRange(2, startDateColumn + i, row_num, 1).setBackgroundColor(COLOR_WEEKDAY);
    }
    
    sheet.getRange(2, startDateColumn + i).setFontSize(8);
    sheet.getRange(2, startDateColumn + i).setValue(newDate);
    sheet.getRange(3, startDateColumn + i).setValue(day);
  }
  sheet.getRange(1, startDateColumn + i).setValue('→');
}
  
  

// 新規シートの作成
function createNewSheet(title, pid){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.insertSheet(title);
  var res = Browser.msgBox(title + 'のガントチャートを書き込みます', Browser.Buttons.OK_CANCEL);
  if(res == 'ok')
    return rewrite();
  else
    ss.toast('Canceled.');
}
    
// 今日の日付を色付け
function setColorOnToday(row_num) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  var today = Utilities.formatDate(new Date(), "GMT+9:00", "yyyy/MM/dd");
  var lastRange = sheet.getRange(2, sheet.getLastColumn());
  var dateData = sheet.getRange("E2:" + lastRange.getA1Notation()).getValues();
  var todayColumn;
  
  for (var i = 0; i < dateData[0].length; i++) {
    var aDate = Utilities.formatDate(new Date(dateData[0][i]), "GMT+9:00", "yyyy/MM/dd");
    // 日付比較をして、今日の日付がある列を探す
    if (today == aDate) {
      todayColumn = 5 + i;
      break;
    }
  }
  
  // 今日の日付があれば、色を塗る
  if (todayColumn != undefined) {
    sheet.getRange(2, todayColumn, row_num, 1).setBackgroundColor(COLOR_TODAY_BACKGROUND).setFontColor(COLOR_TODAY_TEXT);
  }
}


/////////////////////////////////////
// シートへの書き込み
/////////////////////////////////////

//シートの更新処理
function rewrite(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sName = sheet.getName();
  var pid = ScriptProperties.getProperty(sName);
  var db = ScriptDb.getMyDb();
  if(sName == MSHEET){
    writeProjectList();
    ss.toast(MSHEET +' is reloaded');
    return 1;
  }
  
  var res = Browser.msgBox(sName + ' の更新', 'データベースを更新しますか？', Browser.Buttons.YES_NO_CANCEL);
  switch (res){
    case 'yes':
      ss.toast('Reloading '+sName+' start');
      reloadSingle();
      ss.toast('Reloading '+sName+' done');
      break;
    case 'no':
      break;
    case 'cancel':
      ss.toast('Canceled');
      return 1;
    default:
      break;
  }
  
  var mData = getManageData();
  if(mData == 0)
    return mData;
  for(var i in mData){
    if(i == pid)
      var pData = mData[i];
  }
  
  if(pData == undefined){
    Browser.msgBox(sName + ' はアクティブではありません． [' + MSHEET + ']　を確認してください');
    return 1;
  }
  
  var row_num = Utilities.jsonParse(ScriptProperties.getProperty(pid)).counter + HEAD -1;
  formGantt(pData,row_num);
  writeSingle(pData,row_num);
  setColorOnToday(row_num);
  return sheet;
}

// アクティブシートへ書き込み
function writeSingle(pData,row_num){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var pid     = pData.ID;
  var sName   = pData.projectName;
  var start   = pData.start;
  var end     = pData.end;
  var current = pData.current;
  
  writeTitle(pData);
  writeUpdateTime(pid);
  
  ss.toast("Writing start");
  writeStories(sheet, pData);
  setColorOnToday(row_num);
  ss.toast("Writing done!");
}

// ストーリー書き込み
function writeStories(sheet, pData){
  var current = pData.current;
  var start   = pData.start;
  var end     = pData.end;
  var pid     = pData.ID;
  var db = ScriptDb.getMyDb();

  var story = db.query({type:'story',project_id:pid,iteration:db.between(start, end+1)})
  .sortBy('estimate', db.NUMERIC)
  .sortBy('story_type', db.LEXICAL)
  .sortBy('iteration', db.ASCENDING, db.NUMERIC);
  var row = 0; var n = start-1;
  var cols = [1,2];
  while (story.hasNext()){
    var single = story.next();
    var itNum = Number(single.iteration);
    if(itNum != n){
      while(itNum > n){
        writeIteration(sheet,current,row,n+1,pid);
        n++;
        row++;
      }
      cols = getIterationTerm(sheet, itNum, pid);
    }
    writeSingleStory(sheet, row, single, pid);
    drawIterationColor(sheet, row, cols, itNum, pid);
    var type = single.story_type;
    if (type == 'release'){
      var deadline = single.deadline;
      var deadDate = new Date(Date.parse(deadline));
      drawReleaseColor(sheet, row, deadline, pid);
    }
    row++;
  }
}

function writeTitle(pData){
  var start = pData.start;
  var end   = pData.end;
  var title = pData.projectName;
  var range = getArea('title');
  var cell1 = range.getCell(1,1);
  var cell2 = range.getCell(1,2);
  var cell3 = range.getCell(1,3);
  range.setBackgroundColor(COLOR.titlebg).setFontColor(COLOR.titlefg);
  cell1.setValue('Project:').setHorizontalAlignment('right');
  cell2.setValue(title).setHorizontalAlignment('center').setFontSize(15)
  .setFontFamily('cursive').setFontWeight('bold');
  cell3.setValue(start+' - '+end);
}

function writeSingleStory(sheet, row, story, pid){
  row += (HEAD + 1);
  var order = getQuery('story').cNames;
  var range = sheet.getRange(row, 1, 1, order.length).setBackgroundColor('white')
  .setFontColor('black');
  range.breakApart().setBorder(true, true, true, true, true, false);
  for( var j=0; j<order.length; j++ ){
    var key = order[j];
    var val = story[key];
    writeCell(sheet, row, j+1, val);
    if(key == 'name') {
      var url = story.url;
      addLink(sheet, row, j+1, url);
    }
  }
  drawMark(sheet, row, story, pid);
}

function writeIteration(sheet,current,row,itNum,pid){
  var str = ''; var color = '';
  if(itNum <  current){
    str = 'Done';    color = COLOR.done;
  }
  if(itNum == current){
    str = 'Current'; color = COLOR.current;
  }
  if(itNum >  current){
    str = 'Backlog'; color = COLOR.backlog;
  }
  writeLabel(sheet, row, 1, LEFT , [[itNum,str,'','']], color);
  drawIterationChart(sheet, row, itNum, color, pid);
}

function writeUpdateTime(pid){
  pid=638215;
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = getArea('update');
  var date  = Utilities.jsonParse(ScriptProperties.getProperty(pid)).update;
  date = new Date(Date.parse(date));
  var ymd = date.getFullYear()+"/"+(date.getMonth()+1)+"/"+date.getDate();
  var time = date.getHours()+":"+date.getMinutes();
  range.setHorizontalAlignment('right').setValues([['Reloaded:',ymd,time,'']]);
}


function getIterationTerm(sheet, itNum, pid){
  var startDate = getIterationDate(itNum,'start',pid);
  var finishDate = getIterationDate(itNum,'finish',pid);
  var colStart = getColumnByDate(startDate);
  var colFinish = getColumnByDate(finishDate);
  if(colFinish == colStart)colFinish++;
  var cols = [colStart, colFinish];
  return cols;
}  

function writeLabel(sheet, row, col, colNum, str, color ){
  row += (HEAD +1);
  var range = sheet.getRange(row, col, 1, colNum);
  range.clear().clearFormat();
  range.breakApart();
  range
    .setBackgroundColor(color).setFontColor('white')
    .setHorizontalAlignment('center').setFontFamily('arial black')
    .setBorder(true,true,true,true,true,true)
    .setValues(str);
}

function writeCell(sheet,row,col,val){
  var cell = sheet.getRange(row,col);
  cell.setValue(val);
}

function addLink(sheet, i, j, url){
  var cell = sheet.getRange(i,j);
  var str = cell.getValue();
  cell.setFormula('=hyperlink("'+url+'","'+str+'")')
    .setHorizontalAlignment('center');
}


/////////////////////////////////////
// scriptDbまわり
/////////////////////////////////////

// プロジェクト単体のデータを更新
function reloadSingle(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var pid = ScriptProperties.getProperty(sheet.getName());
  var db = ScriptDb.getMyDb();
  var all = db.query({type:db.anyOf(['iteration','story']),project_id:pid});
  ss.toast('please waite...');
  while(all.hasNext()){
    var record = all.next();
    db.remove(record);
  }
  var iteration  = getData('iteration',pid);
  var story = storyIntoIteration(iteration,pid);
  saveObject(iteration, 'iteration');
  saveObject(story    , 'story');
  var obj = {update: new Date()};
  addProperty(pid,obj);
  countStories(pid);
}

// プロジェクトリストのデータを更新
function reloadProjectList(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Reloading projects start');
  var db = ScriptDb.getMyDb();
  
  var all = db.query({type:'project'});
  while(all.hasNext()){
    var project = all.next();
    var res = db.remove(project);
  }
  var pList = getProjectList();
  ss.toast('Reloading projects done');
  return pList;
}

// データベースの全データをリセット
function reloadAllData(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var db = ScriptDb.getMyDb();
  
  ss.toast('please waite...');
  resetDb();
  var mData = getManageData('select');
  
  var pList = getProjectList();
  for(var i in mData){
    ss.toast('Reloading '+ mData[i].projectName);
    var pid = i;
    var iteration  = getData('iteration',pid);
    var story = storyIntoIteration(iteration,pid);
    saveObject(iteration, 'iteration');
    saveObject(story    , 'story');
    var obj = {update: new Date()};
    addProperty(pid,obj);
  }
  for(var i in mData){
    countStories(i);
  }
  ss.toast('Done');
}

// 選択状態プロジェクトのストーリーデータを更新
function reloadStories(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var db = ScriptDb.getMyDb();
  var all = db.query({type:db.anyOf(['iteration','story'])});
  ss.toast('please waite...');
  while(all.hasNext()){
    var record = all.next();
    db.remove(record);
  }
  var mData = getManageData('select');
  for(var i in mData){
    ss.toast('Reloading '+ mData[i].projectName);
    var pid = i;
    var iteration  = getData('iteration',pid);
    var story = storyIntoIteration(iteration,pid);
    saveObject(iteration, 'iteration');
    saveObject(story    , 'story');
    var obj = {update: new Date()};
    addProperty(pid,obj);
  }
  for(var i in mData){
    countStories(i);
  }
  ss.toast('Done');
}

// プロジェクトリストを取得
function getProjectList(){
  var list = {};
  var projects = getData('project');
  for ( var i in projects ) {
    list[i] = projects[i];
  }
  saveObject(projects, 'project');
  return list;
}


/////////////////////////////////////////////////
// UIまわり
/////////////////////////////////////////////////

// プロジェクトの選択
function selectProject(){
  var userProp = UserProperties.getProperties(); 
  
  var token = userProp.pivotal_token;
  // ユーザーデータの取得
  if(token == null || token == false){
    setUserData();
  }
  
  // プロジェクトIDを選択
  checkProjectUi();
}

// プロジェクト管理シートの書込
function writeProjectList(){
  var db = ScriptDb.getMyDb();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sp = ScriptProperties.getProperties();
  var res = Browser.msgBox(MSHEET + ' の更新', 'データベースを更新しますか？', Browser.Buttons.YES_NO_CANCEL);
  switch (res){
    case 'yes':
      ss.toast('Reloading '+MSHEET+' start');
      reloadProjectList();
      ss.toast('Reloading '+MSHEET+' done');
      break;
    case 'no':
      break;
    case 'cancel':
      ss.toast('Canceled');
      return 1;
    default:
      break;
  }
  var sheets = ss.getSheets();
  var k = -1;
  for(var i=0; i<sheets.length; i++){
    if(sheets[i].getSheetName() == MSHEET){
      k = i;
      var mSheet = sheets[i];
    }
  }
  if(k == -1){
    var mSheet = ss.insertSheet(MSHEET);
    mSheet.deleteColumns(PARANUM+1, mSheet.getMaxColumns()-PARANUM);
    mSheet.deleteRows(31, mSheet.getMaxRows()-30);
  }
  var projects = db.query({type:'project'}).sortBy('id',db.NUMERIC);
  var label = [['flag','ID','projectName','current','iterationLength','lastActivity','doneToShow']];
  var labelRange = mSheet.getRange(2,1,1,label[0].length)
  .setBackgroundColor('lightgray').setValues(label);
  
  var data = [];
  var i = 0;
  while(projects.hasNext()){
    var pro = projects.next();
    var array = [];
    array[0] = pro.id;
    array[1] = pro.name;
    array[2] = pro.current_iteration_number;
    array[3] = pro.iteration_length;
    array[4] = new Date(Date.parse(pro.last_activity_at));
    array[5] = pro.number_of_done_iterations_to_show;
    if( 'doneToShow' in sp ) array[5] = sp.doneToShow;
    data[i] = array;
    i++
  }
  var num = Number(i);
  var range = mSheet.getRange(3,2,num,array.length);
  range.setValues(data).setBorder(true,true,true,true,true,true);
}

// プロジェクト選択UI
function checkProjectUi(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication();
  writeProjectList();
  var mData = getManageData('all');
  var wrapper = app.createVerticalPanel();
  for(var i in mData){
    var checkBox = app.createCheckBox(mData[i].projectName).setId(mData[i].ID)
    .setName(mData[i].projectName).setValue(mData[i].flag == true);
    wrapper.add(checkBox);
  }
  var label = app.createLabel('Doneの最大表示数')
  var textBox = app.createTextBox().setName('doneToShow').setValue(mData[i].doneToShow);
  wrapper.add(label).add(textBox);
  var panels = [wrapper];
  var tmp = Template(panels,'プロジェクトの選択','okCheckProject',wrapper);
  ss.show(tmp);
}

// OKの場合発動
function okCheckProject(e){
  var app = UiApp.getActiveApplication();
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var mSheet = ss.getSheetByName(MSHEET);
  
  var mData = getManageData('all');
  var names = [];
  var j = 0;
  for(var i in mData){
    var tmp = mData[i].projectName;
    ScriptProperties.setProperty(tmp, i);
      if(e.parameter[tmp] == 'true'){
        names[j] = tmp;
        j++;
    }
  }
  ScriptProperties.setProperty('doneToShow', e.parameter.doneToShow)
  makeBeTrue(mSheet,names,e.parameter.doneToShow);  
  Browser.msgBox(String(j) +' projects activated');
  return app;
}

// 入力フォームのテンプレート
function Template(panels, title, func, callback) {
  var app = UiApp.getActiveApplication();
  var titlePanel = app.createHorizontalPanel();
  var label = app.createLabel(title);
  titlePanel.add(label);
  
  var buttonPanel = app.createHorizontalPanel();
  var okButton = app.createButton('OK');
  var okButtonHandler = app.createServerHandler(func).addCallbackElement(callback);
  okButton.addClickHandler(okButtonHandler);
  buttonPanel.add(okButton);
  
  var cancelButton = app.createButton('Cancel');
  var cancelButtonHandler = app.createServerHandler("cancelButtonPush");
  cancelButton.addClickHandler(cancelButtonHandler);
  buttonPanel.add(cancelButton);
  
  
  var n = panels.length +2;
  var h = 300;
  var w = 300;
  app.setWidth(w).setHeight(h);
  app.add(titlePanel);
  for(var i=0; i<panels.length; i++){
    app.add(panels[i]);
  }
  app.add(buttonPanel);
  
  return app;
}

// キャンセルの場合発動
function cancelButtonPush(){
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}


////////////////////////////////////////
// subroutine
////////////////////////////////////////

function getArea(key){
  var sheet = SpreadsheetApp.getActiveSheet();
  var array = getAreaData(sheet,key);
  var range = sheet.getRange(array[0],array[1],array[2],array[3]);
  return range;
}


function getAreaData(sheet,key) {
  var sName = sheet.getName();
  var pid   = ScriptProperties.getProperty(sName);
  var mData = getManageData();
  var itNum = mData[pid].end - mData[pid].start + 1;
  if(itNum > 24) itNum = 24;
  var lastCol = LEFT + itNum * 7 + 1;
  
  var headerRow = HEAD;
  var leftCol = LEFT;
  
  var rightCol = lastCol - leftCol + 1;
  var mainRowTop = headerRow +1;
  var mainColTop = leftCol +1;
  var mainCol = lastCol - leftCol;
  
  // 各項目のエリア設定　[開始行, 開始列, 行数, 列数]
  var range = {
    'title'  : [1, 1, 1, 4],
    'update' : [2, 1, 1, 4],
    'label'  : [headerRow, 1, 1, leftCol],
    'month'  : [1, leftCol+1, 1, rightCol],
    'date'   : [2, leftCol+1, 1, rightCol],
    'day'    : [3, leftCol+1, 1, rightCol],
  };  
  
  return range[key];
}

function getData(type,pid){
  var query = getQuery(type,pid);
  var els = getElementFromApi(query.url);
  var parent = els[query.type];
  var obj = {};
  for (var i in parent){
    if(parent.length == undefined){
      obj[i] = getChild(parent);
      Logger.log('!!!except!!! not Array');
      continue;
    }
    obj[i] = getChild(parent[i]);
  }
  return obj;
}

function getQuery(type,pid){
  var queries = {
project : {
  url : 'projects',
  type : 'project',
  cNames : ['id','name','iteration_length']
  },
iteration : {
  url : 'projects/'+pid+'/iterations',
  type : 'iteration',
  cNames : ['id','number','stories','start','finish']
  },
story : {
  url : 'projects/'+pid+'/stories',
  type : 'story',
  cNames : ['story_type','name','current_state','owned_by']
  }
};
  return queries[type];
}

function storyIntoIteration(iterations,pid){
  var array = [];
  var cnt = 0;
  var i = 0;
  for(var i in iterations){
    iterations[i].project_id = pid;
    var stories = iterations[i].stories;
    for(var j in stories){
      stories[j].iteration = iterations[i].number;
      array[cnt] = stories[j];
      cnt++;
    }
    delete iterations[i].stories;
  }
  return array;
}
  

function saveObject(obj, type){
  var db = ScriptDb.getMyDb();
  for( var i in obj ){
    var record = obj[i];
    record.type = type;
    var stored = db.save(record);
  }
}

function countStories(pid){
  var db = ScriptDb.getMyDb();

  var mData = getManageData();
  var start = mData[pid].start;
  var end   = mData[pid].end;
  var story     = db.query({type:'story' ,project_id:pid,iteration:db.between(start, end+1)}).getSize();
  var iteration = db.query({type:'iteration',project_id:pid,number:db.between(start, end+1)}).getSize();
  var cnt = story + iteration;
  Logger.log('counter: '+cnt);
  var obj = {counter:cnt};
  addProperty(pid,obj);

  return cnt;
}
  
function getCurrent(pid){
  var db = ScriptDb.getMyDb();
  var project = db.query({type:'project',id:pid});
  var current = project.next().current_iteration_number;
  current = Number(current);
  return current;
}


function getChild(parent){
  var child = {};
  var parents = parent.getElements();
  for ( var i in parents ) {
    var tmp = parents[i];
    var Name = tmp.getName().getLocalName();
    if(tmp != null) {
      if(tmp.type == 'array'){
        tmp = tmp.getElements();
        var array = [];
        for( var j in tmp ){
          array[j] = arguments.callee(tmp[j]);
        }
        tmp = array;
      }
      else tmp = tmp.Text;
    }
    child[Name] = tmp;
  }
  return child;
}

function resetDb(){
  var db = ScriptDb.getMyDb();
  var all = db.query({});
  while( all.hasNext() ){
    var obj = all.next();
    var res = db.remove(obj);
  }
}

function makeBeTrue(sheet,pNames,limit){
  var rowNum = sheet.getLastRow()-2;
  for(var i=1; i<=rowNum; i++){
    var row = sheet.getRange(i+2,1,1,PARANUM);
    var rowName = row.getCell(1, 3).getValue();
    row.getCell(1,1).setValue(false);
    row.setBackgroundColor('white').setBorder(true,true,true,true,true,true);
    for(var j=0; j<pNames.length; j++){
      if(rowName == pNames[j]){
        row.getCell(1,1).setValue(true);
        row.setBackgroundColor('gold');
        row.getCell(1,7).setValue(limit);
      }
    }
  }
}      


function getManageData(arg){
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var N = PARANUM;
  var mSheet = ss.getSheetByName(MSHEET);
  if(mSheet == null){
    Browser.msgBox('シート [' + MSHEET + '] がありません．「プロジェクト選択」を実行してください');
    return 1;
  }
  var mData = mSheet.getRange(3,1,20,N).getValues();
  var mLabels = mSheet.getRange(2,1,1,N).getValues()[0];

  var data = [];
  for(var i=0; mData[i][1]!=''; i++){
    var obj = {};
    var pid = mData[i][1];
    for(var j=0; j<N; j++){
      obj[mLabels[j]] = mData[i][j];
    }
    
    switch (arg){
      case 'all':
        data[pid] = obj;
        break;
        
      case 'select':
        if(obj.flag == true)
          data[pid] = obj;
        break;
        
      default:
        if(obj.flag == true){
          var start = obj.current - obj.doneToShow;
          if(start > 1){
            obj.start = start;
          }
          else
            obj.start = 1;
          var end = getLastIteration(String(obj.ID));
          if( end == 0 ) {
            Browser.msgBox(obj.projectName + 'のデータがありません．データベースを更新してください');
            return 0;
          }
          
          if((end-start) <= LIMIT)
            obj.end = end;
          else
            obj.end = obj.start + LIMIT -1;
    
          data[pid] = obj;
        }
        break;
    }
  }
  return data;
}

function getLastIteration(pid){
  var db = ScriptDb.getMyDb();

  var iteration = db.query({type:'iteration',project_id:pid}).sortBy('number',db.DESCENDING,db.NUMERIC);
  var lastIteration = iteration.next();
  if ( lastIteration == null ){
    Logger.log('!!! No iteration in DB !!!');
    return 0;
  }
  var num = Number(lastIteration.number);
  return num;
}
  


function drawMark(sheet, row, story){
  var cDate = story.created_at;
  var uDate = story.updated_at;

  
  var cDate = new Date(Date.parse(cDate));
  var uDate = new Date(Date.parse(uDate));
  
  var cCol = getColumnByDate(cDate);
  if (cCol !== undefined){
    var cell = sheet.getRange(row, cCol+LEFT);
    cell.setValue('C').setFontWeight('bold').setFontSize(12).setFontColor(COLOR.create);
  }
  
  if( (uDate -cDate)/(1000 * 60 * 60 * 24) < 1 )
  var uCol = cCol;
  else
    var uCol = getColumnByDate(uDate);
  if (uCol !== undefined){
    var cell = sheet.getRange(row, uCol+LEFT);
    cell.setValue('U').setFontWeight('bold').setFontSize(12).setFontColor(COLOR.update);
  }
  
  if('accepted_at' in story) {
    var aDate = story.accepted_at;
    var aDate = new Date(Date.parse(aDate));
    var aCol = getColumnByDate(aDate);
    if (aCol !== undefined){
      cell = sheet.getRange(row, aCol+LEFT);
      cell.setValue('A').setFontWeight('bold').setFontSize(12).setFontColor(COLOR.accept);
    }
  }
}

function drawReleaseColor(sheet, row, deadline, pid){
  row += (HEAD+1);
  var date = new Date(Date.parse(deadline));
  var col = getColumnByDate(date);
  var range = sheet.getRange(row, 1, 1, col+LEFT);
  range.setBackgroundColor(COLOR.deadline);
  drawDeadLine(sheet, row, col);
}

function drawDeadLine(sheet, row, col){
  var rowNum = row-(HEAD+1);
  var range = sheet.getRange(HEAD+1, col+LEFT, rowNum);
  range.setBackgroundColor(COLOR.deadline);
  
}

function drawIterationChart(sheet, row, itNum, color,pid){
  var finishDate = getIterationDate(itNum,'finish',pid);
  var colFinish = getColumnByDate(finishDate)-1;
  var range = sheet.getRange(row+HEAD+1, 1, 1, colFinish+LEFT);
  range.setBackgroundColor(color);
}

function drawIterationColor(sheet, row, cols, itNum, pid){
  var current = getCurrent(pid);
  var color = '';
  if(itNum <  current){
    color = COLOR.done;
  }
  if(itNum == current){
    color = COLOR.current;
  }
  if(itNum >  current){
    color = COLOR.backlog;
  }

  var range = sheet.getRange(row+HEAD+1, cols[0]+LEFT, 1, cols[1]-cols[0]);
  range.setBackgroundColor(color);
}

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

function getIterationDate(number, path, pid){
  var db = ScriptDb.getMyDb();
  var iteration = db.query({type:'iteration',project_id:pid});
  while(iteration.hasNext()){
    var tmp = iteration.next();
    if(tmp.number == number)
      var record = tmp;
  }
  var dateString = record[path];
  
  var date = new Date(Date.parse(dateString));
  return date;
}

function judgeDate(date1, date2){
  var res = 0;
  if((date1.getYear() == date2.getYear()) 
    && (date1.getMonth() == date2.getMonth())
    &&(date1.getDate() == date2.getDate())){
      res = 1;
    }
    return res;
}


////////////////////////////////////
// utility
///////////////////////////////////
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

function merge(){
  var args = Array.prototype.slice.call(arguments),
      len = args.length,
      ret = {},
      itm;
  
  for( var i = 0; i < len ; i++ ){
    var arg = args[i];
    for (itm in arg) {
      if (arg.hasOwnProperty(itm))
        ret[itm] = arg[itm];
    }
  }
  return ret;
}

// vim: set ft=javascript ts=2 sw=2 sts=2 et :
