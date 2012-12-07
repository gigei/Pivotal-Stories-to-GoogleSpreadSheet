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

