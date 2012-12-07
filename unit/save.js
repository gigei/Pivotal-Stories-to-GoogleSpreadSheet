function saveObject(obj, type){
  var db = ScriptDb.getMyDb();
  
  for( var i in obj ){
    var record = obj[i];
    record.type = type;
    db.save(record);
  }
}

function saveObjectToDb(){
    var pid = $PID,
        story  = getData('story',pid);
    
    saveObject(story , 'story');
}

function getObjectFromDb(){
    var pid = $PID,
        db = ScriptDb.getMyDb(),
        obj = {},
        stories;
    
    stories = db.query({type:'story', project_id:pid, story_type:'release'});
    while(stolies.hasNext()){
        var record = stories.next();
        obj[record.id] = record;
    }
    return obj;
}

function reloadStory(){
    var pid = $PID,
        db = ScriptDb.getMyDb(),
        all,
        story;
    
    all = db.query({type:'story', project_id:pid});
    
    while(all.hasNext()){
        var record = all.next();
        db.remove(record);
    }
    
    story  = getData('story',pid);
    saveObject(story , 'story');
}

