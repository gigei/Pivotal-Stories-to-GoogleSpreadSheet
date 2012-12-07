function getQuery(type,pid){
  var queries = {
project : {
  url : 'projects',
  type : 'project'
  },
iteration : {
  url : 'projects/'+pid+'/iterations',
  type : 'iteration'
  },
story : {
  url : 'projects/'+pid+'/stories',
  type : 'story'
  }
};
  return queries[type];
}

function getData(type,pid){
  var query = getQuery(type,pid),
      els = getElementFromApi(query.url),
      parent = els[type],
      obj = {};
  
  for (var i in parent){
    if(parent.length == undefined){
      obj[i] = getChild(parent);
      continue;
    }
    obj[i] = getChild(parent[i]);
  }
  return obj;
}

function getChild(parent){
  var child = {},
      parents = parent.getElements();
  for ( var i in parents ) {
    var tmp = parents[i],
        Name = tmp.getName().getLocalName();
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

