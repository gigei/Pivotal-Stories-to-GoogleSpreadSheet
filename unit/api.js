// APIトークンの取得
function getToken(user,pass){
  var apiUrl = 'https://www.pivotaltracker.com/services/v3/';

    // リクエストの中身
    var url = 'tokens/active',
        auth_data = Utilities.base64Encode(user + ':' + pass),
        headers = {'Authorization' : 'Basic ' + auth_data},
        params = {'headers' : headers};

  // fetch
  var res = UrlFetchApp.fetch(apiUrl+url , params).getContentText();

  // XML parse
  var elem = Xml.parse(res, true).getElement();
  var token = elem.getElement('guid').getText();  //tokenを取得
  return token;
}

// API fetch & XML parse
function getElementFromApi(extUrl) {
  var apiUrl = 'https://www.pivotaltracker.com/services/v3/',
      token = UserProperties.getProperty('pivotal_token'),
      headers = {'X-TrackerToken'  : token,
        'X-Tracker-Use-UTC': true},
      params ={method : 'get',
        headers: headers};
  // fetch
  var res = UrlFetchApp.fetch(apiUrl+extUrl , params);
  var txt = res.getContentText();
  // XML parse
  var xml = Xml.parse(txt , true);
  var elem = xml.getElement();
  return elem;
}

function useApiTest(){
  var apiUrl = 'https://www.pivotaltracker.com/services/v3/',
      extUrl = 'projects',
      token = getToken($USER, $PASS),
      headers = {'X-TrackerToken'  : token},
      params ={method : 'get', headers: headers};
      
  // fetch
  var res = UrlFetchApp.fetch(apiUrl + extUrl , params),
      txt = res.getContentText(),
      elem = Xml.parse(txt , true).getElement();
  return elem;
}

// vim: set st=2 ts=2 sts=2 et
