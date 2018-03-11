
if (!String.prototype.startsWith) {
    String.prototype.startsWith = function(searchString, position){
      return this.substr(position || 0, searchString.length) === searchString;
  };
}

reposTypes = [
  'github',
  'bitbucket',
  'gitlab',
  ]

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('1 Connect to JIRA', 'showConnectToJira')
      .addItem('2 Build JQL Query', 'showBuildJqlQuery')
      .addItem('3 Display Options', 'showDisplayOptions')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('4 Get JQL Result')
        .addItem('Send Request', 'sendRequest')
        .addItem('Clean Sheet and Send Request', 'cleanSheetAndSendRequest')
        .addItem('Send Request and Insert Data to New Sheet', 'sendRequestAndInsertToNew'))
      .addSeparator()
      .addItem('Setup Storage', 'showSetUpStorage')
      .addSeparator()
      .addItem('Convert to JIRA Link', 'showConvertToJiraLink')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function _getId() {
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}

function sendGaEvent(eventName){
  var payload = {
    'v': '1',
    'tid': 'UA-106946460-1',
    'cid': '555',
    't': 'pageview',
    'dp': '/APP SCRIPT/' + eventName,
  };

  var options = {
    'method' : 'post',
    'payload' : payload
   };

  // Sending the hit
  UrlFetchApp.fetch('http://www.google-analytics.com/collect', options);
}

function showSidebar(name, title) {
  var ui = HtmlService.createHtmlOutputFromFile(name)
      .setTitle(title);
  SpreadsheetApp.getUi().showSidebar(ui);
}

function showConnectToJira() {
  showSidebar('ConnectToJira', 'JIRA2PM :: 1 Connect to JIRA');
}

function showBuildJqlQuery() {
  showSidebar('BuildJqlQuery', 'JIRA2PM :: 2 Build JQL Query');
}

function showDisplayOptions() {
  showSidebar('DisplayOptions', 'JIRA2PM :: 3 Display Options');
}

function showSetUpStorage() {
  showSidebar('SetUpStorage', 'JIRA2PM :: Setup Storage');
}

function showConvertToJiraLink() {
  showSidebar('ConvertToJiraLink', 'JIRA2PM :: Convert to JIRA Link');
}

// Saves current sheet to the user preferences, which can be used later
// Returns SheetId
function _saveSheetId() {
  sheetId = SpreadsheetApp.getActiveSheet().getSheetId();
  PropertiesService.getUserProperties().setProperty('activeSheetId', sheetId);

  return sheetId;
}

// Saves target sheetId to properties
function _saveSheetId(sheetId) {
  PropertiesService.getUserProperties().setProperty('activeSheetId', sheetId);
  return sheetId;
}

// Sets user last desired action\
// Possible options - regular, clear, insert_new
function _setLastAction(action) {
  PropertiesService.getUserProperties().setProperty('lastAction', action);
  return action;
}

function _getLastAction() {
  var action = PropertiesService.getUserProperties().getProperty('lastAction');
  if (action == null)
    return 'regular';
  return action;
}

// Returns sheet with sheetId in current spreadsheet.
// Return null if sheet was not find
function _getSheetById(sheetId) {
  var sheets = SpreadsheetApp.getActive().getSheets();
  for (var n in sheets)
    if (sheets[n].getSheetId() == sheetId)
      return sheets[n];

  return null;
}

// Uses last user action with selected page of results
function useLastAction(pageJSON) {
  // Get sheet ID
  var sheetId = 0;
  var page = JSON.parse(pageJSON);

  // Initialization for first page
  if (page.startAt == 0) {
    var action = _getLastAction();
    // Diferent actions for different user actions
    switch (action) {
      default:
      case "regular":
        sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        break;
      case "clear":
        sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        sheet.clear();
        break;
      case "insert_new":
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
        break;
    }
    sheetId = _saveSheetId(sheet.getSheetId());
  } else {
    sheetId = PropertiesService.getUserProperties().getProperty('activeSheetId');
    sheet = _getSheetById(sheetId);
  }

  return _fetchJira(page);
}

function sendRequest() {
  _setLastAction('regular');
  showSidebar('ProgressMonitor.html', 'JIRA2PM :: Progress Monitor');
}

function cleanSheetAndSendRequest() {
  _setLastAction('clear');
  showSidebar('ProgressMonitor.html', 'JIRA2PM :: Progress Monitor');
}

function sendRequestAndInsertToNew() {
  _setLastAction('insert_new');
  showSidebar('ProgressMonitor.html', 'JIRA2PM :: Progress Monitor');
}

function getCurrentRangeValues() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return JSON.stringify(sheet.getActiveRange().getValues());
}

function convertKeysToLinks(keysJSON) {
  var keys = JSON.parse(keysJSON);
  var connectOptions = JSON.parse(PropertiesService.getUserProperties().getProperty('connectOptions'));

  if (!connectOptions.baseURL)
    throw ("No connection options were found. Please connect to JIRA first");

  return connectOptions.baseURL + "issues/?jql=key%20in%20(" + encodeURIComponent(keys.join()) + ")";
}

function getOptions(name) {
  var props = PropertiesService.getUserProperties();

  // Try getting newer version of local user options
  var options = props.getProperty(_getId() + 'options');
  if (options != null) {
    options = JSON.parse(options);
    if (options[name])
      return options[name];
  }

  // Try getting newer version of user gloabl options
  options = props.getProperty('globalOptions');
  if (options != null) {
    options = JSON.parse(options);
    if (options[name])
      return options[name];
  }

  // Try getting older version of local user options
  var options = props.getProperty(_getId() + name);
  if (options != null)
    return options;

  // Try getting older version of global user options
  return props.getProperty(name);
}

function setOptions(name, value) {
  var props = PropertiesService.getUserProperties();

  // Set to local options first
  var options = props.getProperty(_getId() + 'options');
  options = JSON.parse(options);
  if (options == null)
    options = {}
  options[name] = value;
  props.setProperty(_getId() + 'options', JSON.stringify(options))

  // Set global options
  var options = props.getProperty('globalOptions');
  options = JSON.parse(options);
  if (options == null)
    options = {}
  options[name] = value;
  props.setProperty('globalOptions', JSON.stringify(options))
}

function _getCustomFields() {
  var customFieldsCount = getOptions('customFieldsCount');

  // Legacy custom fields support
  if (customFieldsCount === null)
    return getOptions('customFields');

  var customFields = '';
  for (var i = 0; i < customFieldsCount; i++)
    customFields += getOptions('customFields_' + i)

  return customFields;
}

function getJqlPreferences() {
  return JSON.stringify({jqlOptions: getOptions('jqlOptions'), customFields: _getCustomFields()});
}

function connectJira(optionsJSON, localOptionsJSON) {
  var options = JSON.parse(optionsJSON);

  var baseURL = options.baseURL;

  if (['/', '\\'].indexOf(baseURL.slice(-1)) < 0)
    baseURL = baseURL.concat('/');

  var ennCred = Utilities.base64Encode(options.username + ':' + options.password);

  var fetchArgs = {
    contentType: 'application/json',
    headers: {'Authorization':'Basic ' + ennCred},
    muteHttpExceptions: true
  };

  var httpResponse = UrlFetchApp.fetch(baseURL + 'rest/api/2/search', fetchArgs);
  if (httpResponse) {
    var responseCode = httpResponse.getResponseCode();
    if (responseCode != 200)
      throw "Can't connect!";
  }

  var connectOptions = {};
  connectOptions.username = options.username;
  connectOptions.baseURL = baseURL;
  connectOptions.ennCred = ennCred;
  setOptions('connectOptions', JSON.stringify(connectOptions));
  setOptions('localOptions', localOptionsJSON);

  return true;
}

function updateCustomFields() {
  sendGaEvent('updateCustomFields');
  var connectOptions = JSON.parse(getOptions('connectOptions'));
  if (connectOptions == null) throw 'Setup connection options';

  var customFields = {fields:[], fieldsNames:[]};

  var fetchArgs = {
    contentType: 'application/json',
    headers: {'Authorization':'Basic ' + connectOptions.ennCred},
    muteHttpExceptions: true
  };

  var url = connectOptions.baseURL + 'rest/api/2/field';

  var httpResponse = UrlFetchApp.fetch(url, fetchArgs);
  if (httpResponse) {
    var responseCode = httpResponse.getResponseCode();
    if (responseCode == 200) {
      var data = JSON.parse(httpResponse.getContentText());

      data.map(function(x){
        if (x.custom) {
          customFields.fields.push(getPathForType(x.schema.type, x.id));
          customFields.fieldsNames.push(x.name);
        }
      });
    }
  }

  customFields = JSON.stringify(customFields);

  customFieldsSplit = customFields.match(/.{1,9000}/g);
  setOptions('customFieldsCount', customFieldsSplit.length);
  for (var i = 0; i < customFieldsSplit.length; i++)
    setOptions('customFields_' + i, customFieldsSplit[i]);

  return customFields;
}

function _fetchJira(page) {
  connectOptions = JSON.parse(getOptions('connectOptions'));
  if (connectOptions == null) throw 'Setup connection options';

  jqlOptions = JSON.parse(getOptions('jqlOptions'));
  if (jqlOptions == null) throw 'Setup jql options';

  displayOptions = JSON.parse(getOptions('displayOptions'));
  if (displayOptions == null) throw 'Setup display options';

  setUpStorageOptions = JSON.parse(getOptions('setUpStorageOptions'));
  if (setUpStorageOptions == null) setUpStorageOptions = {refresh: null, refreshCount: 1, refreshMeasurement: 'hours', storage: 'global'};

  localOptions = JSON.parse(getOptions('localOptions'));
  customFields = JSON.parse(_getCustomFields());

  jqlOptions.fields = jqlOptions.fields.concat(jqlOptions.customFields.map(function(x){return x.value;}));
  jqlOptions.fieldsNames = jqlOptions.fieldsNames.concat(jqlOptions.customFields.map(function(x){return x.name;}));

  var baseURL = connectOptions.baseURL;
  var jql = jqlOptions.jql;

  var minDif = setUpStorageOptions.refreshCount;
  switch (setUpStorageOptions.refreshMeasurement) {
    case 'days':
      minDif *= 24;
    case 'hours':
      minDif *= 60;
    case 'minutes':
      minDif *= 60 * 1000;
  }

  baseURL = baseURL.concat('rest/api/2/search');
  var ennCred = connectOptions.ennCred;

  var fetchArgs = {
    contentType: 'application/json',
    headers: {'Authorization':'Basic ' + ennCred},
    muteHttpExceptions: true
  };

  // Encode jql
  jql = encodeURIComponent(jql);

  jql = '?jql='.concat(jql);
  if (jql.indexOf('maxResults') < 0 && jql.length > 0)
    if (page)
      jql = jql.concat('&startAt=' + page.startAt + '&maxResults=' + page.maxResults);
    else
      jql = jql.concat('&maxResults=1000');

  var continuePaging = false;
  var totalResults = 0;
  var curResult = 0;

  var httpResponse = UrlFetchApp.fetch(baseURL + jql, fetchArgs);
  if (httpResponse) {
    var responseCode = httpResponse.getResponseCode();
    if (responseCode == 200) {
      var data = JSON.parse(httpResponse.getContentText());

      if (page) {
        if (page.startAt == 0)
          updateHeadRow();

        appendJira_(data);

        continuePaging = data.total > data.maxResults + data.startAt;
        maxResults = data.total;
        curResult = Math.min(data.maxResults + data.startAt, data.total);
      }
      else {
        // Decide whether we need to append data or to update existant
        var lastTimestamp = getLastTimeStamp(minDif);
        updateHeadRow();
        if (lastTimestamp[1] == -1 || !setUpStorageOptions.refresh)
          appendJira_(data);
        else
          appendJira_(data, lastTimestamp[1]);
      }

      setupFilter(sheet);
    }
    else {
      switch(responseCode){
        case 401:
          throw 'Incorrect username or password';
          break;
        default:
          throw 'Wrong response. Server responded with code: ' + responseCode;
          break;
      }
    }
  }
  else
    throw 'Can not access server.';

  return JSON.stringify({continuePaging: continuePaging, maxResults: maxResults, curResult: curResult});
}

function appendJira_(data, fromIndex) {
  var issues = [];

  var timestamp = new Date();

  for(var id in data['issues'])
    if(data["issues"][id] && data['issues'][id]['fields']) {
      // Fetch data into array from json
      var values = getFieldsFromNode(data['issues'][id]);
      // Add update timestamp
      values.push(timestamp);
      // Create info object
      var info = {};
      info.key = data['issues'][id].key;
      info.epic = data['issues'][id].fields.customfield_10007;
      info.parent = typeof data['issues'][id].fields.parent == 'undefined' ? data['issues'][id].fields.parent : data['issues'][id].fields.parent.key;
      info.type = data['issues'][id].fields.issuetype.name;
      // Add issue to array
      issues.push({value: values, info: info});
    }

  // Delete old values
  if (!isNaN(fromIndex))
    sheet.getRange(fromIndex + 2, 1, sheet.getLastRow() - fromIndex - 1, jqlOptions.fields.length + 1).clear();
  /*
  if (displayOptions.sortByParents) {
    var sortedList = sortIssueList(issues);
    issues = sortedList.resultList;
    formatRows(sortedList.formatList);
  } else */
  issues = issues.map(function(x){return x.value;});

  var lastRow = sheet.getLastRow() + 1;
  for (var i = 0; i < jqlOptions.fields.length + 1; i++) {
    var format = "";
    if (i < jqlOptions.fields.length)
      format = jqlOptions.fields[i].split("|")[0];

    var slice = issues.map(function(value) { return [value[i]]; });
    var range = sheet.getRange(lastRow, i + 1, issues.length, 1);

    if (format == 'text')
      range.setNumberFormat('@')

    if (slice[0][0] == '=')
      range.setFormulas(slice);
    else
      range.setValues(slice);
  }
}

function getFieldsFromNode(node){
  var values = [];

  for (var i = 0, len = jqlOptions.fields.length; i < len; i++) {
    if (displayOptions.fields2links.indexOf(jqlOptions.fieldsNames[i]) > -1)
      values.push(jsonPathToValue(node, jqlOptions.fields[i], jqlOptions.fieldsNames[i]));
    else
      values.push(jsonPathToValue(node, jqlOptions.fields[i], void 0));
  }

  return values;
}

function jsonPathToValue(jsonData, path, toLink) {
  if (!(jsonData instanceof Object) || typeof (path) === 'undefined')
    throw 'Not valid argument: jsonData:' + jsonData + ', path:' + path;

  var input = path.split('|');
  var format = '';
  if (input.length > 1) {
    path = input[1];
    format = input[0];
  }

  path = path.replace(/\[(\w+)\]/g, '.$1'); // convert indexes to properties
  path = path.replace(/^\./, ''); // strip a leading dot
  var pathArray = path.split('.');
  for (var i = 0, n = pathArray.length; i < n; ++i) {
    var key = pathArray[i];
    if (!Array.isArray(jsonData)) {
      if (key in jsonData) {
        if (jsonData[key] !== null) {
          jsonData = jsonData[key];
        } else {
          if (format.valueOf() == 'user')
            return 'Unassigned';
          return null;
        }
      } else {
        return "";
      }
    } else {
      var values = jsonData.map(function(x) { return x[key];});
      if (typeof toLink != 'undefined' && values.length > 0)
        return value2link(values, toLink);
      return values.toString();
    }
  }

  var value = formatValue(jsonData, format, path);

  if (typeof toLink != 'undefined')
    return value2link(value, toLink);

  if (Array.isArray(value))
    return value.toString();

  return value;
}

function getLastTimeStamp(msDif) {
  var curDate = new Date();

  var values = sheet.getDataRange().getValues();

  if (sheet.getLastRow() == 0)
    return [NaN, -1];

  var date = values[sheet.getLastRow() - 1][jqlOptions.fields.length];

  // Unparsable
  if (typeof date == 'undefined')
    return [NaN, -1];

  // If not date, don't search
  if (typeof date.getMonth !== 'function')
    return [NaN, -1];

  var dif = dateDiffInMS(date, curDate);

  // Too old
  if (dif > msDif)
    return [NaN, -1];

  var timestampPos = jqlOptions.fields.length;

  // Return first date instance in list
  for (var i = sheet.getLastRow() - 1; i >= 0 ; i--) {
    var value = values[i][timestampPos];
    if (typeof value.getMonth !== 'function' || dateDiffInMS(value, curDate) > msDif)
      return [date, i];
  }

  return [NaN, -1];
}

function getCustomFields() {
  var customFields = {fields:[], fieldsNames:[]};

  var fetchArgs = {
    contentType: 'application/json',
    headers: {'Authorization':'Basic ' +  connectOptions.ennCred},
    muteHttpExceptions: true
  };

  var url = connectOptions.baseURL + 'rest/api/2/field';

  var httpResponse = UrlFetchApp.fetch(url, fetchArgs);
  if (httpResponse) {
    var responseCode = httpResponse.getResponseCode();
    if (responseCode == 200) {
      var data = JSON.parse(httpResponse.getContentText());

      data.map(function(x){
        if (jqlOptions.customFields.indexOf(x.name) > -1) {

          customFields.fields.push(getPathForType(x.schema.type, x.id));
          customFields.fieldsNames.push(x.name);
        }
      });
    }
  }

  return customFields;
}

function getPathForType(type, id) {
  switch(type) {
    case 'option':
      return 'fields.' + id + '.value';
    case 'number':
      return 'fields.' + id;
    case 'array':
      return 'array|fields.' + id;
    case 'string':
      return 'fields.' + id;
    case 'user':
      return 'user|fields.' + id + '.displayName'
  }

  // Case any
  return 'fields.' + id;
}

function formatValue(value, format, key) {
  if (key.startsWith('fields.customfield_')) {
    for (var i = 0; i < jqlOptions.customFields.length; i++){
      if (jqlOptions.customFields[i].value.indexOf(key) > 0) {
        if (jqlOptions.customFields[i].regex != "")
          return value[0].replace(new RegExp(jqlOptions.customFields[i].regex), '$1');
        else break;
      }
    }
  }

  switch(format) {
    case 'duration':
      if (typeof unit == 'undefined') {
        unit = 1;
        switch (displayOptions.estimationUnit) {
          case 'weeks':
            unit *= 7;
          case 'days':
            unit *= 24;
          case 'hours':
            unit *= 60 * 60;
        }
      }
      var result = (value / unit).toFixed(2).replace('.', localOptions.decimalSeparator);
      return result;
    case 'date':
      // Parse date (Date.parse returns inccorect date with server-side js)
      var a = value.split(/[^0-9]/);
      if (a.length > 3)
        return date2str(new Date(Date.UTC(a[0], a[1] - 1, a[2], a[3]-a[6]/100, a[4], a[5])), displayOptions.dateformat);
      else
        return date2str(new Date(Date.UTC(a[0], a[1] - 1, a[2])), displayOptions.dateformat);
    case 'attachment':
      var fetchArgs = {
        contentType: 'application/json',
        headers: {'Authorization':'Basic ' +  connectOptions.ennCred},
        muteHttpExceptions: true
      };

      var url = connectOptions.baseURL + 'rest/api/2/issue/' + value + '?expand=attachment';

      var httpResponse = UrlFetchApp.fetch(url, fetchArgs);
      if (httpResponse) {
        var responseCode = httpResponse.getResponseCode();
        if (responseCode == 200) {
          var data = JSON.parse(httpResponse.getContentText());
          var mimeTypes = jsonPathToValue(data, 'fields.attachment.mimeType');
          var seen = {};
          return mimeTypes.split(/,|\//).filter(function(item) {
            return seen.hasOwnProperty(item) ? false : (seen[item] = true);
          }).join(', ');
        }
      }

    case 'worklog':
      var fetchArgs = {
        contentType: 'application/json',
        headers: {'Authorization':'Basic ' +  connectOptions.ennCred},
        muteHttpExceptions: true
      };

      var url = connectOptions.baseURL + 'rest/api/2/issue/' + value + '/worklog';
      var httpResponse = UrlFetchApp.fetch(url, fetchArgs);
      if (httpResponse) {
        var responseCode = httpResponse.getResponseCode();
        if (responseCode == 200) {
          var data = JSON.parse(httpResponse.getContentText());
          var worklogs = data.worklogs;
          var worklogsList = [];
          for (var i = 0; i < worklogs.length; i++) {
            var a = worklogs[i].created.split(/[^0-9]/);
            var date;
            if (a.length > 3)
              date = date2str(new Date(Date.UTC(a[0], a[1] - 1, a[2], a[3]-a[6]/100, a[4], a[5])), displayOptions.dateformat);
            else
              date = date2str(new Date(Date.UTC(a[0], a[1] - 1, a[2])), displayOptions.dateformat);

            worklogsList.push(worklogs[i].author.displayName + ' | ' + worklogs[i].timeSpent + ' | ' + date);
          }

          return worklogsList.join('\n');
        }
      }

    case 'prLink':
      var fetchArgs = {
        contentType: 'application/json',
        headers: {'Authorization':'Basic ' +  connectOptions.ennCred},
        muteHttpExceptions: true
      };

      var prList = [];
      for (var i=0; i<reposTypes.length; i++) {
        var url = connectOptions.baseURL + 'rest/dev-status/1.0/issue/detail?issueId=' + value + '&applicationType=' + reposTypes[i] + '&dataType=pullrequest';
        var httpResponse = UrlFetchApp.fetch(url, fetchArgs);
        if (httpResponse) {
          var responseCode = httpResponse.getResponseCode();
          if (responseCode == 200) {
            var data = JSON.parse(httpResponse.getContentText());
            if (typeof data.detail[0] == 'undefined') continue;
            data.detail[0].pullRequests.map(function(x) { prList.push(x.status + ' | ' + x.url);});
          }
        }
      }

      return prList.join('\n');
    case 'text':
      return '"' + value + '"';
    case 'textcap':
      return value.substring(0,50000);
    case 'array':
      // Try casting array to values
      if (value.length <= 0)
        return value;

      if (value[0].hasOwnProperty('value'))
        for (var i = 0; i < value.length; i++)
          value[i] = value[i].value;
      else if (value[0].hasOwnProperty('name'))
        for (var i = 0; i < value.length; i++)
          value[i] = value[i].name;

      return value;
    case 'sprint':
      var re = /.*name=(.*)\,startDate.*/;
      return value[0].replace(re, '$1');
    case 'epic':
      // Getting epics one by one seems to be cheaper due to the fact that they returned unordered and would require two value sorting
      var fetchArgs = {
        contentType: 'application/json',
        headers: {'Authorization':'Basic ' +  connectOptions.ennCred},
        muteHttpExceptions: true
      };

      var url = connectOptions.baseURL + 'rest/api/2/search?jql=key=' + value;

      var httpResponse = UrlFetchApp.fetch(url, fetchArgs);
      if (httpResponse) {
        var responseCode = httpResponse.getResponseCode();
        if (responseCode == 200) {
          var data = JSON.parse(httpResponse.getContentText());
          return data.issues[0].fields.summary;
        }
      }
  }

  return value.toString();
}

function value2link(value, field) {
  var link = '';

  if (Array.isArray(value)) {
    if (value.length < 1)
      return value.toString();

    field = fieldName2jqlName(field);
    var displayedName = value.toString();
    link = connectOptions.baseURL + 'issues/?jql=' + field + '=' + encodeURIComponent('"' + value.pop() + '"');
    value.forEach(function(x) { link += ' AND ' + field + '=' + encodeURIComponent('"' + x + '"'); });
    return '=HYPERLINK("' + link + '";"' + displayedName + '")';
  }
  else {
    if (field == 'Key' || field == 'Epic' || field == 'Parent Key')
      link = connectOptions.baseURL + 'browse/' + value;
    else
      link = connectOptions.baseURL + 'issues/?jql=' + fieldName2jqlName(field) + '=' + encodeURIComponent('"' + value + '"');
  }

  return '=HYPERLINK("' + link + '";"' + value + '")';
}

function date2str(x, y) {

  var z = {
    M: x.getMonth() + 1,
    D: x.getDate(),
    h: x.getHours(),
    m: x.getMinutes(),
    s: x.getSeconds()
  };
  y = y.replace(/(M+|D+|h+|m+|s+)/g, function(v) {
    return ((v.length > 1 ? '0' : '') + eval('z.' + v.slice(-1))).slice(-2)
  });

  return y.replace(/(Y+)/g, function(v) {
    return x.getFullYear().toString().slice(-v.length)
  });
}

function updateHeadRow() {
  var fieldsNames = jqlOptions.fieldsNames;
  fieldsNames.push('Timestamp');

  // Remove leftover fields
  if (sheet.getLastColumn() > fieldsNames.length) {
    sheet.getRange(1, fieldsNames.length + 1, 1, sheet.getLastColumn() - fieldsNames.length + 1).clear();
  }

  // Update old fields
  var range = sheet.getRange(1, 1, 1, fieldsNames.length);
  range.setValues([fieldsNames]);
  sheet.setFrozenRows(1);
  range.setBackground('#111');
  range.setFontColor('#eee');
}

function dateDiffInMS(a, b) {
  // Discard the time and time-zone information.
  var utc1 = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate(), a.getHours(), a.getMinutes(), a.getSeconds());
  var utc2 = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate(), b.getHours(), b.getMinutes(), b.getSeconds());

  return utc2 - utc1;
}

function setupFilter() {
  // Create filters
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var filterSettings = {};

  // The range of data on which you want to apply the filter.
  // optional arguments: startRowIndex, startColumnIndex, endRowIndex, endColumnIndex
  filterSettings.range = {
    sheetId: sheet.getSheetId(),
    startRowIndex: 0,
    endRowIndex: sheet.getLastRow(),
    startColumnIndex: 0,
    endColumnIndex: jqlOptions.fields.length + 1
  };

  var request = {
    'setBasicFilter': {
      'filter': filterSettings
    }
  };
  Sheets.Spreadsheets.batchUpdate({'requests': [request]}, ss.getId());
}

function fieldName2jqlName(fieldName) {
  switch (fieldName) {
    case 'Components':
      return 'component';
    case 'Fix Version/s':
      return 'fixVersion';
    case 'Affected Version/s':
      return 'affectedVersion';
  }
  return fieldName;
}

function formatRows(rowsIndexes) {
  rowsIndexes.forEach(function(x){
    var range = sheet.getRange(x + 1, 1, 1, jqlOptions.fields.length + 1);

    range.setBackground('#111');
    range.setFontColor('#eee');
  });
}

function sortIssueList(list) {
  var resultList = [];

  var emptyRow = Array.apply(null, Array(list[0].value.length)).map(String.prototype.valueOf,"")

  // Do a very very slow bucket sort
  epicBuckets = {'undefined': {'bucketList':[], 'bucketRoot':emptyRow}};

  parentBuckets = {};

  list.forEach(function(x) {
    if (x.info.epic != null)
      if (x.info.epic in epicBuckets)
        epicBuckets[x.info.epic].bucketList.push(x);
      else
        epicBuckets[x.info.epic] = {'bucketList':[x], 'bucketRoot':emptyRow};
    else
      if (x.info.parent != null)
        if (x.info.parent in parentBuckets)
          parentBuckets[x.info.parent].push(x.value);
        else
          parentBuckets[x.info.parent] = [x.value];
      else
        epicBuckets['undefined'].bucketList.push(x);
  });

  // Find all possible bucket roots
  epicBuckets['undefined'].bucketList = epicBuckets['undefined'].bucketList.filter(function(x){
    if (x.info.key in epicBuckets) {
      epicBuckets[x.info.key].bucketRoot = x.value.slice();
      return false;
    }
    return true;
  });

  var undefList = [];
  epicBuckets['undefined'].bucketList.forEach(function(x){
    undefList = undefList.concat([x.value]);
    if (x.info.key in parentBuckets)
      undefList = undefList.concat(parentBuckets[x.info.key]);
  })
  // Create a list from buckets
  resultList = resultList.concat(undefList);
  delete epicBuckets['undefined'];

  // Add all fields without parents
  var keys = list.map(function(x){ return x.info.key; });
  Object.keys(parentBuckets).forEach(function (key) {
    if (keys.indexOf(key) < 0)
      resultList = resultList.concat(parentBuckets[key]);
  });

  var formatList = [];

  Object.keys(epicBuckets).forEach(function (key) {
    resultList.push(epicBuckets[key].bucketRoot);
    formatList.push(resultList.length);
    var tempList = [];
    epicBuckets[key].bucketList.forEach(function(x) {
      tempList = tempList.concat([x.value]);
      if (x.info.key in parentBuckets)
        tempList = tempList.concat(parentBuckets[x.info.key]);
    });
    resultList = resultList.concat(tempList);
  });

  return {resultList: resultList, formatList: formatList};
}
