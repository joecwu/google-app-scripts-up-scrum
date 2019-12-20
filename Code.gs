function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var refresh_entries = [{
    name : "Refresh Member Related Data",
    functionName : "refreshLastUpdate"
  },{
    name : "Refresh Jira Tasks",
    functionName : "refreshJiraTasks"
  },{
    name : "Refresh Planning Tasks",
    functionName : "refreshPlanningTasks"
  }
//TODO: bug, reference cell become #REFERR
//  ,{
//    name : "Refresh Current Cell",
//    functionName : "refreshCurrentCellFormula"
//  }
  ];
  var gen_entries = [{
    name : "Gen Planning Panels",
    functionName : "genPlanningPanelsBySelection"
  },{
    name : "Gen Static Planning Panels",
    functionName : "genStaticPlanningPanelsBySelection"
  },{
    name : "Gen Planning Panels From Jira",
    functionName : "genPlanningPanelsFromJiraCurrentSprint"
  },{
    name : "Gen Static Planning Panels From Jira",
    functionName : "genStaticPlanningPanelsFromJiraCurrentSprint"
  }];
  sheet.addMenu("Refresh", refresh_entries);
  sheet.addMenu("Generate", gen_entries);
};

function refreshCurrentCellFormula() {
  var curr_cell = SpreadsheetApp.getActiveRange();
  var formula = curr_cell.getFormula();
  // clear only the contents, not notes or comments or formatting.
  curr_cell.clearContent();
  curr_cell.setFormula(formula);
}

function refreshLastUpdate() {
  var activeSheet = SpreadsheetApp.getActive();
  var settingsSheet = activeSheet.getSheetByName("Settings");
  settingsSheet.getRange('B4').setValue(new Date().toTimeString());
}

function refreshJiraTasks() {
  var activeSheet = SpreadsheetApp.getActive();
  var settingsSheet = activeSheet.getSheetByName("Settings");
  settingsSheet.getRange('B8').setValue(new Date().toTimeString());
}

function refreshPlanningTasks() {
  var activeSheet = SpreadsheetApp.getActive();
  var settingsSheet = activeSheet.getSheetByName("Settings");
  settingsSheet.getRange('B8').setValue(new Date().toTimeString());
}

function getStoryIdsFromSelection() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selection = activeSheet.getSelection();
  Logger.log('Current Cell: ' + selection.getCurrentCell().getA1Notation());
  Logger.log('Active Range: ' + selection.getActiveRange().getA1Notation());
  var ranges =  selection.getActiveRangeList().getRanges();
  var range = ranges[0];
  var values = range.getValues().map( function(x){ return x[0];} );
  return values;
}

function getStoryIdsFromJiraCurrentSprint() {
  var activeSheet = SpreadsheetApp.getActive();
  var settingsSheet = activeSheet.getSheetByName("Settings");
  var sprintName = settingsSheet.getSheetValues(5, 2, 1, 1);
  var jql = "Sprint = \"UtaPass Sprint " + sprintName + "\" AND issuetype = Story ORDER BY Rank";
  var searchResult = SEARCH_JIRA_TASKS(jql, "", "key")
  var storyIds = searchResult.map(function(r){ return r[0]; });
  return storyIds;
}


function genStaticPlanningPanelsFromJiraCurrentSprint() {
  var storyIds = getStoryIdsFromJiraCurrentSprint();
  genPlanningPanels(storyIds, true);
}

function genStaticPlanningPanelsBySelection() {
  var storyIds = getStoryIdsFromSelection();
  genPlanningPanels(storyIds, true);
}

function genPlanningPanelsFromJiraCurrentSprint() {
  var storyIds = getStoryIdsFromJiraCurrentSprint();
  genPlanningPanels(storyIds, false);
}

function genPlanningPanelsBySelection() {
  var storyIds = getStoryIdsFromSelection();
  genPlanningPanels(storyIds, false);
}

function genPlanningPanels(storyIds, isStaticContent) {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  
  //Template Range B1:K5
  var templateRange = activeSheet.getRange("B1:K5"); //getRange(1,2,5,11);
  var startPosition = activeSheet.getSelection().getActiveRangeList().getRanges()[0].getRowIndex();
  var shiftRowSize = 0;
  
  // Loop for each story
  for (var i = 1; i <= storyIds.length; i++) {
    var rowIndex = startPosition + shiftRowSize;
    var storyId = storyIds[i-1];
    Logger.log('Story Id: ' + storyId);
    
    // retrieve panel value
    var panelRows = GEN_JIRA_STORY_PANEL(storyId, "", false);
    
    var taskCount = Math.max(panelRows.length, 5);
    
    activeSheet.insertRowsAfter(rowIndex, taskCount);
    var copyToRange = activeSheet.getRange(rowIndex, 2, 5, 11);
    templateRange.copyTo(copyToRange);
    if(taskCount > 5) {
      //補其他row的公式
      for(var j = 1; j <= taskCount - 5; j++){
        activeSheet.getRange(3,2,1,11).copyTo(activeSheet.getRange(rowIndex+5+j-1,2,1,11));
      }
    }
    var groupRange = activeSheet.getRange(rowIndex + 1, 2, taskCount-1, 11);
    
    // replace formula by existing value
    //var jiraFormula = activeSheet.getRange(rowIndex,4).getFormula().toString().replace("UTAPASS-",storyId);
    //Logger.log('Formula: ' + activeSheet.getRange(rowIndex,4).getFormula());
    //activeSheet.getRange(rowIndex,4).setFormula(jiraFormula);
    
    // setup formula
    if(isStaticContent) {
      activeSheet.getRange(rowIndex, 4, panelRows.length, panelRows[0].length).setValues(panelRows);
    } else {
      activeSheet.getRange(rowIndex,4).setFormula("=GEN_JIRA_STORY_PANEL(\""+storyId+"\", Settings!$B$8)");
    }
    
    groupRange.shiftRowGroupDepth(1);
    shiftRowSize = shiftRowSize + taskCount + 1;
  }
}

/**
* Imports JSON data to your spreadsheet Ex: IMPORTJSON("http://myapisite.com","city/population")
* @param url URL of your JSON data as string
* @param xpath simplified xpath as string
* @customfunction
*/
function IMPORTJSON(url,xpath){
  try{
    var activeSheet = SpreadsheetApp.getActive();
    var settingsSheet = activeSheet.getSheetByName("Settings");
    var account = settingsSheet.getSheetValues(1, 2, 1, 1);
    var api_token = settingsSheet.getSheetValues(2, 2, 1, 1);
    Logger.log("account:["+account+"]");
    Logger.log("api_token:["+api_token+"]");
    
    var headers = {
        "Authorization" : "Basic " + Utilities.base64Encode(account + ':' + api_token)
      };
    
    var params = {
      "method":"GET",
      "headers":headers
    };
    
    var res = UrlFetchApp.fetch(url, params);
    var content = res.getContentText();
    var json = JSON.parse(content);
    
    var patharray = xpath.split("/");
    //Logger.log(patharray);
    
    for(var i=0;i<patharray.length;i++){
      json = json[patharray[i]];
    }
    
    //Logger.log(typeof(json));
    if(typeof(json) === "undefined"){
      return "Node Not Available";
    } else if(typeof(json) === "object"){
      var tempArr = [];
      
      for(var obj in json){
        tempArr.push([obj,json[obj]]);
      }
      if(tempArr.length == 0){
        return "";
      }
      return tempArr;
    } else if(typeof(json) !== "object") {
      return json;
    }
  }
  catch(err){
    Logger.log(err);
    throw ("Error getting data:" + err);
  }
}

/**
* get jira subtasks Ex: GET_JIRA_SUBTASKS("UTAPASS-0034", Settings!$B$8)
* 回傳會包含 sub tasks 的 id + summary + status
* 請確保下方有足夠的列數，右方也要有空白的二欄。
* Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Member Related Data"
* @param ParentTaskId parent taskid as string
* @customfunction
*/
function GET_JIRA_SUBTASKS(task_id, refresh_key){
  var JIRA_API_URL = "https://kkvideo.atlassian.net/rest/api/3/issue/" + task_id;
  var JIRA_FIELD_STORY_POINTS = "customfield_10005";
  
  try{
    var json = fetchJiraJson(JIRA_API_URL);
    
    subtasks = json["fields"]["subtasks"];
    var subtaskkeys = [];
    for (var i = 0; i < subtasks.length; i++) {
      var data = subtasks[i]["key"];
      var summary = subtasks[i]["fields"]["summary"];
      var status = subtasks[i]["fields"]["status"]["name"];
      var row = [];
      row.push(data);
      row.push(summary);
      row.push(status);
      subtaskkeys.push(row);
    }
    if(subtaskkeys.length == 0) {
      return "";
    }
    return subtaskkeys;
  }
  catch(err){
    Logger.log(err);
    throw ("Error getting data:" + err);
  }
  
}


/**
* get jira sub tasks & linked tasks Ex: GET_JIRA_RELATED_TASKS("UTAPASS-0034", true, false, false, Settings!$B$8)
* 回傳會包含 sub & linked tasks 的 id + summary + status
* 請確保下方有足夠的列數，右方也要有空白的二欄。
* Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Member Related Data"
* @param task_id parent taskid as string
* @param is_only_this_sprint TRUE or FALSE
* @param is_include_self TRUE or FALSE
* @param is_include_linked_issues TRUE or FALSE
* @param refresh_key Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Member Related Data"
* @param fields_string fields definition separated by comma. e.g: "key,fields/summary,fields/status/name
* @customfunction
*/
function GET_JIRA_RELATED_TASKS(task_id, is_only_this_sprint, is_include_self, is_include_linked_issues, refresh_key, fields_string){
  
  try{
    var jql = "parent in (" + task_id + ")";
    if(is_include_linked_issues) {
      jql = jql + " OR issue in linkedissues(" + task_id + ")";
    }
    if(is_include_self) {
      jql = "issuekey in (" + task_id + ") OR " + jql;
    }
    if(is_only_this_sprint) {
      jql = jql + " AND (Sprint = \"UTAPASS Sprint " + sprint + "\")";
    }
    //jql = jql + " ORDER BY issuetype ASC, priority DESC, createdDate DESC";
    jql = jql + " ORDER BY key ASC";
    
    var JIRA_API_URL = "https://kkvideo.atlassian.net/rest/api/3/search?jql=" + encodeURIComponent(jql);
    
    var json = fetchJiraJson(JIRA_API_URL);
  
    if(fields_string == null ) {
      fields = ["key","fields/summary","fields/status/name"];
    } else {
      fields = fields_string.split(",");
    }
    
    issues = json["issues"];
    var subtaskkeys = [];
    for (var i = 0; i < issues.length; i++) {
      var row = [];
      for (var j = 0; j < fields.length; j++) {
        var value = extractJsonProperty(issues[i],fields[j]);
        row.push(value);
      }
      subtaskkeys.push(row);
    }
    if(subtaskkeys.length == 0) {
      return "";
    }
    return subtaskkeys;
  }
  catch(err){
    Logger.log(err);
    throw ("Error getting data:" + err);
  }
  
}

/**
* get jira story panel Ex: GEN_JIRA_STORY_PANEL("UTAPASS-0034", Settings!$B$8)
* 回傳會包含 story id, summary, 並在底下列出所有 sub tasks 的 id + summary + status
* 請確保下方有足夠的列數。
* Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Member Related Data"
* @param task_id parent taskid as string
* @param refresh_key Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Member Related Data"
* @param include_status 回傳值是否包含status
* @customfunction
*/
function GEN_JIRA_STORY_PANEL(story_id, refresh_key, include_status){
  var fields_string = "key,fields/summary,fields/assignee/name,fields/customfield_10005";
  if(include_status) {
    fields_string = "key,fields/summary,fields/assignee/name,fields/customfield_10005,fields/status/name";
  }
  var tasks = GET_JIRA_RELATED_TASKS(story_id, false, true, false, refresh_key, fields_string);
  var total_story_points = 0;
  var story_summary = "";
  var story_status = "";
  var resp = [];

  var task_rows = [];
  for (var i = 0; i < tasks.length; i++) {
    var task_row = [];
    if(tasks[i][0]!=story_id){
      for (var j = 0; j < tasks[i].length; j++) {
        task_row.push(tasks[i][j]);
        if(j==3){
          // story points
          total_story_points = total_story_points + Number(tasks[i][j]);
        }
      }
      task_rows.push(task_row);
    }else{
      story_summary = tasks[i][1];
      story_status = tasks[i][4];
    }
  }
  
  var header_row = [];
  header_row.push(story_id);
  header_row.push(story_summary);
  header_row.push(null);
  header_row.push(total_story_points);
  if(include_status) {
    header_row.push(story_status);
  }

  var title_row = [];
  title_row.push("Task");
  title_row.push("Summary");
  title_row.push("Owner");
  title_row.push("Points");
  if(include_status) {
    title_row.push("Status");
  }
  
  resp.push(header_row);
  resp.push(title_row);
  return resp.concat(task_rows);
}

/**
* get jira task info Ex: GET_JIRA_TASK_INFO("UTAPASS-0034", Settings!$B$8)
* 回傳會包含 sub tasks 的 name + status
* 請確保右方要有二欄空白的欄。
* Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Jira Tasks"
* @param TaskId taskid as string
* @param refresh_key refresh key, Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Jira Tasks"
* @param fields_string fields definition separated by comma. e.g: "key,fields/summary,fields/status/name
* @customfunction
*/
function GET_JIRA_TASK_INFO(task_id, refresh_key, fields_string){
  var JIRA_API_URL = "https://kkvideo.atlassian.net/rest/api/3/issue/" + task_id;
  var JIRA_FIELD_STORY_POINTS = "customfield_10005";
  
  if(fields_string == null ) {
    fields = ["fields/summary","fields/status/name"];
  } else {
    fields = fields_string.split(",");
  }
  
  try{
    var json = fetchJiraJson(JIRA_API_URL);
    
    var row = [];
    for (var j = 0; j < fields.length; j++) {
      var value = extractJsonProperty(json,fields[j]);
      row.push(value);
    }
    
    var result = [];
    result.push(row);
    return result;
  }
  catch(err){
    Logger.log(err);
    throw ("Error getting data:" + err);
  } 
}


/**
* sum by each personal board Ex: SUM_BY_EACH_PERSON_BOARD("J7:J10", Settings!$B$4)
* 回傳會包含每個scrum team member的personal board.
* Settings!$B$4 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Member Related Data"
* @param range sum range.
* @customfunction
*/
function SUM_ALL_SCRUM_MEMBERS(range, refresh_key){
  var activeSheet = SpreadsheetApp.getActive();
  var result = 0;
  var members = GET_SCRUM_TEAM_MEMBERS();
  for(var i = 0; i < members.length; i++){
    var memberName = members[i];
    var memberSheet = activeSheet.getSheetByName(memberName);
    var dataRange = memberSheet.getRange(range);
    var dataValue = dataRange.getValues();
    for(var j = 0; j < dataValue.length; j++){
      result += Number(dataValue[j]);
    }
  }
  return result;
  SpreadsheetApp.flush();
}

/**
* get scrum team members EX: GET_SCRUM_TEAM_MEMBERS(Settings!$B$4)
* 回傳會包含每個scrum team member的name (定義在settings裡的成員名單，有幾位就有幾個row)
* Settings!$B$4 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"refresh"
* @param names.
* @customfunction
*/
function GET_SCRUM_TEAM_MEMBERS(refresh_key){
  var activeSheet = SpreadsheetApp.getActive();
  var settingsSheet = activeSheet.getSheetByName("Settings");
  
 
  //Select the column we will check for the first blank cell
  var columnToCheck = settingsSheet.getRange("E:E").getValues();
  
  // Get the last row based on the data range of a single column.
  var lastRow = getLastRowSpecial(columnToCheck);
  Logger.log(lastRow);
  
  //EXAMPLE: Get the data range based on our selected columns range.
  var dataRange = settingsSheet.getRange(2,5, lastRow-1);
  var dataValues = dataRange.getValues();
  Logger.log(dataValues);
  return dataValues;
  SpreadsheetApp.flush();
}

/************************************************************************
 *
 * Gets the last row number based on a selected column range values
 *
 * @param {array} range : takes a 2d array of a single column's values
 *
 * @returns {number} : the last row number with a value. 
 *
 */ 
 
function getLastRowSpecial(range){
  var rowNum = 0;
  var blank = false;
  for(var row = 0; row < range.length; row++){
 
    if(range[row][0] === "" && !blank){
      rowNum = row;
      blank = true;
    }else if(range[row][0] !== ""){
      blank = false;
    };
  };
  return rowNum;
};


/**
* search jira tasks Ex: SEARCH_JIRA_TASKS("Sprint = "UTAPASS Sprint 113"", Settings!$B$8)
* 回傳會包含 tasks 的 id + summary + status
* 請確保下方有足夠的列數，右方也要有空白的二欄。
* Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Jira Tasks"
* @param jql Jira Query Language
* @param refresh_key refresh key, Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Jira Tasks"
* @param fields_string fields definition separated by comma. e.g: "key,fields/summary,fields/status/name"
* @customfunction
*/
function SEARCH_JIRA_TASKS(jql, refresh_key, fields_string){
  if(fields_string == null ) {
    fields = ["key","fields/summary","fields/status/name"];
  } else {
    fields = fields_string.split(",");
  }
  if(jql.length == 0) {
    return "";
  }
  var JIRA_API_URL = "https://kkvideo.atlassian.net/rest/api/3/search?jql=" + encodeURIComponent(jql);
  
  try{
    var json = fetchJiraJson(JIRA_API_URL);
    
    var issues = json["issues"];
    var return_issues = [];
    for (var i = 0; i < issues.length; i++) {
      var row = [];
      for (var j = 0; j < fields.length; j++) {
        var value = extractJsonProperty(issues[i],fields[j]);
        row.push(value);
      }
      return_issues.push(row);
    }
    return return_issues;
  }
  catch(err){
    Logger.log(err);
    throw ("Error getting data:" + err);
  }
}


/**
* search jira tasks Ex: GET_JIRA_PROPERTY("UTAPASS-1234", Settings!$B$8, "")
* 回傳會包含 tasks 的 id + summary + status
* 請確保下方有足夠的列數，右方也要有空白的二欄。
* Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Jira Tasks"
* @param jql Jira Query Language
* @param refresh_key refresh key, Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Jira Tasks"
* @param fields_string fields definition separated by comma. e.g: "key,fields/summary,fields/status/name"
* @customfunction
*/
function GET_JIRA_PROPERTY(task_id, refresh_key, field_string){
  var JIRA_API_URL = "https://kkvideo.atlassian.net/rest/api/3/issue/" + task_id;
  try{
    var json = fetchJiraJson(JIRA_API_URL);
    
    var value = extractJsonProperty(json, field_string)
    
    var result = [];
    result.push(value);
    return result;
  }
  catch(err){
    Logger.log(err);
    throw ("Error getting data:" + err);
  } 
}

function extractJsonProperty(json, field_string, array_spliter){
  if(array_spliter == null) {
    array_spliter = "\n";
  }
  var fieldpatharray = field_string.split("/");
  var fieldvalue = json;
  for(var k=0;k<fieldpatharray.length;k++){
    if(fieldvalue == null) {
      fieldvalue = "";
      break;
    }
    fieldvalue = fieldvalue[fieldpatharray[k]];
    
    // 特別處理 "Sprint" 的欄位
    if(Array.isArray(fieldvalue) && fieldpatharray[k] == "customfield_10007") {
      fieldvalue = fieldvalue.map(function(v) { 
        if(typeof v === 'string'){
          var arr = v.match(/name=([^,]+),/);
          if(arr != null && arr.length > 0){
            return arr[1];
          }else{
            return v;
          }
        }else if(typeof v === 'object'){
          if(v.name != null) {
            return v.name;
          }else{
            return v;
          }
        }
      });
    }
    
  }
  if(Array.isArray(fieldvalue)){
    fieldvalue = fieldvalue.join(array_spliter);
  }
  return fieldvalue;
}

function fetchJiraJson(url) {
  var activeSheet = SpreadsheetApp.getActive();
  var settingsSheet = activeSheet.getSheetByName("Settings");
  var account = settingsSheet.getSheetValues(1, 2, 1, 1);
  var api_token = settingsSheet.getSheetValues(2, 2, 1, 1);
  Logger.log("account:["+account+"]");
  Logger.log("api_token:["+api_token+"]");
  
  var headers = {
    "Authorization" : "Basic " + Utilities.base64Encode(account + ':' + api_token)
  };
  
  var params = {
    "method":"GET",
    "headers":headers
  };
  
  var res = UrlFetchApp.fetch(url, params);
  var content = res.getContentText();
  var json = JSON.parse(content);
  return json;
}
