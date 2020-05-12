let activeSheet;
let settingsSheet;
let jiraAccount;
let jiraToken;
let jiraHelper;
let sprintNumber;
let sprintFullName;
  
function onOpen() {
  activeSheet = SpreadsheetApp.getActive();
  settingsSheet = activeSheet.getSheetByName("Settings");
  jiraAccount = settingsSheet.getSheetValues(1, 2, 1, 1);
  jiraToken = settingsSheet.getSheetValues(2, 2, 1, 1);
  jiraHelper = new JiraHelper(jiraAccount, jiraToken);
  sprintNumber = settingsSheet.getSheetValues(5, 2, 1, 1);
  sprintFullName = "UtaPass Sprint " + sprintNumber;
}

function getStoryIdsFromSelection() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selection = activeSheet.getSelection();
  console.log('Current Cell: ' + selection.getCurrentCell().getA1Notation());
  console.log('Active Range: ' + selection.getActiveRange().getA1Notation());
  var ranges =  selection.getActiveRangeList().getRanges();
  var range = ranges[0];
  var values = range.getValues().map( function(x){ return x[0];} );
  return values;
}

function getStoryIdsFromJiraCurrentSprint() {
  var jql = "Sprint = \""+ sprintFullName + "\" AND issuetype = Story ORDER BY Rank";
  return jiraHelper.searchJiraIssuesIds(jql);
}

/**
* get jira sub tasks & linked tasks Ex: GET_JIRA_RELATED_TASKS("UTAPASS-0034", true, false, false, Settings!$B$8)
* 回傳會包含 sub & linked tasks 的 id + summary + status
* 請確保下方有足夠的列數，右方也要有空白的二欄。
* Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Member Related Data"
* @param task_id parent taskid as string
* @param is_include_self TRUE or FALSE
* @param is_include_linked_issues TRUE or FALSE
* @param other_query_str 
* @param refresh_key Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Member Related Data"
* @param fields_string fields definition separated by comma. e.g: "key,fields/summary,fields/status/name
* @customfunction
*/
function GET_JIRA_RELATED_TASKS(task_id, is_only_this_sprint, is_include_self, is_include_linked_issues, refresh_key, fields_string){
  let targetSprint = null;
  if(is_only_this_sprint){
    targetSprint = sprintFullName;
  }
  return jiraHelper.getJiraRelatedIssues(task_id, targetSprint, is_include_self, is_include_linked_issues, fields_string);
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
  return jiraHelper.genJiraStoryPanel(story_id, sprintFullName, include_status);
}

/**
* search jira tasks Ex: SEARCH_JIRA_ISSUES("Sprint = "UTAPASS Sprint 113"", Settings!$B$8)
* 回傳會包含 tasks 的 id + summary + status
* 請確保下方有足夠的列數，右方也要有空白的二欄。
* Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Jira Tasks"
* @param jql Jira Query Language
* @param refresh_key refresh key, Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Jira Tasks"
* @param fields_string fields definition separated by comma. e.g: "key,fields/summary,fields/status/name"
* @customfunction
*/
function SEARCH_JIRA_ISSUES(jql, refresh_key, fields_string){
  return jiraHelper.searchJiraIssues(jql, fields_string);
}

/**
* search jira tasks Ex: GET_JIRA_ISSUE("UTAPASS-1234", Settings!$B$8, "")
* 回傳會包含 tasks 的 id + summary + status
* 請確保下方有足夠的列數，右方也要有空白的二欄。
* Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Jira Tasks"
* @param jql Jira Query Language
* @param refresh_key refresh key, Settings!$B$8 是固定值，為了讓來源資料更新時，能強迫連動，但要手動按menu上的"Refresh Jira Tasks"
* @param fields_string fields definition separated by comma. e.g: "key,fields/summary,fields/status/name"
* @customfunction
*/
function GET_JIRA_ISSUE(task_id, refresh_key, fields_string){
  const JIRA_FIELD_STORY_POINTS = "customfield_10005";
  if(fields_string == null ) {
    fields_string = "fields/summary,fields/status/name," + JIRA_FIELD_STORY_POINTS;
  }
  return jiraHelper.getJiraIssue(task_id, fields_string);
}

/**
* QUERY_JIRA_USERS by Jira /rest/api/2/user/search/query API
* @param uql Jira User Query Language. example: is assignee of UTAPASS
* @param fields_string default:"accountId,displayName"
*/
function QUERY_JIRA_USERS(uql, fields_string) {
  return jiraHelper.searchJiraUsers(uql, fields_string);
}