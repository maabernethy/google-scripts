function isoToDate(dateStr){
  var str = dateStr.replace(/-/,'/').replace(/-/,'/').replace(/T/,' ').replace(/\+/,' \+').replace(/Z/,' +00');
  return new Date(str);
}

function fetchRepos(url) {
  if (url == null) {
    return [];
  }

  var newUrl = null,
  response = UrlFetchApp.fetch(url, { method : "GET" }),
  repos = JSON.parse(response.getContentText()),
  linkHeaderArray = response.getHeaders().Link.split(',')[0].split(';');

  if (linkHeaderArray[1].trim() == "rel=\"next\"") {
    newUrl = linkHeaderArray[0].toString().replace('>', '').replace('<', '');
  }

  return repos.concat(fetchRepos(newUrl));
}

function fetchIssues(ownerName, repo) {
  var queryString = "?state=open&sort=updated&direction=desc";
  var url = "https://api.github.com/repos/" + ownerName + "/" + repo.name + "/issues" + queryString;
  var response = UrlFetchApp.fetch(url, { method: "GET" }),
  issuesDesc = JSON.parse(response.getContentText());

  if (response.getAllHeaders().Link) { // has next page
    queryString = "?state=open&sort=updated&direction=asc";
    url = "https://api.github.com/repos/" + ownerName + "/" + repo.name + "/issues" + queryString;
    response = UrlFetchApp.fetch(url, { method : "GET" });
    var issuesAsc = JSON.parse(response.getContentText());

    return [issuesDesc[0], issuesAsc[0]];
  } else {
    return [issuesDesc[0], issuesDesc[issuesDesc.length - 1]];
  }
}

function getLatestCommitDate(ownerName, repo) {
  var url = "https://api.github.com/repos/" + ownerName + "/" + repo.name + "/commits";
  var response = UrlFetchApp.fetch(url, { method : "GET", "muteHttpExceptions":true }),
  commits = JSON.parse(response.getContentText());

  if (commits[0] == undefined) {
    return formattedCommitDate = "-";
  } else {
    return formattedCommitDate = Utilities.formatDate(isoToDate(commits[0].commit.committer.date), "EST", "MM/dd/yyyy");
  }
}

function populateSpreadSheetRow(ownerName, repo, sheet) {
  var formattedOldIssueDate,
  formattedNewIssueDate,
  issuesCount = repo.open_issues_count,
  watchersCount = repo.stargazers_count,
  formattedCommitDate = getLatestCommitDate(ownerName, repo);

  if (issuesCount > 0) {
    var issues = fetchIssues(ownerName, repo);
    formattedOldIssueDate = Utilities.formatDate(isoToDate(issues[1].updated_at), "EST", "MM/dd/yyyy");
    formattedNewIssueDate = Utilities.formatDate(isoToDate(issues[0].updated_at), "EST", "MM/dd/yyyy");    
  } else {
    formattedOldIssueDate = formattedNewIssueDate = "-";
  }

  var quot = "\"";
  var title = "=hyperlink(\""  + repo.html_url + quot + ";" + quot + repo.name + quot + ")";
  sheet.appendRow([title, formattedCommitDate, watchersCount, issuesCount, formattedOldIssueDate, formattedNewIssueDate ])
}

function formatSheet(sheet, sheetRange, repoCount) {
  //resize columns
  for (var i=1; i <= 6; i++) {
    sheet.autoResizeColumn(i);
  }

  // center text in each column
  sheetRange.setHorizontalAlignment("center");

  // sort repos by name (alphabetized)
  sheetRange.sort(1);

  //for each cell in latest commit column - background red if < 2014
  var thresholdDate = new Date(2014, 1, 31);
  for (var i = 2; i <= repoCount + 1; i++) {
    var date = sheet.getRange(i, 2).getValue();

    if (date.valueOf() < thresholdDate.valueOf()) {
      sheet.getRange(i, 2).setBackground('red');
    }
  }
}

function runScript() {
  var ownerName = "ORG_NAME";
  var url = "https://api.github.com/orgs/" + ownerName + "/repos";
  var repos = fetchRepos(url),
  repoCount = repos.length,
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Output"),
  headerNames = ["Project Name", "Lastest Commit", "# Watchers", "# Open Issues/PRs","Least Recently Updated Issue", "Most Recently Updated Issue"],
  headerRange = sheet.getRange(1, 1, 1, headerNames.length),
  sheetRange = sheet.getRange(2, 1, repoCount + 1, headerNames.length);

  headerRange.setValues([headerNames]).setFontWeight("bold").setHorizontalAlignment("center");
  sheetRange.clear();

  for (var i = 0; i < repos.length; i++) {
    populateSpreadSheetRow(ownerName, repos[i], sheet);
  }

  formatSheet(sheet, sheetRange, repoCount);
}
