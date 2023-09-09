 * @overview Export keyword performance data along with quality score  
 * parameters to the give spreadsheet.  

// Path to the folder in Google Drive where all the reports are to be created
var REPORTS_FOLDER_PATH = 'Google Ads Keyword Quality Scores';
var EMAIL_ADDRESSES = {{EMAIL ADDRESS}}; //insert your email
var EMAIL_SUBJECT = "Haftalık Keyword Analizi"
var EMAIL_BODY = "Google Ads haftalık raporu hazır. Lütfen EMAIL ADDRESS hesabının Google Drive'ına girerek otomatik oluşturulmuş olan Nebim sheet'inden analizleri yap."; // add any text you would like to add to your email alerts
    MailApp.sendEmail(EMAIL_ADDRESSES, EMAIL_SUBJECT, EMAIL_BODY);
 
// Specify a date range for the report 
var DATE_RANGE = "LAST_90_DAYS"; // Other allowed values are: LAST_<NUM>_DAYS (Ex: LAST_90_DAYS) or TODAY, YESTERDAY, THIS_WEEK_SUN_TODAY, THIS_WEEK_MON_TODAY, LAST_WEEK, LAST_WEEK, LAST_BUSINESS_WEEK, LAST_WEEK_SUN_SAT, THIS_MONTH, LAST_MONTH 

// Or specify a custom date range. Format is: yyyy-mm-dd
var USE_CUSTOM_DATE_RANGE = false;
var START_DATE = "<Date in yyyy-mm-dd format>"; // Example "2016-02-01"
var END_DATE = "<Date in yyyy-mm-dd format>"; // Example "2016-02-29"

// Set this to true to only look at currently active campaigns. 
// Set to false to include campaigns that had impressions but are currently paused.
var IGNORE_PAUSED_CAMPAIGNS = true;
 
// Set this to true to only look at currently active ad groups.
// Set to false to include ad groups that had impressions but are currently paused.  
var IGNORE_PAUSED_ADGROUPS = true;

var IGNORE_PAUSED_KEYWORDS = true;
var REMOVE_ZERO_IMPRESSIONS_KW = true;

// Number of top keywords (by impressions) to export to spreadsheet
var MAX_KEYWORDS = 1000;

/*-- More filter for MCC account --*/
//Is your account a MCC account
var IS_MCC_ACCOUNT = false;

var FILTER_ACCOUNTS_BY_LABEL = false;
var ACCOUNT_LABEL_TO_SELECT = "INSERT_LABEL_NAME_HERE";

var FILTER_ACCOUNTS_BY_IDS = false;
var ACCOUNT_IDS_TO_SELECT = ['INSERT_ACCOUNT_ID_HERE', 'INSERT_ACCOUNT_ID_HERE'];
/*---------------------------------*/

//////////////////////////////////////////////////////////////////////////////
function main() {
  var reportsFolder = getFolder(REPORTS_FOLDER_PATH);
  
  if (!IS_MCC_ACCOUNT) {
    processCurrentAccount(reportsFolder);
  } else {
    var childAccounts  = getManagedAccounts();
    while(childAccounts .hasNext()) {
      var childAccount  = childAccounts .next()
      AdsManagerApp.select(childAccount);
      processCurrentAccount(reportsFolder);
    }
  }
  Logger.log("Done!");
  Logger.log("=========================");
  Logger.log("All the reports are available in the Google Drive folder at following URL: ");
  Logger.log(reportsFolder.getUrl());
  Logger.log("=========================");
}

function getManagedAccounts() {
  var accountSelector = AdsManagerApp.accounts();
  if (FILTER_ACCOUNTS_BY_IDS) {
    accountSelector = accountSelector.withIds(ACCOUNT_IDS_TO_SELECT);
  }
  if (FILTER_ACCOUNTS_BY_LABEL) {
    accountSelector = accountSelector.withCondition("LabelNames CONTAINS '" + ACCOUNT_LABEL_TO_SELECT + "'")
  }
  return accountSelector.get();  
}

function processCurrentAccount(reportsFolder) {
  var adWordsAccount = AdsApp.currentAccount();
  var spreadsheet = getReportSpreadsheet(reportsFolder, adWordsAccount);

  var accountName = adWordsAccount.getName();
  var currencyCode = adWordsAccount.getCurrencyCode();
  Logger.log("Accesing AdWord account: " + accountName);
  Logger.log("Fetching data from AdWords..");
  var keywordReport = getKeywordReport();

  Logger.log("Computing..");
  var keywordArray = compute(keywordReport);
  var summary = computeSummary(keywordArray);

  Logger.log("Exporting results to spreadsheet..");
  var dateString = Utilities.formatDate(new Date(), adWordsAccount.getTimeZone(), "yyyyMMdd");
  // Export keyword level stats to a spreadsheet
  var newSheetName = dateString + "-Details";
  var sheet = spreadsheet.getSheetByName(newSheetName);
  if (sheet != null) {
    clearDataAndCharts(sheet);
  } else {
    sheet = spreadsheet.insertSheet(newSheetName, 0);
  }
  //keywordReport.exportToSheet(sheet);
  exportToSpreadsheet(keywordArray, sheet, accountName);

  Logger.log("Exporting summary & charts to spreadsheet..");
  // Export summary data and charts to another sheet
  var newSummarySheetName = dateString + "-Summary";
  var summarySheet = spreadsheet.getSheetByName(newSummarySheetName);
  if (summarySheet != null) {
    clearDataAndCharts(summarySheet);
  } else {
    summarySheet = spreadsheet.insertSheet(newSummarySheetName, 0);
  }
  exportSummaryStatsToSpreadsheet(summary, summarySheet);
}

function compute(keywordReport) {
  var reportIterator = keywordReport;
  var keywordArray = new Array();
  while (reportIterator.hasNext()) {
    var row = reportIterator.next();
    var kw = new Object();
    kw['CampaignId'] = row.campaign.id;
    kw['AdGroupId'] = row.adGroup.id;
    kw['Id'] = row.adGroupCriterion.criterionId;
    kw['CampaignName'] = row.campaign.name;
    kw['AdGroupName'] = row.adGroup.name;
    kw['Criteria'] = row.adGroupCriterion.keyword.text;
    kw['KeywordMatchType'] = row.adGroupCriterion.keyword.matchType;

    var hasQualityScore = false;
    var qualityScore = -1;
    var searchPredictedCtr = "";
    var creativeQualityScore = "";
    var postClickQualityScore = "";
    if (row.adGroupCriterion.qualityInfo) {
        hasQualityScore = true;
        qualityScore = row.adGroupCriterion.qualityInfo.qualityScore;
        searchPredictedCtr = row.adGroupCriterion.qualityInfo.searchPredictedCtr.toString();
        creativeQualityScore = row.adGroupCriterion.qualityInfo.creativeQualityScore.toString();
        postClickQualityScore = row.adGroupCriterion.qualityInfo.postClickQualityScore.toString();
    }
    kw['QualityScore'] = qualityScore;
    kw['SearchPredictedCtr'] = searchPredictedCtr;
    kw['CreativeQualityScore'] = creativeQualityScore;
    kw['PostClickQualityScore'] = postClickQualityScore;

    kw['Clicks'] = row.metrics.clicks;
    kw['Impressions'] = row.metrics.impressions;
    kw['Ctr'] = row.metrics.ctr;
    kw['AverageCpc'] = row.metrics.averageCpc / 1000000;
    kw['Cost'] = row.metrics.costMicros / 1000000;
    kw['Conversions'] = row.metrics.conversions;
    if (row.metrics.costPerConversion) {
        kw['CostPerConversion'] = row.metrics.costPerConversion / 1000000;
    } else {
        kw['CostPerConversion'] = "--";
    }
    if (hasQualityScore) {
      keywordArray.push(kw);
    }
  }
  Logger.log("Total keywords found: " + keywordArray.length);
  keywordArray.sort(getComparator("Impressions", true));
  
  // Truncate  the array after MAX_KEYWORDS limit
  keywordArray = keywordArray.slice(0, MAX_KEYWORDS);
  return keywordArray;
}

function getKeywordReport() {
  var dateRange = getDateRange(" AND ");
  
  var whereStatements = "";
  if (IGNORE_PAUSED_CAMPAIGNS) {
    whereStatements += " AND campaign.status = 'ENABLED' ";
  } else {
    whereStatements += " AND campaign.status IN ['ENABLED','PAUSED'] ";
  }
  
  if (IGNORE_PAUSED_ADGROUPS) {
    whereStatements += " AND ad_group.status = 'ENABLED' ";
  } else {
    whereStatements += " AND ad_group.status IN ['ENABLED','PAUSED'] ";
  }

  if (IGNORE_PAUSED_KEYWORDS) {
    whereStatements += " AND ad_group_criterion.status = 'ENABLED' "; 
  } else {
    whereStatements += " AND ad_group_criterion.status IN ['ENABLED','PAUSED'] ";
  }
  
  if (REMOVE_ZERO_IMPRESSIONS_KW) {
    whereStatements += " AND metrics.impressions > 0 "; 
  }
  
  var query = "SELECT campaign.id, ad_group.id, ad_group_criterion.criterion_id, campaign.name, ad_group.name, ad_group_criterion.keyword.text, ad_group_criterion.keyword.match_type, ad_group_criterion.quality_info.quality_score, ad_group_criterion.quality_info.search_predicted_ctr, ad_group_criterion.quality_info.creative_quality_score, ad_group_criterion.quality_info.post_click_quality_score,  metrics.clicks, metrics.impressions, metrics.ctr, metrics.average_cpc, metrics.cost_micros, metrics.conversions, metrics.cost_per_conversion " +
    "FROM  keyword_view " +
    "WHERE segments.date BETWEEN " + dateRange + whereStatements;
    Logger.log("Query: " + query);
  return AdsApp.search(query);
}

function getDateRange(seperator) {
  var dateRange = DATE_RANGE;
  if (USE_CUSTOM_DATE_RANGE) {
    dateRange = "'" + START_DATE + "'" + seperator + "'"+ END_DATE + "'";
  } else if (dateRange.match(/LAST_(.*)_DAYS/)) {
    var adWordsAccount = AdWordsApp.currentAccount();
    var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    var numDaysBack = parseInt(dateRange.match(/LAST_(.*)_DAYS/)[1]);
    var today = new Date();
    var endDate = Utilities.formatDate(new Date(today.getTime() - MILLIS_PER_DAY), adWordsAccount.getTimeZone(), "yyyy-MM-dd");// Yesterday
    var startDate = Utilities.formatDate(new Date(today.getTime() - (MILLIS_PER_DAY * numDaysBack)), adWordsAccount.getTimeZone(), "yyyy-MM-dd");
    dateRange = "'" + startDate + "'" + seperator + "'" + endDate + "'";
  }
  return dateRange;
}

function getComparator(sortFieldName, reverse) {
  return function(obj1, obj2) {
    var retVal = 0;
    var val1 = parseInt(obj1[sortFieldName]);
    var val2 = parseInt(obj2[sortFieldName]);
    if (val1 < val2)
      retVal = -1;
    else if (val1 > val2)
      retVal = 1;
    else 
      retVal = 0;
    
    if (reverse) {
      retVal = -1 * retVal;
    }
    return retVal;
  }
}

function exportToSpreadsheet(keywordArray, sheet, accountName) {
  var rowsArray = new Array();
  for (var i=0; i<keywordArray.length; i++) {
    var kw = keywordArray[i];
    if (kw["SearchPredictedCtr"]!="Not applicable") {
      rowsArray.push(
          [ kw["CampaignName"], kw["AdGroupName"], kw["Criteria"], kw["KeywordMatchType"], 
            kw["Clicks"], kw["Impressions"], kw["Ctr"], kw["AverageCpc"], kw["Conversions"], 
            kw["Cost"], kw["CostPerConversion"], kw["QualityScore"], 
            kw["SearchPredictedCtr"].replace(/\s/g, "_").toUpperCase(), 
            kw["CreativeQualityScore"].replace(/\s/g, "_").toUpperCase(), 
            kw["PostClickQualityScore"].replace(/\s/g, "_").toUpperCase()
          ]
      );
    }
  }

  var colTitleColor = "#03cfcc"; // Aqua
  var summaryRowColor = "#D3D3D3"; // Grey
  var headers = ["Campaign Name", "Ad Group Name", "Keyword", "Match Type", "Clicks", "Impressions", "Ctr", "Avg CPC", "Conversions", "Cost", "Cost Per Conversion", "Quality Score", "Expected CTR", "Ad Relevance", "Landing Page Experience"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground(colTitleColor).setFontWeight("BOLD");
  if (rowsArray.length > 0) {
    sheet.getRange(2, 1, rowsArray.length, headers.length).setValues(rowsArray);
    applyColorCoding(sheet, rowsArray, 13, 1);
    applyColorCoding(sheet, rowsArray, 14, 1);
    applyColorCoding(sheet, rowsArray, 15, 1);
    sheet.getRange(1, 5, rowsArray.length).setNumberFormat("#,##0");
    sheet.getRange(1, 6, rowsArray.length).setNumberFormat("#,##0");
    sheet.getRange(1, 7, rowsArray.length).setNumberFormat("#,##0.00%");
    sheet.getRange(1, 8, rowsArray.length).setNumberFormat("#,##0.00");
    sheet.getRange(1, 9, rowsArray.length).setNumberFormat("#,##0.00");
    sheet.getRange(1, 10, rowsArray.length).setNumberFormat("#,##0.00");
    sheet.getRange(1, 11, rowsArray.length).setNumberFormat("#,##0.00");
    sheet.getRange(1, 13, rowsArray.length).setNumberFormat("#,##0");
  }
  for(var i=1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);  
  }
  sheet.setFrozenRows(1);
}

function applyColorCoding(sheet, rowsArray, colIdx, rowOffset) {
  for(var i=0; i<rowsArray.length; i++) {
    var bgColor = getBgColorFor(rowsArray[i][colIdx]);
    sheet.getRange(i+rowOffset+1, colIdx+1).setBackground(bgColor);
  }
}

function getBgColorFor(value) {
  if (value == "BELOW_AVERAGE") return "#FACDB1"; // Red
  if (value == "ABOVE_AVERAGE") return "#CDF6BE"; // Green
  if (value == "AVERAGE") return "#EAEAEA";       // Grey
}

/* ******************************************* */
function computeSummary(keywordArray) {
  var summary = getNewEmptySummaryObject();
  for(var i=0; i < keywordArray.length; i++) {
    addStats(summary, keywordArray[i]);
  }
  setAverages(summary);
  return summary;
}

function addStats(summary, kw) {
  var clicks = cleanAndParseInt(kw["Clicks"]);
  var impressions = cleanAndParseInt(kw["Impressions"]);
  var cost = cleanAndParseFloat(kw["Cost"]);
  var conversions = cleanAndParseFloat(kw["Conversions"]);
  groupByValue(clicks, impressions, conversions, cost, summary.byQS, parseInt(kw["QualityScore"]));
  groupByValue(clicks, impressions, conversions, cost, summary.byLPE, kw["PostClickQualityScore"]);
  groupByValue(clicks, impressions, conversions, cost, summary.byExpectedCTR, kw["SearchPredictedCtr"]);
  groupByValue(clicks, impressions, conversions, cost, summary.byAdRelevance, kw["CreativeQualityScore"]);
}

function groupByValue(clicks, impressions, conversions, cost, map, key) {
  var aggStats = map[key];
  if (!aggStats) {
    aggStats = getNewEmptyStatsObject();
    map[key] = aggStats;
  }
  aggStats["Clicks"] += clicks;
  aggStats["Impressions"] += impressions;
  aggStats["Conversions"] += conversions;
  aggStats["Cost"] += cost;
};

function setAverages(summary) {
  for(groupByName in summary) {
    var map = summary[groupByName]
    for(key in map) {
      var stats = map[key];
      if (stats["Clicks"] > 0) {
        stats["AverageCpc"]= stats["Cost"] / stats["Clicks"];
      }
      if (stats["Conversions"])
      stats["CostPerConversion"] = stats["Cost"] / stats["Conversions"];
    }
  }
}

function getNewEmptyStatsObject() {
  return {
    "Clicks":0,
    "Impressions":0,
    "Conversions":0,
    "Cost":0,
    "AverageCpc":0,
    "CostPerConversion":0
  };
}

function getNewEmptySummaryObject() {
  return {
    "byQS": new Object(),
    "byLPE": new Object(),
    "byExpectedCTR": new Object(),
    "byAdRelevance": new Object()
  };
};

function exportSummaryStatsToSpreadsheet(summary, sheet) {
  var startRow = 1;
  exportAggStats(sheet, summary.byQS, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], startRow, "Quality Score", "#FFFF80", Charts.ChartType.COLUMN);
  
  var qsParamValues = ["Above average", "Average", "Below average"];
  var chartOptions = {"slices":{"0":{"color":"#00ff00"},"1":{"color":"#3366cc"},"2":{"color":"#dc3912"}}};
  exportAggStats(sheet, summary.byLPE, qsParamValues, startRow+= 13, "Lading Page Experience", "#f9cda8", Charts.ChartType.PIE, chartOptions);
  exportAggStats(sheet, summary.byExpectedCTR, qsParamValues, startRow+=6, "Expected CTR", "#a8d4f9", Charts.ChartType.PIE, chartOptions);
  exportAggStats(sheet, summary.byAdRelevance, qsParamValues, startRow+=6, "Ad Relevance", "#9aff9a", Charts.ChartType.PIE, chartOptions);
}

function exportAggStats(sheet, map, keyArray, startRow, keyColHeader, bgColor, chartType, chartOptions) {
  var rowsArray = new Array();
  for (var i=0; i < keyArray.length; i++) {
    var keyDisplayVal = keyArray[i];
    var key = keyDisplayVal;
    if ((typeof keyDisplayVal) == "string")  {
        key = keyDisplayVal.replace(/\s/g, "_").toUpperCase();
    }
    var stats = map[key];
    if (!stats) {
      stats = getNewEmptyStatsObject();
    }
    rowsArray.push(
      [
        keyDisplayVal, stats["Clicks"], stats["Impressions"], stats["Cost"], stats["AverageCpc"],  stats["Conversions"],  stats["CostPerConversion"]
      ]
    );
  }
  
  var headerRow = startRow;
  var headers = [keyColHeader, "Clicks", "Impressions", "Cost", "Avg CPC", "Conversions",  "Cost Per Conversion"];
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]).setBackground(bgColor).setFontWeight("BOLD");

  var firstDataRow = headerRow + 1;
  sheet.getRange(firstDataRow, 1, rowsArray.length, headers.length).setValues(rowsArray).setBackground(bgColor);
//  sheet.getRange(firstDataRow, 1, rowsArray.length).setFontWeight("BOLD");
  sheet.getRange(firstDataRow, 2, rowsArray.length).setNumberFormat("#,##0");
  sheet.getRange(firstDataRow, 3, rowsArray.length).setNumberFormat("#,##0");
  sheet.getRange(firstDataRow, 4, rowsArray.length).setNumberFormat("#,##0.00");
  sheet.getRange(firstDataRow, 5, rowsArray.length).setNumberFormat("#,##0.00");
  sheet.getRange(firstDataRow, 6, rowsArray.length).setNumberFormat("#,##0");
  sheet.getRange(firstDataRow, 7, rowsArray.length).setNumberFormat("#,##0.00");

  for(var i=1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);  
  }
  
  drawChart(sheet, chartType, startRow, rowsArray.length, keyColHeader, chartOptions);
}

function drawChart(sheet, chartType, startRow, numOfRows, colName2, chartOptions) {
  // Creates a column chart for values in range
  var colName1 = "Impressions";
  var range1 = sheet.getRange(startRow+1, 1, numOfRows, 1);
  var range2 = sheet.getRange(startRow+1, 3, numOfRows, 1);

  var chartBuilder = sheet.newChart();
  chartBuilder.addRange(range1).addRange(range2).setChartType(chartType);
  var chartTitle = colName1 + " Vs. " + colName2;
  if (chartType === Charts.ChartType.PIE) {
    chartTitle = colName1 + " Chart for " + colName2;
  }
  chartBuilder.setOption('title', chartTitle);
  chartBuilder.setOption('useFirstColumnAsDomain', true);
  chartBuilder.setOption('legend', {position: 'bottom'});
  chartBuilder.setOption("treatLabelsAsText", true);
  chartBuilder.setOption("hAxis", { "useFormatFromData": true, "title": colName2});
  chartBuilder.setOption("vAxis", { "useFormatFromData": true, "title": colName1});
  if (chartOptions) {
    for(chartOptionName in chartOptions) {
      chartBuilder.setOption(chartOptionName, chartOptions[chartOptionName]);
    }
  }

  chartBuilder.setPosition(startRow, 8, 70-(2*startRow), 2);
  sheet.insertChart(chartBuilder.build());
}

/*
 * Gets the report file (spreadsheet) for the given Adwords account in the given folder.
 * Creates a new spreadsheet if doesn't exist.
 */
function getReportSpreadsheet(folder, adWordsAccount) {
  var accountId = adWordsAccount.getCustomerId();
  var accountName = adWordsAccount.getName();
  var spreadsheet = undefined;
  var files = folder.searchFiles(
      'mimeType = "application/vnd.google-apps.spreadsheet" and title contains "'+ accountId + '"');
  if (files.hasNext()) {
    var file = files.next();
    spreadsheet = SpreadsheetApp.open(file);
  }
  
  if (!spreadsheet) {
    var fileName = accountName + " (" + accountId + ")";
    spreadsheet = SpreadsheetApp.create(fileName);
    var file = DriveApp.getFileById(spreadsheet.getId());
    var oldFolder = file.getParents().next();
    folder.addFile(file);
    oldFolder.removeFile(file);
  }
  return spreadsheet;
}

/*
* Gets the folder in Google Drive for the given folderPath.  
* Creates the folder and all the internediate folders if needed.
*/
function getFolder(folderPath) {
  var folder = DriveApp.getRootFolder();
  var folderNamesArray = folderPath.split("/");
  for(var idx=0; idx < folderNamesArray.length; idx++) {
    var newFolderName = folderNamesArray[idx];
    // Skip if new folder name is empty (possiblly due to slash at the end) 
    if (newFolderName.trim() == "") { 
      continue;
    }
    var folderIterator = folder.getFoldersByName(newFolderName);
    if (folderIterator.hasNext()) {
      folder = folderIterator.next();
    } else {
      Logger.log("Creating folder '" + newFolderName + "'");
      folder = folder.createFolder(newFolderName);
    }
  }
  return folder;
}

function clearDataAndCharts(sheet) {
  sheet.clear();
  var charts = sheet.getCharts();
  for(var i=0; i<charts.length; i++) {
    sheet.removeChart(charts[i]);
  }
}
/* ******************************************* */
function cleanAndParseFloat(valueStr) {
  // valueStr = cleanValueStr(valueStr);
  return parseFloat(valueStr);
}

function cleanAndParseInt(valueStr) {
  // valueStr = cleanValueStr(valueStr);
  return parseInt(valueStr);
}

function cleanValueStr(valueStr) {
  if (valueStr.charAt(valueStr.length - 1) == '%') {
    valueStr = valueStr.substring(0, valueStr.length - 1);
  }
  valueStr = valueStr.replace(/,/g,'');
  return valueStr;
}


