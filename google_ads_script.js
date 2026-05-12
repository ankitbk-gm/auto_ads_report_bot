/**
 * Google Ads Script — Ad Group Daily Metrics → Google Sheet
 *
 * Sheet :   / 'YOUR_GOOGLE_SHEET_ID_HERE'; Tab: Google_Ads
 * Logic : Initial run (header-only sheet) → backfill Apr 1 2026 → yesterday
 *         Regular run → yesterday only
 * Upsert key : Date + Campaign + Ad_Group
 */

var SHEET_ID           = 'YOUR_GOOGLE_SHEET_ID_HERE';
var TAB_NAME           = 'Google_Ads';
var INITIAL_START_DATE = '2026-04-01';

var HEADERS = [
  'Date', 'Last_Updated', 'Campaign', 'Ad_Group',
  'Spend_INR', 'Impressions', 'Unique_Users', 'Clicks',
  'Conversions', 'CPC', 'CTR', 'Conv_Rate', 'CPM'
];

function main() {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(TAB_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(TAB_NAME);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }

    var lastRow      = sheet.getLastRow();
    var isInitialRun = lastRow <= 1;

    var tz        = AdsApp.currentAccount().getTimeZone();
    var today     = new Date();
    var yesterday = new Date(today);
    yesterday.setDate(today.getDate() - 1);

    var datesToProcess = [];

    if (isInitialRun) {
      var cur = parseDateStr(INITIAL_START_DATE);
      while (dateToStr(cur) <= dateToStr(yesterday)) {
        datesToProcess.push(dateToStr(cur));
        cur.setDate(cur.getDate() + 1);
      }
    } else {
      datesToProcess.push(dateToStr(yesterday));
    }

    if (datesToProcess.length === 0) {
      Logger.log('SUCCESS: No dates to process.');
      return;
    }

    // Build eligible ad groups map — MTD spend > 0
    var eligibleMap = buildEligibleAdGroupsMap();

    var existingRows  = [];
    var existingIndex = {};

    if (lastRow > 1) {
      existingRows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
      for (var r = 0; r < existingRows.length; r++) {
        var k = upsertKey(existingRows[r][0], existingRows[r][2], existingRows[r][3]);
        existingIndex[k] = r;
      }
    }

    var nowStr       = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    var upsertCount  = 0;
    var skippedCount = 0;

    for (var d = 0; d < datesToProcess.length; d++) {
      var date  = datesToProcess[d];
      var query = buildDailyQuery(date);
      var rpt   = AdsApp.report(query);
      var rows  = rpt.rows();

      while (rows.hasNext()) {
        var row      = rows.next();
        var campaign = row['campaign.name'];
        var adGroup  = row['ad_group.name'];
        var lKey     = campaign + '||' + adGroup;

        if (!eligibleMap.hasOwnProperty(lKey)) {
          skippedCount++;
          continue;
        }

        var costMicros  = parseFloat(row['metrics.cost_micros'])  || 0;
        var impressions = parseInt(row['metrics.impressions'])    || 0;
        var clicks      = parseInt(row['metrics.clicks'])         || 0;
        var conversions = parseFloat(row['metrics.conversions'])  || 0;
        var avgCpcMicro = parseFloat(row['metrics.average_cpc'])  || 0;
        var ctrRaw      = parseFloat(row['metrics.ctr'])          || 0;
        var avgCpmMicro = parseFloat(row['metrics.average_cpm'])  || 0;

        var spendINR = costMicros  / 1e6;
        var cpc      = avgCpcMicro / 1e6;
        var cpm      = avgCpmMicro / 1e6;
        var ctr      = ctrRaw * 100;
        var convRate = clicks > 0 ? (conversions / clicks * 100) : 0;

        var newRow = [
          date, nowStr, campaign, adGroup,
          spendINR, impressions, '', clicks,
          conversions, cpc, ctr, convRate, cpm
        ];

        var uk = upsertKey(date, campaign, adGroup);

        if (existingIndex.hasOwnProperty(uk)) {
          existingRows[existingIndex[uk]] = newRow;
        } else {
          existingIndex[uk] = existingRows.length;
          existingRows.push(newRow);
        }

        upsertCount++;
      }
    }

    if (existingRows.length > 0) {
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
      }
      sheet.getRange(2, 1, existingRows.length, HEADERS.length).setValues(existingRows);
    }

    Logger.log(
      'SUCCESS: Processed ' + datesToProcess.length + ' date(s). ' +
      'Upserted ' + upsertCount + ' rows. ' +
      'Skipped ' + skippedCount + ' rows. ' +
      'Total sheet rows: ' + existingRows.length + '.'
    );

  } catch (e) {
    Logger.log('ERROR: ' + e.message + (e.stack ? '\n' + e.stack : ''));
  }
}

function buildEligibleAdGroupsMap() {
  var map = {};

  // Get first day of current month
  var today     = new Date();
  var firstOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  var startDate = dateToStr(firstOfMonth);
  var endDate   = dateToStr(new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1));

  // Query MTD spend — only include ad groups with cost_micros > 0
  var query = [
    'SELECT campaign.name, ad_group.name, metrics.cost_micros',
    'FROM ad_group',
    "WHERE segments.date BETWEEN '" + startDate + "' AND '" + endDate + "'",
    'AND metrics.cost_micros > 0'
  ].join(' ');

  var rpt  = AdsApp.report(query);
  var rows = rpt.rows();

  while (rows.hasNext()) {
    var row  = rows.next();
    var k    = row['campaign.name'] + '||' + row['ad_group.name'];
    map[k]   = true;
  }

  return map;
}

function buildDailyQuery(dateStr) {
  return [
    'SELECT',
    '  campaign.name,',
    '  ad_group.name,',
    '  metrics.cost_micros,',
    '  metrics.impressions,',
    '  metrics.clicks,',
    '  metrics.conversions,',
    '  metrics.average_cpc,',
    '  metrics.ctr,',
    '  metrics.average_cpm',
    'FROM ad_group',
    "WHERE segments.date = '" + dateStr + "'"
  ].join(' ');
}

function upsertKey(date, campaign, adGroup) {
  var d = (date instanceof Date) ? dateToStr(date) : String(date).substring(0, 10);
  return d + '|' + campaign + '|' + adGroup;
}

function dateToStr(date) {
  var y  = date.getFullYear();
  var m  = padTwo(date.getMonth() + 1);
  var dd = padTwo(date.getDate());
  return y + '-' + m + '-' + dd;
}

function parseDateStr(s) {
  var parts = s.split('-');
  return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
}

function padTwo(n) {
  return n < 10 ? '0' + n : String(n);
}