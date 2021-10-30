// Source: https://git.io/Jthwt
// warwickallen/GMail-Archive-old-messages-by-label-and-category
// ef742d671a2e03ea1ca8025c80b34b4a14067979
// This code was based on John Day's example at:
// https://www.johneday.com/422/time-based-gmail-filters-with-google-apps-script
// Why Search per label vs total inbox?
// Big(0) Notation:
// Inbox: Log(n^2)
// - Great at small inboxes, bad at large ones
// - # operations grow by the size of the inbox (uncontrolled varaible)
// Label: Log(n)
// - More consistent execution even as an inbox grows
// - # of operations is determined by a fixed number of labels/categories

/*jshint esversion: 6 */

function go() {
  console.time("go() time");
  let fname = "GMail Inbox Cleanup (gScript)";
  let files = DriveApp.getFilesByName(fname);
  let spreadsheet;

  while ((typeof spreadsheet !== "object") && files.hasNext()) {
    try {
      spreadsheet = SpreadsheetApp.openById(files.next().getId());
    } catch (e) {
      console.warn(e.message);
    }
  }

  if (typeof spreadsheet !== "object") {
    console.error("Cannot find a Google speadsheet called '%s'.", fname);
    return -1;
  }

  let count = archiveThreadsByLabel(spreadsheet);
  count += archiveThreadsByCategory(spreadsheet);
  console.timeEnd("go() time");
  console.info("Archived %d threads.", count);
}

function archiveThreadsByLabel(spreadsheet) {
  let archiveItems = getSheetData(spreadsheet, "Archive by Label");
  return archiveThreads("label", archiveItems);
}

function archiveThreadsByCategory(spreadsheet) {
  let archiveItems = getSheetData(spreadsheet, "Archive by Category");
  return archiveThreads("category", archiveItems);
}

function getSheetData(spreadsheet, sheetName) {
  let dataArray = spreadsheet.getSheetByName(sheetName)
    .getDataRange()
    .getValues();
  dataArray.shift(); // Remove the headings row.
  // Filter rows missing a name or age
  let filteredArray = dataArray.filter(function(datum) {
    if (datum[0] === "") {
      // If the spreadsheet has empty rows, this warning blows up.
      // console.info('Label/Category has no "name". Skipping...');
      return false;
    }

    if (datum[1] === "") {
      console.warn("Label/Category '%s' has no 'age'. Skipping...", datum[0]);
      return false;
    }

    return true;
  });

  let dataObjects = filteredArray.map((datum) => {
    return {
      "name": datum[0],
      "queryName": datum[0].replace(/[\/() ]/g, "-"),
      "age": datum[1],
      "markAsRead": datum[2]
    };
  });

  // SETTING: Sort by age ASC
  Date.prototype.subHours = function(h) {
    this.setTime(this.getTime() - (h * 60 * 60 * 1000));
    return this;
  };

  Date.prototype.subDays = function(d) {
    this.setTime(this.getTime() - (d * 24 * 60 * 60 * 1000));
    return this;
  };

  Date.prototype.subMonths = function(m) {
    this.setTime(this.getTime() - (m * 30 * 24 * 60 * 60 * 1000));
    return this;
  };

  Date.prototype.subYears = function(y) {
    this.setTime(this.getTime() - (y * 12 * 30 * 24 * 60 * 60 * 1000));
    return this;
  };

  dataObjects.sort(function(x, y) {
    let xDate = new Date();
    let yDate = new Date();
    let xNumber = x.age.slice(0, -1);
    let yNumber = y.age.slice(0, -1);
    let xUnit = x.age.substr(x.age.length - 1);
    let yUnit = y.age.substr(y.age.length - 1);

    switch (xUnit) {
      case "h": xDate.subHours(xNumber); break;
      case "d": xDate.subDays(xNumber); break;
      case "m": xDate.subMonths(xNumber); break;
      case "y": xDate.subYears(xNumber); break;
    }

    switch (yUnit) {
      case "h": yDate.subHours(yNumber); break;
      case "d": yDate.subDays(yNumber); break;
      case "m": yDate.subMonths(yNumber); break;
      case "y": yDate.subYears(yNumber); break;
    }

    return yDate - xDate;
  });

  // console.log('dataObjects: %s', JSON.stringify(dataObjects, null ,2));
  return dataObjects;
}

function archiveThreads(archiveType, archiveItems) {
  // SETTING: Limit query & archive/read to 100 threads at a time
  let batchSize = 100;
  let archivedCount = 0;

  if (archiveItems === undefined || archiveItems.length == 0) {
    console.warn("No rules found. Check 'Archive by %s'.", archiveType);
    return archivedCount;
  }

  archiveItems.forEach(function(archiveObj) {
    // console.log('archiveObj: %s', JSON.stringify(archiveObj, null, 2));
    let batchedCount = 0;
    let query = "in:inbox -is:starred ";
    query += `${archiveType}:${archiveObj.queryName} older_than:${archiveObj.age}`;
    console.time(`archiveThreads(${archiveObj.name}) time`);

    while (1) {
      let batchThreads = GmailApp.search(query, 0, batchSize);

      if (batchThreads === undefined || batchThreads.length == 0) {
        break;
      }

      GmailApp.moveThreadsToArchive(batchThreads);

      if (archiveObj.markAsRead === true) {
        GmailApp.markThreadsRead(batchThreads);
      }

      batchedCount += batchThreads.length;

      // SETTING: Throttle the 2-3x GmailApp API calls to once every second
      Utilities.sleep(1000);
      // break; // DEBUG
    }

    console.timeEnd(`archiveThreads(${archiveObj.name}) time`);
    console.info(
      "Archived %d threads (%s) '%s' that are older than %s.",
      batchedCount, archiveType, archiveObj.name, archiveObj.age
    );
    archivedCount += batchedCount;
  });

  return archivedCount;
}
