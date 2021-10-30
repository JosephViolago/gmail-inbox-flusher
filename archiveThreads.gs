function archiveLabeled() {
  var myLabels  = {
    '"calendar"'                   : "1d",
    '"dev/ci"'                     : "1d",
    '"dev/clubhouse"'              : "1d",
    '"dev/cr"'                     : "1d",
    '"dev/datadog"'                : "1d",
    '"dev/deploy"'                 : "1h",
    '"dev/devops"'                 : "1d",
    '"dev/dojo"'                   : "1d",
    '"dev/errors"'                 : "1d",
    '"dev/github"'                 : "1d",
    '"dev/hipchat"'                : "1h",
    '"dev/jira"'                   : "1h",
    '"dev/ooo"'                    : "1d",
    '"feedback"'                   : "1d",
    '"gdrive"'                     : "1d",
    '"office/bamboohr"'            : "1d",
    '"office/birthday"'            : "1d",
    '"office/cleanup"'             : "1d",
    '"office/perks"'               : "1d",
    '"alertos"'                    : "1d",
    '"alertos/deals"'              : "1d",
    '"alertos/deals/am expansion"' : "1d",
    '"alertos/deals/kickoffs"'     : "1d",
    '"alertos/culture"'            : "1d",
    '"alertos/forest"'             : "1d",
    '"alertos/new hires"'          : "1d",
    '"alertos/ocean"'              : "1d",
    '"vendors"'                    : "1d",
  };
  var batchSize = 100; // Process up to 100 threads at once

  for (aLabel in myLabels) {
    var mySearch = "'label:inbox label:" + aLabel + "older_than:" + myLabels[aLabel] + "'";

    while (GmailApp.search(mySearch, 0, 1).length == 1) {
      GmailApp.moveThreadsToArchive(GmailApp.search(mySearch, 0, batchSize));
    }
  }
}
