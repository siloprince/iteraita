function onEdit(ev) {
  return iteraita.onEdit(ev);
}
function cron() {
  var spread = SpreadsheetApp.getActive();
  return iteraita.atprocess(spread, false);
}
function hand() {
  var spread = SpreadsheetApp.getActive();
  return iteraita.atprocess(spread, true);
}
function clear(all) {
  var spread = SpreadsheetApp.getActive();
  return iteraita.clear(spread, all);
}
function onOpen() {
  var spread = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  return iteraita.onOpen(spread,ui,false);
}
function openSidebar() {
  var spread = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  return iteraita.onOpen(spread,ui,true);
}