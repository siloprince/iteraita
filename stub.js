function onEdit(ev) {
  iteraita.onEdit(ev);
}
function cron() {
  var spread = SpreadsheetApp.getActive();
  iteraita.atprocess(spread, false);
}
function hand() {
  var spread = SpreadsheetApp.getActive();
  iteraita.atprocess(spread,true);
}
function onOpen() {
  iteraita.onOpen();
}