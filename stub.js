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
function clear(all) {
  var spread = SpreadsheetApp.getActive();
  iteraita.clear(spread,all);
}
function onOpen() {
    var sidebar = iteraita.getSidebar();
    SpreadsheetApp.getUi() 
      .showSidebar(sidebar);
}