
  var button = '<button width="15" height="200" onclick=\'document.querySelector("iframe#conflu").contentWindow.postMessage("HELLO!", "https://confluence.gree-office.net/pages/viewpage.action?pageId=188158149");return true;\'>貼付ける</button>';
  var script = '<script type="text/javascript">window.addEventListener("message", function(event) { \
  var tmp = document.createElement("div");\
  tmp.innerHTML="<table>"+event.data+"</table>";\
  var json={　\
    "conflu":tmp.querySelector("td#conflu_td").textContent,\
    "jira":tmp.querySelector("td#jira_td").textContent\
  };\
  google.script.run.withFailureHandler(function(err){console.error(err);}).withSuccessHandler(function(){google.script.host.close();}).paste(JSON.stringify(json))}, false);</script>';

  var pageId = 187132295;
  var url = 'https://xxxxx/pages/viewpage.action?pageId=188158149&noedit';
  var html = HtmlService.createHtmlOutput('<table width="100%"><tr><td width="30"　valign="top">'+button+'</td><td><div style="width:100%;height:100%;"><div style="width:100%;height:100%;"><iframe sandbox="allow-top-navigation allow-scripts allow-same-origin" id="conflu" style="border:none;width:100%;height:100%;padding:0;margin:0;" src="'+url+'"></iframe></div></div></td></tr></table>'+script)
      .setWidth(800)
      .setHeight(160);
  SpreadsheetApp.getUi()
      .showModalDialog(html, '画面に表示される記録をシートに貼り付けるには、「貼付け」ボタンを押して下さい');