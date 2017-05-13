
  window.addEventListener('message', function(event) {
    if (!/googleusercontent.com$/.test(event.origin) ) return;
    event.source.postMessage(document.querySelector('table#weekly').innerHTML, event.origin);
  }, false);