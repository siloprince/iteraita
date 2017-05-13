// dummy
function info() {
}
function onEdit(ev) {
  var targetRange = ev.range;
  var sheet = targetRange.getSheet();
  var sheetName = sheet.getName();
  var spread = sheet.getParent();
  var itemNameListRange = spread.getRangeByName("'"+sheetName+"'!__itemNameList__");
  var itemNameListRow = itemNameListRange.getRow();
  var formulaListRange = spread.getRangeByName("'"+sheetName+"'!__formulaList__");
  var formulaListRow = formulaListRange.getRow();

  var targetRow = targetRange.getRow();
  var targetHeight = targetRange.getHeight();
  
  var namedRanges = spread.getNamedRanges();

  if (targetRow <= itemNameListRow && itemNameListRow <= targetRow + targetHeight - 1) {
    var rawItemNameList = itemNameListRange.getValues()[0];
    var itemNameList = [];
    var nameDupHash = {};
    for (var ri = 0; ri < rawItemNameList.length; ri++) {
      var conved = convertItemName(rawItemNameList[ri].toString());
      if (itemNameList.indexOf(conved) === -1) {
        itemNameList.push(conved);
      } else if (conved.length === 0) {
        itemNameList.push(conved);
      } else {
        if (!(conved in nameDupHash)) {
          nameDupHash[conved]=2;
        }
        for (var rj=0;rj<rawItemNameList.length;rj++) {
          var indexed = conved+'_'+nameDupHash[conved];
          if (itemNameList.indexOf(indexed)===-1) {
            itemNameList.push(indexed);
            break;
          }
          nameDupHash[conved]++;
        }
      }
    }
    // update entire names
    updateRangeNames(itemNameList, itemNameListRange,namedRanges);
    recoverFromFormulas(formulaListRange);

  }
  if (targetRow <= formulaListRow && formulaListRow <= targetRow + targetHeight - 1) {
    var rawFormulaList = formulaListRange.getValues()[0];
    var sheet = targetRange.getSheet();
    var frozenRows = sheet.getFrozenRows();
    var valRow = frozenRows + 1;
    var formulaList = [];
    for (var ri = 0; ri < rawFormulaList.length; ri++) {
      var conved = convertFormula(rawFormulaList[ri], valRow);
      formulaList.push(conved);
    }
    var startw = targetRange.getColumn();
    var endw = startw + targetRange.getWidth() - 1;
    updateFormulas(formulaList, formulaListRange, startw, endw);
  }
  return true;

  function recoverFromFormulas(range) {

    var sheet = range.getSheet();
    var frozenRows = sheet.getFrozenRows();
    var valRow = range.getRow() + 1;
    var width = range.getWidth();
    Logger.log('range:'+valRow+' '+width);
    var formulas = sheet.getRange(valRow, 1, 1, width).getFormulas()[0];
    for (var fi = 0; fi < formulas.length; fi++) {
      var raw = formulas[fi];
      if (raw.length === 0) {
        continue;
      }
      formulas[fi] = raw.replace(/^=/, '');
      if (formulas[fi].indexOf('iferror(T(N(to_text(') === 0) {
        formulas[fi] = formulas[fi].replace(/^iferror\(T\(N\(to_text\(/, '').replace(/\)\)\),""\)$/, '');
      }
      if (formulas[fi].indexOf('+N("__formula__")),"")') > -1) {
        var parts = formulas[fi].split('+N("__formula__")),"")');

        for (var pi = 0; pi < parts.length; pi++) {
          if (pi === parts.length - 1) {
            continue;
          }
          Logger.log(parts[pi]);
          // mod(1103515245 * iferror(index(擬似乱数,N("__")+row()-1
          if (/^([\s\S]*)iferror\(index\(([^;,{&\s\+\-\*\(]+),N\("__"\)\+row\(\)\-([0-9]+)$/.test(parts[pi])) {
            var rest = RegExp.$1;
            var item = RegExp.$2;
            var dcount = RegExp.$3;
            dcount = parseInt(dcount, 10);
            var darray = [];
            for (var di = 0; di < dcount; di++) {
              darray.push('\'');
            }
            parts[pi] = rest + item + darray.join('');
          }
        }
        var ff = parts.join('');
        formulas[fi] = ff;
      }
      // =iferror(T(N(to_text(if(isnumber(222),222,T(N("__@2__"))&222)))),"")
      if (formulas[fi].indexOf('T(N("__@') > -1) {
        if (/\T\(N\("__@(@|[0-9\.]+)__"\)\)&([\s\S]*)\)$/.test(formulas[fi])) {
          formulas[fi] = '@' + RegExp.$1 + ' ' + RegExp.$2;
        }
      }
    }
    range.setValues([formulas]);
  }

  function convertFormula(str, valRow) {
    if (str.toString().trim().length === 0) {
      return '';
    } else {
      var timeout = 0;
      if (/^@(@|[0-9\.]+)\s+/.test(str)) {
        timeout = RegExp.$1;
        str = str.replace(/^@(@|[0-9\.]+)\s+/, '');
        str = 'if(isnumber(' + str + '),' + str + ',T(N("__@' + timeout + '__"))&' + str + ')';
      }
      str = str.toString().replace(/(import(feed|range|html|data|xml)\s*\([^\)]+\))/gi, 'iferror(index($1,N("__")+row()-0+N("__formula__")),"")');
      if (str.indexOf('\'') === -1) {
        return str;
      } else {
        var formulaArray = [];
        var dollers = str.split('\'');
        var dcount = 0;
        for (var di = 0; di < dollers.length; di++) {
          var dj = dollers.length - di - 1;
          var doll = dollers[dj];
          if (doll.length === 0) {
            if (dj !== 0) {
              dcount++;
            }
          } else {
            if (dj === dollers.length - 1) {
              formulaArray.unshift(doll);
              dcount = 1;
              continue;
            }
            var item = '';
            var rest = '';
            if (/[;,{&\s\+\-\*\/\(]/.test(doll)) {
              if (/[;,{&\s\+\-\*\/\(]+([^;,{&\s\+\-\*\/\(]+)$/.test(doll)) {
                item = RegExp.$1;
                rest = doll.replace(new RegExp(item + '$'), '');
              }
            } else {
              item = doll;
              rest = '';
            }
            if (item.length > 0) {
              if (dcount === 0) {
                formulaArray.unshift(rest + item);
              } else {
                formulaArray.unshift(rest + 'iferror(index(' + item + ',N("__")+row()-' + dcount + '+N("__formula__")),"")');
              }
            }
            dcount = 1;
          }
        }
        return formulaArray.join('');
      }
    }
  }
  function updateFormulas(formulaList, range, startw, endw) {
    var sheet = range.getSheet();
    var frozenRows = sheet.getFrozenRows();
    var maxRows = sheet.getMaxRows();
    var dollerHeight = maxRows - frozenRows;
    var dollerRow = frozenRows + 1;
    var row = range.getRow();
    var valRow = row + 2;
    var valHeight = maxRows -row;
    var width = range.getWidth();
    var now = new Date();
    var year = now.getFullYear();
    var month = now.getMonth() + 1;
    var day = now.getDate();
    var hour = now.getHours();
    var minute = now.getMinutes();
    var second = now.getSeconds();
    var datestr = year + '/' + month + '/' + day;
    var timestr = hour + ':' + minute + ':' + second;
    var minrate = 1 / (24 * 60);
    for (var wi = startw - 1; wi < endw; wi++) {
      var input = false;
      var f = formulaList[wi].toString().trim();
      if (f.length === 0) {
        var ff = sheet.getRange(row + 1, wi + 1, valHeight, 1).getFormulas();
        var lastffv = ff[dollerRow].toString().trim();
        for (var fi = dollerRow; fi < ff.length; ff++) {
          var ffv = ff[fi].toString().trim();
          if (ffv.length === 0) {
            input = true;
            break;
          }
          if (ffv !== lastffv) {
            input = true;
            break;
          }
        }
      }
      if (!input) {
        if (f.length === 0) {
          sheet.getRange(row + 1, wi + 1, valHeight, 1).setFormula(f);
          sheet.getRange(4,wi+1).setFormula('iferror(sparkline(indirect(address(9,column(),4)&":"&address(108,column(),4))),"")');

        } else {
          var _row = dollerRow;
          var _col = wi + 1;
          var _height = dollerHeight;
          var _width = 1;
          // set to protocode
          sheet.getRange(row + 1, _col).setFormula('iferror(T(N(to_text(' + f + '))),"")');
          if (f.indexOf('N("__formula__")') > -1) {
            // remove errors on initals
            var errors = sheet.getRange(5,_col,dollerRow-5,1).getValues();
            var corrects = [];
            var hasError=false;
            for (var ei=0;ei<errors.length;ei++) {
              if (/^#.*!$/.test(errors[ei][0])) {
                corrects.push(['']);
                hasError=true;
              } else {
                corrects.push([errors[ei][0]])
              }
            }
            if (hasError) {
              sheet.getRange(5,_col,dollerRow-5,1).setValues(corrects);
            }
          } else {
            sheet.getRange(row + 3, _col, frozenRows+1-(row+3), _width).setFormula('iferror('+f+',"")');
          }
          if (f.indexOf('__@') > -1) {
            sheet.getRange(_row, _col, _height, _width).setFormula('');
            var therange = sheet.getRange(_row, _col);
            therange.setFormula(f);
            var val = therange.getValue();
            var time = (new Date()).getTime();
            therange.setFormula('"' + val + '"&T(N("__#' + time + '__"))');

          } else {
            sheet.getRange(_row, _col, _height, _width).setFormula(f);
          }
        }
      }
    }
  }

  function updateRangeNames(itemNameList, itemRange, namedRanges) {
    var sheet = itemRange.getSheet();
    var spread = sheet.getParent();
    var valRow = 1;
    var maxRows = sheet.getMaxRows();
    var valHeight = maxRows;
    var maxCols = sheet.getMaxColumns();

    var maxHeightHash = {};
    for (var ni = 0; ni < namedRanges.length; ni++) {
      var rangeName = namedRanges[ni].getName();
      if (rangeName.indexOf('_') !== 0) {
        var range = namedRanges[ni].getRange();
        var width = range.getWidth();
        if (width === 0) {
          namedRanges[ni].remove();
          continue;
        }
        if (width !== 1) {
          continue;
        }
        var row = range.getRow();
        if (row !== valRow) {
          continue;
        }
        var key = range.getColumn().toString();
        var height = range.getHeight();
        if (!(key in maxHeightHash)) {
          maxHeightHash[key] = { height: height, namedRange: namedRanges[ni] };
        } else if (maxHeightHash[key].height <= height) {
          maxHeightHash[key].namedRange.remove();
          maxHeightHash[key].height = height;
          maxHeightHash[key].namedRange = namedRanges[ni];
        }
      }
    }
    for (var ii = 0; ii < itemNameList.length; ii++) {
      var col = ii + 1;
      var key = col.toString();
      var range = sheet.getRange(valRow, col, valHeight, 1);
      var rangeName = itemNameList[ii];
      if (rangeName.length > 0) {
        if (!(key in maxHeightHash)) {
          spread.setNamedRange(rangeName, range);
        } else {
          maxHeightHash[key].namedRange.setName(rangeName);
          maxHeightHash[key].namedRange.setRange(range);
        }
      } else {
        if (key in maxHeightHash) {
          maxHeightHash[key].namedRange.remove();
        }
      }
    }
    itemRange.setValues([itemNameList]);
  }
  function convertItemName(str) {
    // TODO: support more bad characters
    if (/[０１２３４５６７８９]/) {
      str = str.replace(/０/g, '0');
      str = str.replace(/１/g, '1');
      str = str.replace(/２/g, '2');
      str = str.replace(/３/g, '3');
      str = str.replace(/４/g, '4');
      str = str.replace(/５/g, '5');
      str = str.replace(/６/g, '6');
      str = str.replace(/７/g, '7');
      str = str.replace(/８/g, '8');
      str = str.replace(/９/g, '9');
    }
    if (/^[a-zA-Z]$/.test(str)) {
      str = '英' + str.toUpperCase();
    } else if (/^[0-9]/.test(str)) {
      str = '数' + str;
    } else if (/[_\s<>=~!#'"%&;:,\(\)\|\.\\\^\+\-\*\/\?\$　＜＞＝〜！＃’”％＆；：，（）｜．＼＾＋＊／？＄]/.test(str)) {
      str = str.replace(/[\s<>=~!#'"%&;:,\(\)\|\.\\\^\+\-\*\/\?\$　＜＞＝〜！＃’”％＆；：，（）｜．＼＾＋＊／？＄]/g, '＿');
    }
    return str;
  }
}

function atprocess(spread, byhand) {
  var sheet = spread.getActiveSheet();
  var sheetName = sheet.getName();
  var range = spread.getRangeByName("'"+sheetName+"'!__formulaList__");
  var frow = range.getRow() + 1;
  var formulaList = range.getValues()[0];
  var sheet = range.getSheet();
  var frozenRows = sheet.getFrozenRows();
  var startRow = frozenRows + 1;
  var maxRows = sheet.getMaxRows();
  for (var fi = 0; fi < formulaList.length; fi++) {
    var col = fi + 1;
    if (!(formulaList[fi])) {
      continue;
    }
    if (formulaList[fi].indexOf('@') === 0) {
      var timeout = -1;
      if (!byhand) {
        if (/^@([0-9\.]+)/.test(formulaList[fi])) {
          timeout = RegExp.$1;
          timeout = parseFloat(timeout);
        }
      } else {
        if (/^@@/.test(formulaList[fi])) {
          timeout = 0;
        }
      }
      if (timeout !== -1) {
        var formulas = sheet.getRange(startRow, col, maxRows - frozenRows, 1).getFormulas();
        for (var fj = 0; fj < formulas.length; fj++) {
          if (fj === formulas.length - 1) {
            break;
          }
          var fv = formulas[fj][0];
          if (!fv) {
            fv = '';
          } else {
            fv = fv.toString();
          }
          var nfv = formulas[fj + 1][0];
          if (!nfv) {
            nfv = '';
          } else {
            nfv = nfv.toString();
          }
          if (fv.indexOf('&T(N("__#') > -1 && nfv.length === 0) {

            if (/&T\(N\("__#([0-9\.]+)__/.test(fv)) {
              var timestamp = RegExp.$1;
              timestamp = parseFloat(timestamp);
              var time = (new Date()).getTime();

              if (time > timestamp + timeout * 60 * 1000) {

                var therange = sheet.getRange(startRow + fj + 1, col);
                setTimestamp(therange, frow);
              }
            }
            break;
          } else if (fv.length === 0) {

            var therange = sheet.getRange(startRow + fj, col);
            setTimestamp(therange, frow);
            break;
          }
        }
      }
    }
  }
  return 'atprocess ' + byhand;
}
function setTimestamp(range, frow) {
  var sheet = range.getSheet();
  var col = range.getColumn();
  var f = sheet.getRange(frow, col).getFormula().replace(/^=/, '').replace(/iferror\(T\(N\(to_text\(/, '').replace(/\)\)\),""\)$/, '');
  range.setFormula(f);
  var time = (new Date()).getTime();
  var val = range.getValue();
  range.setFormula('"' + val + '"&T(N("__#' + time + '__"))');
}
function clear(spread, all) {
  var range = spread.getActiveRange();
  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  var formulaRange = spread.getRangeByName("'"+sheetName+"'!__formulaList__");
  var frozenRows = sheet.getFrozenRows();
  var maxRows = sheet.getMaxRows();
  var values;
  var offset = 1;
  if (!all) {
    offset = range.getColumn();
    values = sheet.getRange(formulaRange.getRow(), range.getColumn(), 1, range.getWidth()).getValues();
  } else {
    values = formulaRange.getValues();
  }
  for (var vi = 0; vi < values[0].length; vi++) {
    var val = values[0][vi];
    if (/^\s*@/.test(val.toString())) {
      sheet.getRange(frozenRows + 1, vi + offset, maxRows - frozenRows, 1).setValue('');
    }
  }
  return "clear:" + all;
}
function onOpen(spread, ui, sidebar) {
  if (!sidebar) {
    ui.createMenu('Iteraita').addItem('サイドバーを開く', 'openSidebar').addItem('リセット', 'reset').addToUi();
    return;
  }
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Iteraita')
    .setWidth(300);
  ui.showSidebar(html);
  return true;
}
function reset(spread) {
  var sheet = spread.getActiveSheet();
  var sheetName = sheet.getName();
  var namedRanges = spread.getNamedRanges();
  for (var ni=0;ni<namedRanges.length;ni++) {
    if (namedRanges[ni].getName().indexOf('__')!==0) {
      namedRanges[ni].remove();
    }
  }
  var rows = 100+8;
  var cols = 26;
  sheet.clearContents();
  var maxRows = sheet.getMaxRows();
  if (maxRows < rows ) {
    sheet.insertRows(maxRows,rows-maxRows);
  } else if (maxRows > rows) {
    sheet.deleteRows(rows+1,maxRows-rows);
  }
   var maxCols = sheet.getMaxColumns();
  if (maxCols < cols ) {
    sheet.insertColumns(maxCols,cols-maxCols);
  } else if (maxCols > cols) {
    sheet.deleteColumns(cols+1,maxCols-cols);
  }
  for (var ci=0;ci<cols;ci++) {
    sheet.setColumnWidth(ci+1,120);
  }
  var formulaRange = spread.getRange("'"+sheetName+"'!__formulaList__");
  formulaRange.setWrap(true);
  // TODO
  for (var ri=0;ri<rows;ri++) {
    var height=21;
    if (ri===1) {
      height = 70;
    } else if (ri===2) {
      height = 2;
    } else if (ri===3) {
      height = 70;
    }
    sheet.setRowHeight(ri+1,height);
  }
  sheet.getRange(4,1,1,cols).setFormula('iferror(sparkline(indirect(address(9,column(),4)&":"&address(108,column(),4))),"")');
}