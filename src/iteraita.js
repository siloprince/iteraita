// dummy
function info() {
}
function onEdit(ev) {
  var targetRange = ev.range;
  var sheet = targetRange.getSheet();
  var sheetName = sheet.getName();
  if (sheetName !== 'Sheet1') {
    return;
  }
  var spread = sheet.getParent();
  var targetRow = targetRange.getRow();
  var targetHeight = targetRange.getHeight();
  var targetColumn = targetRange.getColumn();
  var targetWidth = targetRange.getWidth();
  var itemNameListRange = spread.getRangeByName("'" + sheetName + "'!__itemNameList__");
  var formulaListRange = spread.getRangeByName("'" + sheetName + "'!__formulaList__");
  processNameRange(spread, sheet, targetRow, targetHeight, itemNameListRange, formulaListRange);
  processSingleEmptyFormula(spread, sheet, ev, targetRow, targetColumn, formulaListRange);
  processFormulaList(spread, sheet, targetRow, targetHeight, targetColumn, targetWidth, itemNameListRange, formulaListRange);
  return true;
}
function getObjectType(object) {
  return Object.prototype.toString.call(object).replace(/\[object (\w+)\]$/, '$1');
}

function processSingleEmptyFormula(spread, sheet, ev, targetRow, targetColumn, formulaListRange) {
  // if onEdit for multiple cell ev.value = {}, ev.oldValue = undefined
  // if onEdit for single cell ev.value = { oldValue: "x"}, ev.oldValue = "x" if changed from "x" to ""
  //                           ev.value = "x", ev.oldValue = undefined if chagned from "" to "x"
  //Logger.log(JSON.stringify(ev.value)+':'+JSON.stringify(ev.oldValue));
  // value -> empty
  if (getObjectType(ev.value) === 'Object' && getObjectType(ev.oldValue) === 'String') {
    if (targetRow === formulaListRange.getRow()) {
      var maxRows = sheet.getMaxRows();
      sheet.getRange(targetRow + 1, targetColumn, maxRows - targetRow, 1).setValue('');
    }
  }
}
function getColumnLabel(index) {
  var mod = ((index - 1) % 26);
  var modStr = String.fromCharCode("A".charCodeAt(0) + mod);
  var quot = ((index - 1 - mod) / 26);
  var quotStr = String.fromCharCode("A".charCodeAt(0) - 1 + quot);
  if (quot == 0) { quotStr = ""; }
  return quotStr + modStr;
}
function processNameRange(spread, sheet, targetRow, targetHeight, itemNameListRange, formulaListRange) {
  var itemNameListRow = itemNameListRange.getRow();
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
        if (conved.indexOf('_') <= 0) {
          if (!(conved in nameDupHash)) {
            nameDupHash[conved] = 2;
          }
        } else {
          // TODO: multiple _\d_\d
          conved = conved.replace(/_[0-9]+$/, '');
          if (!(conved in nameDupHash)) {
            nameDupHash[conved] = 2;
          }
        }
        for (var rj = 0; rj < rawItemNameList.length; rj++) {
          var indexed = conved + '_' + nameDupHash[conved];
          if (itemNameList.indexOf(indexed) === -1) {
            itemNameList.push(indexed);
            break;
          }
          nameDupHash[conved]++;
        }
      }
    }
    // update entire names
    updateRangeNames(itemNameList, itemNameListRange, namedRanges);
    recoverFromFormulas(formulaListRange, itemNameList);
  }
  return;

  function recoverFromFormulas(range, itemNameList) {

    var sheet = range.getSheet();
    var frozenRows = sheet.getFrozenRows();
    var formulaRow = range.getRow();
    var valRow = formulaRow + 1;
    var width = range.getWidth();
    var formulas = sheet.getRange(valRow, 1, 1, width).getFormulas()[0];
    var tn_header = 'iferror(T(N(to_text(';
    var tn_footer = '))),"")';
    for (var fi = 0; fi < formulas.length; fi++) {
      var raw = formulas[fi];
      if (raw.length === 0) {
        continue;
      }
      if (raw.indexOf('N("__IMPORTRANGE__")') !== -1) {
        if (/T\(N\("__IMPORTRANGE__"\)\)&T\(N\("([^"]+)"\)\)/.test(raw)) {
          var filename = RegExp.$1;
          var itemName = itemNameList[fi];
          formulas[fi] = 'import(' + filename + ')';
          if (updateImport(spread, itemName, filename)) {
            sheet.getRange(formulaRow + 3, fi + 1).setFormula('offset(_' + itemName + ',' + ((formulaRow + 3) - 1) + ',0)');
          } else {
            sheet.getRange(formulaRow + 3, fi + 1).setFormula('');
          }
        }
        continue;
      }
      if (raw.indexOf('=') === 0) {
        formulas[fi] = raw.slice(1);
      }
      if (formulas[fi].indexOf(tn_header) === 0) {
        formulas[fi] = formulas[fi].slice(tn_header.length, -tn_footer.length);
      }
      if (formulas[fi].indexOf('+N("__formula__")),1/0)') > -1) {
        var parts = formulas[fi].split('+N("__formula__")),1/0)');

        for (var pi = 0; pi < parts.length; pi++) {
          if (pi === parts.length - 1) {
            continue;
          }
          // mod(1103515245 * iferror(index(擬似乱数,N("__n__")+row()-1
          if (/^([\s\S]*)iferror\(index\((filter\(offset\(|offset\(offset\(|)(.*),([\.0-9--]*)\+N\("__([^_]+)__"\)([\s\S]*)$/.test(parts[pi])) {
            var rest = RegExp.$1;
            var item = RegExp.$3;
            var num = RegExp.$4;
            var func = RegExp.$5;
            var end = RegExp.$6;
            if (func.indexOf('prev') === 0) {
              var p1 = 0;
              var p2 = 0;
              var dcount = 0;
              if (/if\(""="([0-9--]*)",N\("__param___"\)\+len\("([^"]+)"\)/.test(end)) {
                p1 = RegExp.$1;
                p2 = RegExp.$2;
                if (p1 === '') {
                  dcount = p2.length;
                } else {
                  dcount = parseInt(p1, 10);
                }
              }
              var darray = [];
              if (dcount <= 0) {
                darray.push('\'{' + dcount + '}');
              } else {
                for (var di = 0; di < dcount; di++) {
                  darray.push('\'');
                }
              }
              var argvend = ')';
              if (item.indexOf(',' + argvend) !== -1) {
                parts[pi] = rest + darray.join('');
              } else {
                item = item.slice(0, -argvend.length);
                item = item.slice(item.indexOf(',') + 1);
                item = item.slice(item.indexOf(',') + 1);
                parts[pi] = rest + item + darray.join('');
              }
            } else if (func.indexOf('left') === 0) {
              // -len("$2")
              // if(""="$4",N("__param___")+len("$2")
              var p1 = 0;
              var p2 = 0;
              var dcount = 0;
              if (/if\(""="([0-9--]*)",N\("__param___"\)\+len\("([^"]+)"\)/.test(item)) {
                p1 = RegExp.$1;
                p2 = RegExp.$2;
                if (p1 === '') {
                  dcount = p2.length;
                } else {
                  dcount = parseInt(p1, 10);
                }
              }
              var darray = [];
              if (dcount <= 0) {
                darray.push('`{' + dcount + '}');
              } else {
                for (var di = 0; di < dcount; di++) {
                  darray.push('`');
                }
              }
              var argvend = ')';
              if (item.indexOf(',' + argvend) !== -1) {
                parts[pi] = rest + darray.join('');
              } else {
                item = item.slice(item.indexOf('"') + 1);
                item = item.sliace(0, -(item.length - item.indexOf('"')));
                parts[pi] = rest + item + darray.join('');
              }
            } else if (func.indexOf('argv') === 0) {
              num = -parseInt(num, 10);
              var argvend = ')';
              if (item.indexOf(',' + argvend) !== -1) {
                parts[pi] = rest + '$' + num + '';
              } else {
                item = item.slice(0, -argvend.length);
                item = item.slice(item.indexOf(',') + 1);
                item = item.slice(item.indexOf(',') + 1);
                parts[pi] = rest + item + '$' + num + '';
              }
            } else if (func.indexOf('head') === 0 || func.indexOf('last') === 0) {
              parts[pi] = rest + func + '(' + item + ')';
              if (num !== '') {
                num = parseInt(num, 10);
                if (func.indexOf('last') === 0) {
                  num = -num;
                  if (end.indexOf('+1-if(""="",1,0)') > -1) {
                    num += 1;
                  }
                } else {
                  if (end.indexOf('-1+if(""="",1,0)') > -1) {
                    num += 1;
                  }
                }
                if (num !== 1) {
                  parts[pi] = rest + func + '(' + item + ',' + num + ')';
                }
              }
            } else {
              parts[pi] = rest + func + '(' + item + ')';
            }
          }
        }
        var ff = parts.join('');
        formulas[fi] = ff;
      }
      // 
      // =T("__EMPTY__")+N("__SIDE__")+T("__EMPTY__")+N("__SIDE__")+22+N("__SIDE__")+3
      var sidesep = '+N("__SIDE__")+';
      var inits = [];
      if (formulas[fi].indexOf(sidesep) > -1) {
        var farray = [];
        var parts = formulas[fi].split(sidesep);
        parts.pop();
        parts.shift();
        var nonempty = 0;
        for (var pi = 0; pi < parts.length; pi++) {
          if (pi === 0) {
            farray.push(parts[pi]);
          } else if (parts[pi] === 'T("__EMPTY__")') {
            if (nonempty) {
              inits.push('[]');
            }
          } else {
            inits.push('[' + parts[pi] + ']');
            nonempty = 1;
          }
        }
        formulas[fi] = farray.join('\n');
      }
      // =iferror(iferror(N("__if__")+if(and( 自然数>10,isnumber(自然数)),N("__then__")+自然数  ,1/0)),N("__if__")+0/1),"")
      var ifsep = 'N("__if__")+if(and(';
      if (formulas[fi].indexOf(ifsep) > -1) {
        var farray = [];
        var elsep1 = ',isnumber(';
        var elsep2 = '),N("__then__")+';
        var tailsep = ',1/0)),';
        var parts = formulas[fi].split(ifsep);
        for (var pi = 0; pi < parts.length; pi++) {
          if (pi === 0 || pi === parts.length - 1) {
            continue;
          }
          var condval = parts[pi].split(elsep2);
          var cond = condval[0].split(elsep1);
          var val = condval[1].slice(0, -tailsep.length).trim();
          if (parts.length === 3) {
            farray.push(val + ' | ' + cond[0].trim() + '');
          } else {
            farray.push(val + ' | { ' + cond[0].trim() + ' }');
          }
        }
        formulas[fi] = farray.join('\n');
      }
      if (inits.length > 0) {
        formulas[fi] += '\n' + inits.join('\n');
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
    str = str.toUpperCase();
    var head = str.charCodeAt(0);
    if (str.length === 1) {
      // C and R is not available
      if ('A'.charCodeAt(0) <= head && head < 'Z'.charCodeAt(0)) {
        str = '英' + str;
        return str;
      }
    }
    if (str.indexOf('０') > -1) {
      str = str.replace(/０/g, '0');
    }
    if (str.indexOf('１') > -1) {
      str = str.replace(/１/g, '1');
    }
    if (str.indexOf('２') > -1) {
      str = str.replace(/２/g, '2');
    }
    if (str.indexOf('３') > -1) {
      str = str.replace(/３/g, '3');
    }
    if (str.indexOf('４') > -1) {
      str = str.replace(/４/g, '4');
    }
    if (str.indexOf('５') > -1) {
      str = str.replace(/５/g, '5');
    }
    if (str.indexOf('６') > -1) {
      str = str.replace(/６/g, '6');
    }
    if (str.indexOf('７') > -1) {
      str = str.replace(/７/g, '7');
    }
    if (str.indexOf('８') > -1) {
      str = str.replace(/８/g, '8');
    }
    if (str.indexOf('９') > -1) {
      str = str.replace(/９/g, '9');
    }
    if (/[_\s<>=~!#'"%&;:,\(\)\|\.\\\^\+\-\*\/\?\$　＜＞＝〜！＃’”％＆；：，（）｜．＼＾＋＊／？＄]/.test(str)) {
      str = str.replace(/[\s<>=~!#'"%&;:,\(\)\|\.\\\^\+\-\*\/\?\$　＜＞＝〜！＃’”％＆；：，（）｜．＼＾＋＊／？＄]/g, '＿');
    }
    head = str.charCodeAt(0);
    var last = str.charCodeAt(str.length - 1);
    if ('A'.charCodeAt(0) <= head && head < 'Z'.charCodeAt(0) && '0'.charCodeAt(0) <= last && last < '9'.charCodeAt(0) && str.indexOf('_') === -1) {
      var sarray = [];
      var numflag = 0;
      for (var si = 0; si < str.length; si++) {
        var sv = str.charCodeAt(si);
        if ('A'.charCodeAt(0) <= sv && sv < 'Z'.charCodeAt(0)) {
          if (numflag) {
            numflag = 0;
            break;
          }
          sarray.push(str[si]);
        } else if ('0'.charCodeAt(0) <= sv && sv < '9'.charCodeAt(0)) {
          if (numflag === 0) {
            sarray.push('_');
            numflag = 1;
          }
          sarray.push(str[si]);
        } else {
          numflag = 0;
          break;
        }
      }
      if (numflag) {
        str = sarray.join('');
      }
    }
    if ('0'.charCodeAt(0) <= head && head < '9'.charCodeAt(0)) {
      str = '数' + str;
      return str;
    }
    return str;
  }
}

function processFormulaList(spread, sheet, targetRow, targetHeight, targetColumn, targetWidth, itemNameListRange, formulaListRange) {
  var itemNameList = itemNameListRange.getValues()[0];
  var formulaListRow = formulaListRange.getRow();
  if (targetRow <= formulaListRow && formulaListRow <= targetRow + targetHeight - 1) {
    var rawFormulaList = formulaListRange.getValues()[0];
    var frozenRows = sheet.getFrozenRows();
    var valRow = frozenRows + 1;
    var formulaList = [];
    for (var ri = 0; ri < rawFormulaList.length; ri++) {
      var raw = rawFormulaList[ri].toString().trim();
      var conved = convertFormula(raw, valRow);
      formulaList.push(conved);
      if (raw !== conved) {
        sheet.getRange(formulaListRow, ri + 1).setValue(conved);
      }
    }
    var startw = targetColumn;
    var endw = startw + targetWidth - 1;
    updateFormulas(formulaList, formulaListRange, startw, endw, itemNameList);
  }
  return;
  function replaceAt(str, char, at) {
    return str.substr(0, at) + char + str.substr(at + 1, str.length);
  }
  function convertFormula(str, valRow) {
    if (str.length === 0) {
      return '';
    } else {
      // zen to han

      for (var si = 0; si < str.length; si++) {
        var code = str.charCodeAt(si);
        var char = 0;
        if (code === '　'.charCodeAt(0)) {
          char = ' ';
        } else if (code === '０'.charCodeAt(0)) {
          char = '0';
        } else if (code === '１'.charCodeAt(0)) {
          char = '1';
        } else if (code === '２'.charCodeAt(0)) {
          char = '2';
        } else if (code === '３'.charCodeAt(0)) {
          char = '3';
        } else if (code === '４'.charCodeAt(0)) {
          char = '4';
        } else if (code === '５'.charCodeAt(0)) {
          char = '5';
        } else if (code === '６'.charCodeAt(0)) {
          char = '6';
        } else if (code === '７'.charCodeAt(0)) {
          char = '7';
        } else if (code === '８'.charCodeAt(0)) {
          char = '8';
        } else if (code === '９'.charCodeAt(0)) {
          char = '9';
        } else if (code === '＋'.charCodeAt(0)) {
          char = '+';
        } else if (code === '＊'.charCodeAt(0)) {
          char = '*';
        } else if (code === '｀'.charCodeAt(0)) {
          char = '`';
        } else if (code === '"'.charCodeAt(0)) {
          char = '"';
        } else if (code === '.'.charCodeAt(0)) {
          char = '.';
        } else if (code === '，'.charCodeAt(0)) {
          char = ',';
        } else if (code === '（'.charCodeAt(0)) {
          char = '(';
        } else if (code === '）'.charCodeAt(0)) {
          char = ')';
        } else if (code === '＜'.charCodeAt(0)) {
          char = '<';
        } else if (code === '＝'.charCodeAt(0)) {
          char = '=';
        } else if (code === '＞'.charCodeAt(0)) {
          char = '>';
        } else if (code === '｛'.charCodeAt(0)) {
          char = '{';
        } else if (code === '｝'.charCodeAt(0)) {
          char = '}';
        } else if (code === '｜'.charCodeAt(0)) {
          char = '|';
        } else if (code === '？'.charCodeAt(0)) {
          char = '?';
        } else if (code === '＄'.charCodeAt(0)) {
          char = '$';
        } else if (code === '＆'.charCodeAt(0)) {
          char = '&';
        } else if (code === '％'.charCodeAt(0)) {
          char = '%';
        } else if (code === '＃'.charCodeAt(0)) {
          char = '#';
        } else if (code === '！'.charCodeAt(0)) {
          char = '!';
        } else if (code === '＾'.charCodeAt(0)) {
          char = '^';
        } else if (code === '＠'.charCodeAt(0)) {
          char = '@';
        } else if (code === ';'.charCodeAt(0)) {
          char = ';';
        } else if (code === '：'.charCodeAt(0)) {
          char = ':';
        } else if (code === '’'.charCodeAt(0)) {
          char = '\'';
        }
        if (char) {
          str = replaceAt(str, char, si);
        }
      }
      // double quote to single quote
      // zen - to han - 
      var strArray = [];
      if (str.indexOf('ー') > -1) {
        var lastcode = 0;
        var code = 0;
        for (var si = 0; si < str.length; si++) {
          lastcode = code;
          code = str.charCodeAt(si);
          if (code === 'ー'.charCodeAt(0)) {
            if (lastcode < 128) {
              strArray.push('-');
            }
          }
        }
        str = strArray.join('');
      }
      var timeout = 0;
      if (/^@(@|[0-9\.]+)\s+/.test(str)) {
        timeout = RegExp.$1;
        str = str.replace(/^@(@|[0-9\.]+)\s+/, '');
        str = 'if(isnumber(' + str + '),' + str + ',T(N("__@' + timeout + '__"))&' + str + ')';
      }
      return str;
    }
  }
  function updateFormulas(formulaList, range, startw, endw, itemNameList) {
    var sheet = range.getSheet();
    var frozenRows = sheet.getFrozenRows();
    var maxRows = sheet.getMaxRows();
    var dollerHeight = maxRows - frozenRows;
    var dollerRow = frozenRows + 1;
    var row = range.getRow();
    var valRow = row + 2;
    var valHeight = maxRows - row;
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
      var itemName = itemNameList[wi].toString().trim();
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
          sheet.getRange(row + 1, wi + 1, 1, 1).setFormula('');
          sheet.getRange(4, wi + 1).setFormula('iferror(sparkline(indirect(address(9,column(),4)&":"&address(' + maxRows + ',column(),4))),"")');
        } else if (f.indexOf('import(') >= 0) {
          var filename = f.slice(f.indexOf('(') + 1, -f.length + f.indexOf(')')).trim();
          sheet.getRange(row + 1, wi + 1, 1, 1).setFormula('T(N("__IMPORTRANGE__"))&T(N("' + filename + '"))');
          sheet.getRange(4, wi + 1).setFormula('iferror(sparkline(indirect(address(9,column(),4)&":"&address(' + maxRows + ',column(),4))),"")');
        } else {
          var side = 0;
          var constval = 4;
          var formulaArray = [];
          var convArray = [];
          var sideBadArray = [];
          var orgf = f;
          if (orgf.indexOf('[') > -1) {
            side = 1;
            var form = '';
            var splitArray = orgf.split(']');
            splitArray.pop();
            for (var si = 0; si < constval; si++) {
              var sj = splitArray.length - 1 - si;
              if (sj >= 0) {
                if (sj === 0) {
                  form = splitArray[sj].slice(0, splitArray[sj].indexOf('[')).trim();
                }
                var val = splitArray[sj].slice(splitArray[sj].indexOf('[') + 1);
                formulaArray.unshift(val.trim());
              } else {
                formulaArray.unshift('');
              }
            }
            if (form === '') {
              form = splitArray[0].slice(0, splitArray[0].indexOf('[')).trim();
            }
            formulaArray.unshift(form);
          } else {
            formulaArray.push(orgf);
          }
          for (var fi = 0; fi < formulaArray.length; fi++) {

            f = formulaArray[fi];
            sideBad = 0;
            // side unsupported
            {
              if (f.indexOf('|') > -1) {
                sideBad = 1;

                var splitArray = f.split('|');
                var valueArray = [];
                var condArray = [];
                valueArray.push(splitArray.shift());
                if (splitArray[0].trim().indexOf('{') !== 0) {
                  if (splitArray.length === 1) {
                    condArray.push(splitArray[0]);
                  }
                } else {
                  for (var si = 0; si < splitArray.length; si++) {
                    var detailArray = splitArray[si].split('}');
                    var condtmp = '';
                    var nextvalueflag = 0;
                    var nextvaluetmp = '';
                    detailArray[0] = detailArray[0].trim().slice(1);
                    for (var di = 0; di < detailArray.length; di++) {
                      if (detailArray[di].trim().lastIndexOf('{') === -1) {
                        nextvalueflag++;
                      }
                      if (nextvalueflag === 0) {
                        condtmp += detailArray[di] + '}';
                      } else if (nextvalueflag === 1) {
                        condtmp += detailArray[di];
                      } else {
                        if (di === detailArray.length - 1) {
                          nextvaluetmp += detailArray[di];
                        } else {
                          nextvaluetmp += detailArray[di] + '}';
                        }
                      }
                    }
                    condArray.push(condtmp.trim());
                    valueArray.push(nextvaluetmp);
                  }
                }
                var farray = ['iferror('];
                for (var ci = 0; ci < condArray.length; ci++) {
                  farray.push('iferror(');
                }
                farray.push('(');
                for (var ci = 0; ci < condArray.length; ci++) {
                  farray.push('N("__if__")+if(and(' + condArray[ci] + ',isnumber(' + valueArray[ci] + ')),N("__then__")+' + valueArray[ci] + ' ,1/0)),');
                }
                farray.push('N("__if__")+if(and(1,1/0,1))),"")');
                f = farray.join('');
              }
              if (f.indexOf('`') > -1) {
                sideBad = 1;

                var left = '$4.0+if(""="$4",N("__param___")+len("$2"),0)';
                var addr = 'regexreplace(address(row($1),column($1)-(' + left + '),4),"[0-9]+","")';
                var itemLabel = 'indirect(' + addr + '&":"&' + addr + ')';
                var rep = 'iferror(index(if("$1"="$1",' + itemLabel + ',$1),$4.0+N("__left__")-len("$2")+N("__prev__")-1-($4.0)+len("$2")+row()+N("__formula__")),1/0)';

                f = f.replace(/([^=<>\|`'"\$;,{&\s\+\-\*\/\(]*)(`+)({([0-9--]+)}|([0-9--]*))/g, rep);
              }
            }
            sideBadArray.push(sideBad);
            // +N("__formula__")),1/0)

            if (f.indexOf('\'') > -1) {
              var prev = 'if(""="$4",N("__param___")+len("$2"),1)';
              var collabel = getColumnLabel(wi + 1);
              var itemLabel = collabel + ':' + collabel;
              // =iferror(index(電卓A,-1+N("__prev__")+row()+N("__formula__")),1/0) 
              // =iferror(index(電卓,-2+N("__prev__")+row()+N("__formula__")),1/0)
              var rep = 'iferror(index(if("$1"="",' + itemLabel + ',$1),-$4.0+N("__prev__")-(' + prev + ')+row()+N("__formula__")),1/0)';

              f = f.replace(/([^=<>\|`'"\$;,{&\s\+\-\*\/\(]*)('+)({([0-9]+)}|([0-9]*))/g, rep);
            }
            if (f.indexOf('pack') > -1) {
              var target = 'offset($1,' + frozenRows + '+N("__pack__"),0,' + (maxRows - frozenRows) + ',1)';
              var subtarget = target.replace('+N("__pack__")', '');
              var rep = 'iferror(index(filter(' + target + ',' + subtarget + '<>""),if(row()-' + (frozenRows) + '>0,row()-' + (frozenRows) + ',-1)+N("__formula__")),1/0)';
              f = f.replace(/pack\s*\(\s*([^\s\)]+)\s*\)/g, rep);
            }
            if (f.indexOf('subseq') > -1) {
              var target = 'offset($1,' + frozenRows + '+N("__subseq__"),0,' + (maxRows - frozenRows) + ',1)';
              var subtarget = target.replace('+N("__subseq__")', '');
              var start = 'match(index(filter(' + subtarget + ',' + subtarget + '<>""),1,1),' + subtarget + ',0)';
              var end = 'match("_",arrayformula(if(offset(' + subtarget + ',' + start + '-1,0)="","_",offset(' + subtarget + ',' + start + '-1,0))),0)';
              var rep = 'iferror(index(offset(' + target + ',' + start + '-1,0,' + end + ',1),if(row()-' + (frozenRows) + '>0,row()-' + (frozenRows) + ',-1)+N("__formula__")),1/0)';
              f = f.replace(/subseq\s*\(\s*([^\s\)]+)\s*\)/g, rep);
            }
            if (f.indexOf('$') > -1) {
              var collabel = getColumnLabel(wi + 1);
              var itemLabel = collabel + ':' + collabel;

              //var rep = 'iferror(index(if("$1"="",' + itemLabel + ',$1),-$3.0+N("__argv__")+N("__prev__")-1+' + (frozenRows + 1) + '+N("__formula__")),1/0)';
              var rep = 'iferror(index(if("$1"="",' + itemLabel + ',$1),-$3.0+N("__argv__")+if(index(if("$1"="",' + itemLabel + ',$1),-$3.0-1+' + (frozenRows + 1) + ')="",1/0)+N("__prev__")-1+' + (frozenRows + 1) + '+N("__formula__")),1/0)';

              f = f.replace(/([^=><\|`'"\$;,{&\s\+\-\*\/\(]*)(\$+)([0-9]+)/g, rep);
              if (f.indexOf('${') > -1) {
                f = f.replace(/([^=><\|`'"\$;,{&\s\+\-\*\/\(]*)(\$+){([0-9]+)}/g, rep);
              }
            }
            if (f.indexOf('head') > -1) {
              var rep = 'iferror(index($1,$2+N("__head__")+N("$1")-1+if("$2"="",1,0)+' + (frozenRows + 1) + '+N("__formula__")),1/0)';
              f = f.replace(/head\s*\(\s*([^\s\),]+)\s*,*((-|\+)*[0-9]*)\s*\)/g, rep);
              if (itemName.length > 0) {
                var headname = 'N("__head__")+N("' + itemName + '")';
                if (f.indexOf(headname) > -1) {
                  f = f.replace(headname, 'N("__head__")+N("__prev__")');
                }
              }
            }
            if (f.indexOf('last') > -1) {
              var rep = 'iferror(index($1,-$3.0+N("__last__")+1-if("$3"="",1,0)+' + (maxRows) + '+N("__formula__")),1/0)';
              f = f.replace(/last\s*\(\s*([^\s\),]+)\s*,*(-|\+)*([0-9]*)\s*\)/g, rep);
            }
            if (f.trim() === '') {
              f = 'T("__EMPTY__")';
            }
            convArray.push(f);
          }
          var _row = dollerRow;
          var _col = wi + 1;
          var _height = dollerHeight;
          var _width = 1;
          // set to protocode
          if (!side) {
            sheet.getRange(row + 1, _col).setFormula('iferror(T(N(to_text(' + f + '))),"")');
          } else {
            var _widthcount = 1;
            // _col is +1 rigth neighbor
            for (var ii = _col; ii < itemNameList.length; ii++) {
              if (itemNameList[ii].trim() !== "") {
                break;
              }
              _widthcount++;
            }
            var store = '+N("__SIDE__")+' + convArray.join('+N("__SIDE__")+') + '+N("__SIDE__")+0';
            sheet.getRange(row + 1, _col).setFormula('iferror(T(N(to_text(' + store + '))),"")');
            for (var ci = 1; ci <= constval; ci++) {
              if (ci >= formulaArray.length) {
                sheet.getRange(row + 3 + (ci - 1), _col + 1, 1, _widthcount).setValue('');
              } else {
                orgf = formulaArray[ci];
                f = convArray[ci];
                var wval = 1;
                var vval = '';
                if (sideBadArray[ci] === 0) {
                  if (f.indexOf('row()') > -1) {
                    f = f.replace(/row\(\)/g, 'column()+(' + (_row - _col) + ')');
                    if (orgf.indexOf(')') === orgf.length - 1) {
                      if (
                        orgf.indexOf('pack(') > -1 ||
                        orgf.indexOf('subseq(') > -1) {
                        vval = f;
                      }
                    } else if (orgf.indexOf('\'') === orgf.length - 1) {
                      if (orgf.indexOf(' ') === -1 && orgf.indexOf('/') === -1
                        && orgf.indexOf('+') === -1 && orgf.indexOf('-') === -1
                        && orgf.indexOf('*') === -1) {
                        vval = f;
                      }
                    }
                  } else if (orgf.indexOf('last(') > -1 || orgf.indexOf('head(') > -1) {
                    if (orgf.indexOf(')') === orgf.length - 1) {
                      vval = f;
                    }
                  } else {
                    if (orgf.indexOf('\(') === 0 && orgf.indexOf('\)') === orgf.length - 1) {
                      orgf = orgf.slice(1, -1);
                    }
                    if (orgf === f) {

                      vval = f;
                      var startcode = '('.charCodeAt(0);
                      var endcode = '9'.charCodeAt(0);
                      for (var oi = 0; oi < f.length; oi++) {
                        var code = f.charCodeAt(oi);
                        if (!(startcode <= code && code <= endcode)) {
                          vval = '';
                          break;
                        }
                      }
                      if (vval === '') {
                        if (orgf.indexOf(' ') === -1 && orgf.indexOf('/') === -1
                          && orgf.indexOf('+') === -1 && orgf.indexOf('-') === -1
                          && orgf.indexOf('*') === -1) {
                          vval = 'index(' + f + ',column()+(' + (_row - _col) + '))';
                        }
                      }
                    }
                  }
                }
              }
              sheet.getRange(row + 3 + (ci - 1), _col, 1, _widthcount).setFormula(vval);
            }
          }
          f = convArray[0];
          if (f.indexOf('N("__prev__")') > -1) {
            // remove errors on initals
            var errors = sheet.getRange(row + 3, _col, frozenRows + 1 - (row + 3), 1).getValues();
            var corrects = [];
            var hasError = false;
            for (var ei = 0; ei < errors.length; ei++) {
              if (/^#.*!$/.test(errors[ei][0])) {
                corrects.push(['']);
                hasError = true;
              } else {
                corrects.push([errors[ei][0]])
              }
            }
            if (hasError) {
              sheet.getRange(row + 3, _col, frozenRows + 1 - (row + 3), 1).setValues(corrects);
            }
          } else {
            var ff = 'iferror(if(' + f + '="","",+' + f + '),"")';
            sheet.getRange(row + 3, _col, frozenRows + 1 - (row + 3), _width).setFormula(ff);
          }
          if (f.indexOf('__@') > -1) {
            sheet.getRange(_row, _col, _height, _width).setFormula('');
            var therange = sheet.getRange(_row, _col);
            therange.setFormula(f);
            var val = therange.getValue();
            var time = (new Date()).getTime();
            therange.setFormula('"' + val + '"&T(N("__#' + time + '__"))');

          } else {
            var ff = 'iferror(' + f + ',"")';
            sheet.getRange(_row, _col, _height, _width).setFormula(ff);
          }

        }
      }
    }
  }
}

function atprocess(spread, byhand) {
  var sheet = spread.getActiveSheet();
  var sheetName = sheet.getName();
  var range = spread.getRangeByName("'" + sheetName + "'!__formulaList__");
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
  var formulaRange = spread.getRangeByName("'" + sheetName + "'!__formulaList__");
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
function draw(spread) {
  var formatListRange = spread.getRangeByName('__formulaList__');
  var sheet = formatListRange.getSheet();
  var formulaValues = spread.getRangeByName('__formulaList__').getValues()[0];
  var startColumn = 0;
  var endColumn = sheet.getMaxColumns();

  for (var fi = 0; fi < formulaValues.length; fi++) {
    if (formulaValues[fi].toString().indexOf('$3+1-mod') !== -1) {
      if (startColumn === 0) {
        startColumn = fi + 1;
      }
    } else {
      if (startColumn !== 0) {
        endColumn = fi;
        break;
      }
    }
  }
  if (startColumn === 0) {
    return;
  }
  if ((endColumn - startColumn + 1) <= 0) {
    return;
  }
  var frozenRows = sheet.getFrozenRows();
  var maxRows = sheet.getMaxRows();
  var ymax = maxRows - frozenRows;
  var range = sheet.getRange(frozenRows + 1, startColumn, ymax, (endColumn - startColumn + 1));
  var values = range.getValues();
  var rects = [];
  var grid = 7;
  var end = 0;
  var valcount = [];
  for (var vi = 0; vi < values[0].length; vi++) {
    valcount[vi] = 0;
    for (var vj = 0; vj < values.length; vj++) {
      if (values[vj][vi]) {
        valcount[vi]++;
        if (valcount[vi] > 1) {
          break;
        }
      }
    }
    if (valcount[vi] === 0) {
      end = vi;
      break;
    }
  }
  if (end === 0) {
    return;
  }
  for (var vj = 0; vj < values.length; vj++) {
    values[vj][end] = values[vj][0];
  }
  for (var vi = 0; vi < end; vi++) {
    for (var vj = 0; vj < values.length; vj++) {
      if (values[vj][vi]) {
        var val = Math.round(parseFloat(values[vj][vi].toString())); // this      
        var val1;
        if (vi - 1 >= 0) { val1 = values[vj][vi - 1]; } //prev
        var val2;
        if (vi + 1 <= end) { val2 = values[vj][vi + 1]; } //next
        var val3;
        if (vj - 1 >= 0) { val3 = values[vj - 1][vi]; } // up
        var val4;
        if (vj + 1 < values.length) { val4 = values[vj + 1][vi]; } // down
        var val5;
        if (vj - 1 >= 0 && vi - 1 >= 0) { val5 = values[vj - 1][vi - 1]; } // up prev 5,9
        var val6;
        if (vj + 1 < values.length && vi + 1 <= end) { val6 = values[vj + 1][vi + 1]; } // down next 6,10,
        var val7;
        if (vj - 1 >= 0 && vi + 1 <= end) { val7 = values[vj - 1][vi + 1]; } // up next 7,11
        var val8;
        if (vj + 1 < values.length && vi - 1 >= 0) { val8 = values[vj + 1][vi - 1]; } // down prev  8,12
        var val9;
        if (vj - 2 >= 0 && vi - 1 >= 0) { val9 = values[vj - 2][vi - 1]; } // upup prev 
        var val10;
        if (vj + 2 < values.length && vi + 1 <= end) { val10 = values[vj + 2][vi + 1]; } // downdown next
        var val11;
        if (vj - 2 >= 0 && vi + 1 <= end) { val11 = values[vj - 2][vi + 1]; } // upup next
        var val12;
        if (vj + 2 < values.length && vi - 1 >= 0) { val12 = values[vj + 2][vi - 1]; } // downdown prev
        if (valcount[vi] === 1) {
          if (val2) {
            var start;
            var max;
            val2 = Math.round(parseFloat(val2.toString()));
            if (val2 < val) {
              start = val2;
              if (val1) {
                val1 = Math.round(parseFloat(val1.toString()));
                if (val1 < val) {
                  max = val;
                } else {
                  max = val + 1;
                }
              } else {
                max = val + 1;
              }
            } else {
              max = val2 + 1;
              if (val1) {
                val1 = Math.round(parseFloat(val1.toString()));
                if (val1 < val) {
                  start = val;
                } else {
                  start = val + 1;
                }
              } else {
                start = val + 1;
              }
            }
            for (var vk = start; vk < max; vk++) {
              rects.push('<rect opacity="1" fill="#ff0000" x="' + (vk * grid) + '" y="' + ((vj + 1) * grid) + '" width="' + (grid) + '" height="' + (grid) + '"/>');
            }
          }
        } else {
          var skip = false;
          if (valcount[vi + 1] === 1 && val4 && !val3 && val6) {
            skip = true;
          } else if (valcount[vi - 1] === 1 && val3 && !val4 && val5) {
            skip = true;
          } else if (val2 && !val3) {
            val2 = Math.round(parseFloat(val2.toString()));
            if (val >= val2) {
              skip = true;
            }
          } else if (val1 && !val4) {
            val1 = Math.round(parseFloat(val1.toString()));
            if (val >= val1) {
              skip = true;
            }
          }
          if (!skip) {
            rects.push('<rect opacity="1" fill="#ff0000" x="' + (val * grid) + '" y="' + ((vj + 1) * grid) + '" width="' + (grid) + '" height="' + (grid) + '"/>');
          }
        }
      }
    }
  }
  var stroke;
  var xmax = ymax * 1.85;
  for (var i = 0; i < ymax + 1; i++) {
    stroke = "#aaaaaa";
    if (i % 10 === 0) {
      stroke = "#000000";
    }
    rects.push('<line x1="0" y1="' + (i * grid) + '" x2="' + (xmax * grid) + '" y2="' + (i * grid) + '" stroke="' + stroke + '" stroke-width="1"/>');
  }
  for (var j = 0; j < xmax + 1; j++) {
    stroke = "#aaaaaa";
    if (j % 10 === 0) {
      stroke = "#000000";
    }
    rects.push('<line x1="' + (j * grid) + '" y1="0" x2="' + (j * grid) + '" y2="' + (ymax * grid) + '" stroke="' + stroke + '" stroke-width="1"/>');
  }
  var svg = '<svg width="1000" height="1000"><g transform="translate(0,0)"><g transform="scale(1,1)">' + rects.join('') + '</g></g></svg>';
  var htmlOutput = HtmlService
    .createHtmlOutput(svg)
    .setWidth(1050)
    .setHeight(1610);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '方眼紙');
}
function importRange(spread) {
  var namedRanges = spread.getNamedRanges();
  var itemNameRange = spread.getRange('__itemNameList__');
  var itemNameList = itemNameRange.getValues()[0];
  var formulaRange = spread.getRange('__formulaList__');
  var formulaList = formulaRange.getValues()[0];
  var formulaRow = formulaRange.getRow();
  var sheet = formulaRange.getSheet();
  var maxRows = sheet.getMaxRows();
  var nameIdHash = {};
  var importHash = {};
  var importWidth = 0;
  var importSheet = spread.getSheetByName("_import");
  if (!importSheet) {
    importSheet = spread.insertSheet('_import', 2);
    spread.setActiveSheet(sheet);
  } else {
    var importFormulas = importSheet.getRange("1:1").getFormulas()[0];
    for (var ii = 0; ii < importFormulas.length; ii++) {
      if (/importrange\("([^"]+)","([^"]+)"&T\(N\("([^"]+)"\)\)\)/.test(importFormulas[ii])) {
        var fileId = RegExp.$1;
        var itemName = RegExp.$2;
        var filename = RegExp.$3;
        var range;
        itemName = itemName.trim();
        importWidth = ii + 1;
        importHash[itemName] = { fileId: fileId, row: importWidth, filename: filename };
      }
    }
    for (var ni = 0; ni < namedRanges.length; ni++) {
      var rangeName = namedRanges[ni].getName();
      if (rangeName.indexOf('_') === 0 && rangeName.indexOf('__') === -1) {
        var itemName = rangeName.slice(1);
        if (itemName in importHash) {
          importHash[itemName].namedRange = namedRanges[ni];
        }
      }
    }
  }
  {
    var importMax = importSheet.getMaxRows();
    var tobedel = importMax - maxRows;
    if (tobedel > 0) {
      importSheet.deleteRows(maxRows + 1, importMax - maxRows);
    }
  }
  var parentFolder; {
    var rootId = DriveApp.getRootFolder().getId();
    var parents = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();
    while (parents.hasNext()) {
      var candidate = parents.next();
      if (candidate.getId() !== rootId) {
        parentFolder = candidate;
      }
    }
    if (!parentFolder) {
      return;
    }
  }
  for (var fi = 0; fi < formulaList.length; fi++) {
    if (formulaList[fi].toString().trim().indexOf('import(') === 0 &&
      itemNameList[fi].toString().trim().length > 0
    ) {
      if (/import\(([\s\S]+)\)/i.test(formulaList[fi])) {
        var filename = RegExp.$1;
        filename = filename.trim();
        var fileid;
        if (filename in nameIdHash) {
          fileid = nameIdHash[filename];
        } else {
          var fileIter = parentFolder.getFilesByName(filename);
          if (fileIter && fileIter.hasNext()) {
            fileId = fileIter.next().getId();
            nameIdHash[filename] = fileId;
          }
        }
        if (!fileId) {
          Logger.log('parent not found, please open from folder');
          continue;
        }
        var itemName = itemNameList[fi].toString().trim();
        var importData;
        var importFormula = 'importrange("' + fileId + '","' + itemName + '"&T(N("' + filename + '")) )';
        if (itemName in importHash) {
          importData = importHash[itemName];
          if (importData.fileId !== fileId) {
            var range = importHash[itemName].namedRange.getRange();
            var cell = importSheet.getRange(1, range.getWidth());
            cell.setFormula(importFormula);
          }
        } else {
          var cell = importSheet.getRange(1, importWidth + 1);
          cell.setFormula(importFormula);
          var collab = getColumnLabel(importWidth + 1);
          importWidth++;
          var range = importSheet.getRange(collab + ':' + collab);
          spread.setNamedRange('_' + itemName, range);
        }
        sheet.getRange(formulaRow + 1, fi + 1).setFormula('T(N(offset(_' + itemName + ',0,0,1,1)))&T(N("__IMPORTRANGE__"))&T(N("' + filename + '"))');
        sheet.getRange(formulaRow + 3, fi + 1, maxRows - (formulaRow + 3) + 1, 1).setValue('');
        sheet.getRange(formulaRow + 3, fi + 1).setFormula('offset(_' + itemName + ',' + ((formulaRow + 3) - 1) + ',0)');
      }
    }
  }
}
function updateImport(spread, _itemName, _filename) {

  var importSheet = spread.getSheetByName("_import");
  if (!importSheet) {
    return false;
  }

  var importFormulas = importSheet.getRange("1:1").getFormulas()[0];
  var _fileId = '';
  var importWidth = 0;
  for (var ii = 0; ii < importFormulas.length; ii++) {
    if (importFormulas[ii].replace('=').length === 0) {
      break;
    }
    importWidth = ii + 1;
    // importrange("fileid","itemName"&T(N("filename"))
    if (importFormulas[ii].indexOf('importrange') > -1 &&
      importFormulas[ii].indexOf('"&T(N("') > -1
    ) {
      var filename = importFormulas[ii].slice(
        importFormulas[ii].indexOf('"&T(N("') + '"&T(N("'.length,
        -importFormulas[ii].length + importFormulas[ii].indexOf('"))')
      );
      var fileId = importFormulas[ii].slice(
        importFormulas[ii].indexOf('importrange("') + 'importrange("'.length,
        -importFormulas[ii].length + importFormulas[ii].indexOf('","')
      );
      var itemName = importFormulas[ii].slice(
        importFormulas[ii].indexOf('","') + '","'.length,
        -importFormulas[ii].length + importFormulas[ii].indexOf('"&T(N("')
      );
      if (_filename === filename) {
        _fileId = fileId;
      }
      if (_itemName === itemName && _filename === filename) {
        return true;
      }
    }
  }
  if (_fileId === '') {
    return false;
  }
  var importFormula = 'importrange("' + _fileId + '","' + _itemName + '"&T(N("' + _filename + '")) )';
  importSheet.getRange(1, importWidth + 1).setFormula(importFormula);
  var collab = getColumnLabel(importWidth + 1);
  spread.setNamedRange('_' + _itemName, importSheet.getRange(collab + ':' + collab));
  return true;
}
function onOpen(spread, ui, sidebar) {
  if (getObjectType(sidebar)==='Object') {
      return;
  }
  if (!sidebar) {
    ui.createMenu('[れん卓]').addItem('方眼紙を開く', 'draw').addSeparator().addItem('インポート', 'importRange').addItem('リフレッシュ', 'refresh').addToUi();
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
  for (var ni = 0; ni < namedRanges.length; ni++) {
    if (namedRanges[ni].getName().indexOf('__') !== 0) {
      namedRanges[ni].remove();
    }
  }
  var rows = 100 + 8;
  var cols = 26;
  sheet.clearContents();
  var maxRows = sheet.getMaxRows();
  if (maxRows < rows) {
    sheet.insertRows(maxRows, rows - maxRows);
  } else if (maxRows > rows) {
    sheet.deleteRows(rows + 1, maxRows - rows);
  }
  var maxCols = sheet.getMaxColumns();
  if (maxCols < cols) {
    sheet.insertColumns(maxCols, cols - maxCols);
  } else if (maxCols > cols) {
    sheet.deleteColumns(cols + 1, maxCols - cols);
  }
  for (var ci = 0; ci < cols; ci++) {
    sheet.setColumnWidth(ci + 1, 120);
  }
  var formulaRange = spread.getRange("'" + sheetName + "'!__formulaList__");
  formulaRange.setVerticalAlignment('top');
  formulaRange.setHorizontalAlignment('right');
  formulaRange.setWrap(true);
  // TODO
  for (var ri = 0; ri < rows; ri++) {
    var height = 21;
    if (ri === 1) {
      height = 70;
    } else if (ri === 2) {
      height = 2;
    } else if (ri === 3) {
      height = 70;
    }
    sheet.setRowHeight(ri + 1, height);
  }
  sheet.getRange(4, 1, 1, cols).setFormula('iferror(sparkline(indirect(address(9,column(),4)&":"&address(' + maxRows + ',column(),4))),"")');
}

function refresh(spread) {
  var sheet = spread.getActiveSheet();
  var import = spread.getSheetByName('_import');
  if (import) {
    var importrange = import.getRange('1:1');
    var formulas = importrange.getFormulas();
    importrange.setValue('');
    importrange.setFormulas(formulas);
  }
  var sheetName = sheet.getName();
  if (sheetName === '_import') {
    return;
  }
  var formulaRange = spread.getRange("'" + sheetName + "'!__formulaList__");
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  var targetRow = formulaRange.getRow();
  var targetHeight = formulaRange.getHeight();
  var targetColumn = formulaRange.getColumn();
  var targetWidth = formulaRange.getWidth();
  var itemNameListRange = spread.getRangeByName("'" + sheetName + "'!__itemNameList__");
  var formulaListRange = spread.getRangeByName("'" + sheetName + "'!__formulaList__");
  processFormulaList(spread, sheet, targetRow, targetHeight, targetColumn, targetWidth, itemNameListRange, formulaListRange);
  sheet.getRange(4, 1, 1, maxCols).setFormula('iferror(sparkline(indirect(address(9,column(),4)&":"&address(' + maxRows + ',column(),4))),"")');
}