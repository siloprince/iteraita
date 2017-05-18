// dummy
function info() {
}
function onEdit(ev) {
  var targetRange = ev.range;
  var sheet = targetRange.getSheet();
  var sheetName = sheet.getName();
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
    recoverFromFormulas(formulaListRange);
  }
  return;

  function recoverFromFormulas(range) {

    var sheet = range.getSheet();
    var frozenRows = sheet.getFrozenRows();
    var valRow = range.getRow() + 1;
    var width = range.getWidth();
    var formulas = sheet.getRange(valRow, 1, 1, width).getFormulas()[0];
    var tn_header = 'iferror(T(N(to_text(';
    var tn_footer = '))),"")';
    for (var fi = 0; fi < formulas.length; fi++) {
      var raw = formulas[fi];
      if (raw.length === 0) {
        continue;
      }
      if (raw.indexOf('=') === 0) {
        formulas[fi] = raw.slice(1);
      }
      if (formulas[fi].indexOf('iferror(T(N(to_text(') === 0) {
        formulas[fi] = formulas[fi].slice(tn_header.length, -tn_footer.length);
      }
      if (formulas[fi].indexOf('+N("__formula__")),"")') > -1) {
        var parts = formulas[fi].split('+N("__formula__")),"")');

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
              if (/if\(""="([0-9--]*)",N\("__param___"\)\+len\("([^"]+)"\)/.test(item)) {
                p1 = RegExp.$1;
                p2 = RegExp.$2;
                if (p1==='') {
                  dcount = p2.length;
                } else {
                  dcount = parseInt(p1,10);
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
                if (p1==='') {
                  dcount = p2.length;
                } else {
                  dcount = parseInt(p1,10);
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
      // 
      // =T("__EMPTY__")+N("__SIDE__")+T("__EMPTY__")+N("__SIDE__")+22+N("__SIDE__")+3
      var sidesep = '+N("__SIDE__")+';
      if (formulas[fi].indexOf(sidesep) > -1) {
        var farray = [];
        var parts = formulas[fi].split(sidesep);
        for (var pi = 0; pi < parts.length; pi++) {
          if (parts[pi] !== 'T("__EMPTY__")') {
            farray.push('[' + parts[pi] + ']');
          }
        }
        formulas[fi] = farray.join('\n');
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
      var conved = convertFormula(rawFormulaList[ri].toString().trim(), valRow);
      formulaList.push(conved);
    }
    var startw = targetColumn;
    var endw = startw + targetWidth - 1;
    updateFormulas(formulaList, formulaListRange, startw, endw, itemNameList);
  }
  return;

  function convertFormula(str, valRow) {
    if (str.length === 0) {
      return '';
    } else {
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
          sheet.getRange(row + 1, wi + 1, 1, 1).setFormula(f);
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
            var splitArray = orgf.split(']');
            splitArray.pop();
            for (var si = 0; si < constval; si++) {
              var sj = splitArray.length - 1 - si;
              if (sj >= 0) {
                var val = splitArray[sj].slice(splitArray[sj].indexOf('[') + 1);
                formulaArray.unshift(val.trim());
              } else {
                formulaArray.unshift('');
              }
            }
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
                    for (var di = 0; di < detailArray.length; di++) {
                      if (detailArray[di].trim().lastIndexOf('{') === -1) {
                        nextvalueflag++;
                      }
                      if (nextvalueflag <= 1) {
                        condtmp += detailArray[di] + '}';
                      } else {
                        nextvaluetmp += detailArray[di] + '}';
                      }
                    }
                    condArray.push(condtmp.trim().slice(1, -1));
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
                var rep = 'iferror(index(if("$1"="$1",' + itemLabel + ',$1),$4.0+N("__left__")-len("$2")+N("__prev__")-1-($4.0)+len("$2")+row()+N("__formula__")),"")';

                f = f.replace(/([^=\|`'"\$;,{&\s\+\-\*\/\(]*)(`+)({([0-9--]+)}|([0-9--]*))/g, rep);
              }
            }
            sideBadArray.push(sideBad);
            // +N("__formula__")),"")
            if (f.indexOf('\'') > -1) {
              var prev = 'if(""="$4",N("__param___")+len("$2"),1)';
              var collabel = getColumnLabel(wi + 1);
              var itemLabel = collabel + ':' + collabel;
              // =iferror(index(電卓A,-1+N("__prev__")+row()+N("__formula__")),"") 
              // =iferror(index(電卓,-2+N("__prev__")+row()+N("__formula__")),"")
              var rep = 'iferror(index(if("$1"="",' + itemLabel + ',$1),-$4.0+N("__prev__")-(' + prev + ')+row()+N("__formula__")),"")';

              f = f.replace(/([^=\|`'"\$;,{&\s\+\-\*\/\(]*)('+)({([0-9]+)}|([0-9]*))/g, rep);
            }
            if (f.indexOf('pack') > -1) {
              var target = 'offset($1,' + frozenRows + '+N("__pack__"),0,' + (maxRows - frozenRows) + ',1)';
              var subtarget = target.replace('+N("__pack__")', '');
              var rep = 'iferror(index(filter(' + target + ',' + subtarget + '<>""),if(row()-' + (frozenRows) + '>0,row()-' + (frozenRows) + ',-1)+N("__formula__")),"")';
              f = f.replace(/pack\s*\(\s*([^\s\)]+)\s*\)/g, rep);
            }
            if (f.indexOf('subseq') > -1) {
              var target = 'offset($1,' + frozenRows + '+N("__subseq__"),0,' + (maxRows - frozenRows) + ',1)';
              var subtarget = target.replace('+N("__subseq__")', '');
              var start = 'match(index(filter(' + subtarget + ',' + subtarget + '<>""),1,1),' + subtarget + ',0)';
              var end = 'match("_",arrayformula(if(offset(' + subtarget + ',' + start + '-1,0)="","_",offset(' + subtarget + ',' + start + '-1,0))),0)';
              var rep = 'iferror(index(offset(' + target + ',' + start + '-1,0,' + end + ',1),if(row()-' + (frozenRows) + '>0,row()-' + (frozenRows) + ',-1)+N("__formula__")),"")';
              f = f.replace(/subseq\s*\(\s*([^\s\)]+)\s*\)/g, rep);
            }
            if (f.indexOf('$') > -1) {
              var collabel = getColumnLabel(wi + 1);
              var itemLabel = collabel + ':' + collabel;

              var rep = 'iferror(index(if("$1"="",' + itemLabel + ',$1),-$3.0+N("__argv__")+N("__prev__")-1+' + (frozenRows + 1) + '+N("__formula__")),"")';

              f = f.replace(/([^=\|`'"\$;,{&\s\+\-\*\/\(]*)(\$+)([0-9]+)/g, rep);
              if (f.indexOf('${') > -1) {
                f = f.replace(/([^=\|`'"\$;,{&\s\+\-\*\/\(]*)(\$+){([0-9]+)}/g, rep);
              }
            }
            if (f.indexOf('head') > -1) {
              var rep = 'iferror(index($1,$2+N("__head__")+N("$1")-1+if("$2"="",1,0)+' + (frozenRows + 1) + '+N("__formula__")),"")';
              f = f.replace(/head\s*\(\s*([^\s\),]+)\s*,*((-|\+)*[0-9]*)\s*\)/g, rep);
              if (itemName.length > 0) {
                var headname = 'N("__head__")+N("' + itemName + '")';
                if (f.indexOf(headname) > -1) {
                  f = f.replace(headname, 'N("__head__")+N("__prev__")');
                }
              }
            }
            if (f.indexOf('last') > -1) {
              var rep = 'iferror(index($1,-$3.0+N("__last__")+1-if("$3"="",1,0)+' + (maxRows) + '+N("__formula__")),"")';
              f = f.replace(/last\s*\(\s*([^\s\),]+)\s*,*(-|\+)*([0-9]*)\s*\)/g, rep);
            }
            if (f === '') {
              f = 'T("__EMPTY__")';
            }
            convArray.push(f);
          }
          var _row = dollerRow;
          var _col = wi + 1;
          var _height = dollerHeight;
          var _width = 1;
          // set to protocode
          if (side) {
            var _widthcount = 1;
            for (var ii = _col + 1; ii < itemNameList.length; ii++) {
              if (itemNameList[ii].trim() !== "") {
                break;
              }
              _widthcount++;
            }
            var store = convArray.join('+N("__SIDE__")+');
            sheet.getRange(row + 1, _col).setFormula('iferror(T(N(to_text(' + store + '))),"")');
            for (var ci = 0; ci < constval; ci++) {
              if (ci >= formulaArray.length) {
                sheet.getRange(row + 3 + ci, _col + 1, 1, _widthcount).setValue('');
              } else {
                orgf = formulaArray[ci];
                f = convArray[ci];
                var wval = 1;
                var vval = '';
                if (sideBadArray[ci] === 0) {
                  if (f.indexOf('row()') > -1) {
                    f = f.replace(/row\(\)/g, 'column()+(' + (_row - _col - 1) + ')');
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
                          vval = 'index(' + f + ',column()+(' + (_row - _col - 1) + '))';
                        }
                      }
                    }
                  }
                }
              }
              sheet.getRange(row + 3 + ci, _col + 1, 1, _widthcount).setFormula(vval);
            }
          } else {
            f = convArray[0];
            sheet.getRange(row + 1, _col).setFormula('iferror(T(N(to_text(' + f + '))),"")');
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
              sheet.getRange(_row, _col, _height, _width).setFormula(f);
            }
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
function onOpen(spread, ui, sidebar) {
  if (!sidebar) {
    ui.createMenu('Iteraita').addItem('サイドバーを開く', 'openSidebar').addItem('リフレッシュ', 'refresh').addItem('リセット', 'reset').addToUi();
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
  var sheetName = sheet.getName();
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