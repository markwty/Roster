function onEdit(e){
  const rg = e.range;
  if (rg.isChecked()){
    if(rg.getSheet().getName() === "schedule"){
      if(rg.getA1Notation() === "AA13")
      {
        merge();
        rg.uncheck();
      }
      else if(rg.getA1Notation() === "AA14")
      {
        unmerge();
        rg.uncheck();
      }
      else if(rg.getA1Notation() === "AA15")
      {
        colour();
        rg.uncheck();
      }
      else if(rg.getA1Notation() === "AA16")
      {
        uncolour();
        rg.uncheck();
      }
      else if(rg.getA1Notation() === "AA17")
      {
        outline();
        rg.uncheck();
      }
      else if(rg.getA1Notation() === "AA18")
      {
        outline_10pm();
        rg.uncheck();
      }
      else if(rg.getA1Notation() === "AA19")
      {
        copy();
        rg.uncheck();
      }
      else if(rg.getA1Notation() === "AA20")
      {
        paste();
        rg.uncheck();
      }
    }
  }
  else if(rg.getSheet().getName() === "schedule" && rg.getA1Notation() === "M4")
  {
    changeTimeTable(rg.getValues()[0][0]);
    rg.clearContent();
  }
}

function RangeIntersect (R1, R2) {
  return (R1.getLastRow() >= R2.getRow()) && (R2.getLastRow() >= R1.getRow()) && (R1.getLastColumn() >= R2.getColumn()) && (R2.getLastColumn() >= R1.getColumn());
}

function Find(data, term){
  return data.findIndex(([item]) => {return item == term}); 
}

function FindDate(data, term){
  return data.findIndex(([item]) => {return item.valueOf() == term}); 
}

function onSelectionChange(e){
  var r, r2, offsetCell;
  var rangestr;
  var dat;
  var monthNum, row, idx;
  var ws;
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  r = e.range;
  if (spreadsheet.getActiveSheet().getName() !== "schedule"){
    return;
  } else if (r.getValues().length > 1 || r.getValues().length == 0){
    return;
  } else if (!RangeIntersect(spreadsheet.getRange("E4:K9"), r)){
    return;
  }
  
  var value = r.getCell(1, 1).getValue();
  if (isNaN(value) || value < 1 || value > 31){
    return;
  }
  var mDate = new Date("1 " + spreadsheet.getRange("A1:D2").getCell(2, 4).getValue() + " 2020");
  monthNum = mDate.getMonth()+1;
  dat = new Date(monthNum + " " + value + " " + spreadsheet.getRange("A1:D1").getCell(1, 4).getValue());
  dat.setDate(dat.getDate() - dat.getDay());// dat - (dat.getDay() % 7);
  //console.log(Date(dat));
  
  ws = SpreadsheetApp.getActive().getSheetByName("records");
  r2 = spreadsheet.getRange("B11");
  rangestr = spreadsheet.getRange("A11").getValue();
  if(r2.getValue() !== ""){
    idx = FindDate(ws.getRange("A:A").getValues(), r2.getValue().valueOf());
    if(idx != -1){
      r = ws.getRange("A" + (idx + 1));
      spreadsheet.getRange(rangestr).copyTo(r.offset(1, 1));
    } else {
      idx = ws.getRange("A:A").getValues().length;//Find(ws.getRange("A:A").getValues(), "/")
      if(idx != -1){
        r = ws.getRange("A" + (idx + 1));
        row = r.row + 19;
      } else {
        row = -18;
      }
      ws.getRange("A:A").getCell(row + 19, 1) = r2.value;
      r = ws.getRange("A:A").getCell(row + 19, 1);
      spreadsheet.getRange(rangestr).copyTo(r.offset(1, 1));
    }
  } else {
      r2.setValue(Utilities.formatDate(dat, "GMT+8", "MM/dd/yyyy"));
      return;
  }
  
  r2.setValue(Utilities.formatDate(dat, "GMT+8", "MM/dd/yyyy"));
  dat = r2.getValue();
  idx = FindDate(ws.getRange("A:A").getValues(), dat.valueOf());
  r2 = spreadsheet.getRange(rangestr);
  if(idx != -1){
      r = ws.getRange("A" + (idx + 1));
      offsetCell = r.offset(1, 1);
      ws.getRange(offsetCell.getA1Notation() + ":" + offsetCell.offset(r2.getNumRows()-1, r2.getNumColumns()-1).getA1Notation()).copyTo(r2);
  } else {
      spreadsheet.getRange(rangestr).clear();
  }
}

function changeTimeTable(value){
  var r, r2, offsetCell;
  var rangestr;
  var dat;
  var monthNum, row, idx;
  var ws;
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (isNaN(value) || value < 1 || value > 31){
    //console.log(spreadsheet.getName());
    return;
  }
  
  var mDate = new Date("1 " + spreadsheet.getRange("A1:D2").getCell(2, 4).getValue() + " 2020");
  monthNum = mDate.getMonth()+1;
  dat = new Date(monthNum + " " + value + " " + spreadsheet.getRange("A1:D1").getCell(1, 4).getValue());
  dat.setDate(dat.getDate() - dat.getDay());// dat - (dat.getDay() % 7);
  //console.log(Date(dat));
  
  ws = SpreadsheetApp.getActive().getSheetByName("records");
  r2 = spreadsheet.getRange("B11");
  rangestr = spreadsheet.getRange("A11").getValue();
  if(r2.getValue() !== ""){
    idx = FindDate(ws.getRange("A:A").getValues(), r2.getValue().valueOf());
    if(idx != -1){
      r = ws.getRange("A" + (idx + 1));
      spreadsheet.getRange(rangestr).copyTo(r.offset(1, 1));
    } else {
      idx = ws.getRange("A:A").getValues().length;//Find(ws.getRange("A:A").getValues(), "/")
      if(idx != -1){
        r = ws.getRange("A" + (idx + 1));
        row = r.row + 19;
      } else {
        row = -18;
      }
      ws.getRange("A:A").getCell(row + 19, 1) = r2.value;
      r = ws.getRange("A:A").getCell(row + 19, 1);
      spreadsheet.getRange(rangestr).copyTo(r.offset(1, 1));
    }
  } else {
      r2.setValue(Utilities.formatDate(dat, "GMT+8", "MM/dd/yyyy"));
      return;
  }
  
  r2.setValue(Utilities.formatDate(dat, "GMT+8", "MM/dd/yyyy"));
  dat = r2.getValue();
  idx = FindDate(ws.getRange("A:A").getValues(), dat.valueOf());
  r2 = spreadsheet.getRange(rangestr);
  if(idx != -1){
      r = ws.getRange("A" + (idx + 1));
      offsetCell = r.offset(1, 1);
      ws.getRange(offsetCell.getA1Notation() + ":" + offsetCell.offset(r2.getNumRows()-1, r2.getNumColumns()-1).getA1Notation()).copyTo(r2);
  } else {
      spreadsheet.getRange(rangestr).clear();
  }
}

function CountMerge(r, v, values) {
  var count;
  var ranges, vals;

  count = 0;
  mcount = 1;

  v = v.toLowerCase();
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName("schedule");
 
  r = spreadsheet.getRange(r);
  values = r.getValues().map(function(row) {
    return row.map(function(value) {
      return TRIM(value).toLowerCase();
    });
  });
  for(var row = 0; row < r.getNumRows(); row++){
    for(var col = 0; col < r.getNumColumns(); col++){
      if(values[row][col] === v){
        count++;
      }
    }
  }
  
  ranges = r.getMergedRanges();
  for(var i = 0; i < ranges.length; i++){
    vals = ranges[i].getValues();
    if(vals[0][0].toLowerCase() === v){
      count = count + vals.length - 1;
    }
  }
  return count;
}

function TRIM(str){
  var start = 0, end = 0;
  var chars = str.split('');

  for(var i = 0; i < chars.length; ++i){
    if(chars[i] != ' '){
      break;
    }
    start++;
  }
  for(var i = chars.length-1; i >= 0; --i){
    if(chars[i] != ' '){
      break;
    }
    end++;
  }
  return str.substring(start, str.length - end);
}

function PROPER_CASE(str) {
  if (typeof str != "string")
    throw 'Expected string but got a ${typeof str} value.';
  
  str = str.toLowerCase();

  var arr = str.split(/.-:?—/ );
  
  return arr.reduce(function(val, current) {
    return val += (current.charAt(0).toUpperCase() + current.slice(1));
  }, "");
}

function merge() {
  var c, r, rr;
  var rangestr, v;
  var rCount = 0;
  var values;

  SpreadsheetApp.getActive().getSheetByName("schedule").activate();
  count = 1;
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  rangestr = spreadsheet.getRange("A11").getValue();
  rr = spreadsheet.getRange(rangestr);
  rr.setValues(rr.getValues().map(function(row) {
    return row.map(function(value) {
      return PROPER_CASE(TRIM(value));
    });
  }));

  values = rr.getValues();
  for(var col = 1; col <= rr.getNumColumns(); col++){
    v = ".................";
    rCount = 0;
    for(var row = 1; row <= rr.getNumRows(); row++){
      if (values[row-1][col-1] === v && v !== ""){
        rCount++;
      } else {
        if(rCount > 0){
          c = rr.getCell(row, col);
          r = spreadsheet.getRange(c.offset(-rCount - 1, 0).getA1Notation() + ":" + c.offset(-1, 0).getA1Notation());
          r.merge();
          r.setHorizontalAlignment("center");
          r.setVerticalAlignment("middle");
          rCount = 0;
        }
        v = values[row-1][col-1];
      }
    }
    if(rCount > 0){
      r.merge();
      r.setHorizontalAlignment("center");
      r.setVerticalAlignment("middle");
    }
  }
  
  SpreadsheetApp.flush();
}

function unmerge() {
  var r;
  var curr, rangestr;
  var count;
   
  SpreadsheetApp.getActive().getSheetByName("schedule").activate();
  count = 1;
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  rangestr = spreadsheet.getRange("A11").getValue();
  r = spreadsheet.getRange(rangestr);
  r.setValues(r.getValues().map(function(row) {
    return row.map(function(value) {
      return PROPER_CASE(TRIM(value));
    });
  }));

  /*
  var values = r.getValues();
  for(var col = 1; col <= r.getNumColumns(); col++){
    for(var row = 1; row <= r.getNumRows(); row++){
      c = r.getCell(row, col);
      if(count > 1){
        count--;
      } else if(c.isPartOfMerge()){
        count = c.getMergedRanges()[0].getValues().length;
        curr = values[row-1][col-1];
        for(var i = row; i < row + count; i ++){
          values[i-1][col-1] = curr;
        }
      }
    }
  }
  */
  var values = r.getValues();
  var ranges = r.getMergedRanges();
  var rvalues, rrow, rcol, row, col;
  rrow = r.getRow();
  rcol = r.getColumn();
  for(var i = 0; i < ranges.length; i++){
    rvalues = ranges[i].getValues();
    count = rvalues.length;
    curr = rvalues[0];
    row = ranges[i].getRow() - rrow;
    col = ranges[i].getColumn() - rcol;
    for(var ii = row; ii < row + count; ii++){
      values[ii][col] = curr;
    }
  }
  r.breakApart();
  r.setValues(values);
  
  SpreadsheetApp.flush();
}

function colour() {
  var colorRange, r;
  var rangestr;
  var dict = {};
  
  SpreadsheetApp.getActive().getSheetByName("schedule").activate();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  colorRange = spreadsheet.getRange("P34");
  for(var idx = 0; idx < 16; idx++){
    if(colorRange.offset(idx, -11).getValue() != undefined){
      dict[PROPER_CASE(TRIM(colorRange.offset(idx, -11).getValue()))] = colorRange.offset(idx, 0).getBackgroundColor();
    }
  }
  
  ignore = 0;
  rangestr = spreadsheet.getRange("A11").getValue();
  r = spreadsheet.getRange(rangestr);
  r.setValues(r.getValues().map(function(row) {
    return row.map(function(value) {
      return PROPER_CASE(TRIM(value));
    });
  }));
  /*
  for(var col = 1; col < r.getNumColumns(); col++){
    for(var row = 1; row < r.getNumRows(); row++){
      console.log("hi2:" + row);
      r.getCell(row, col).setValue(PROPER_CASE(TRIM(r.getCell(row, col).getValue())));
    }
  }
  */
  r.setBackgrounds(r.getValues().map(function(row) {
    return row.map(function(str) {
      var value = PROPER_CASE(TRIM(str));
      if(dict[value] != undefined){
        return dict[value];
      }
    });
  }));
  /*
  for(var col = 1; col < r.getNumColumns(); col++){
    for(var row = 1; row < r.getNumRows(); row++){
      if(ignore > 0) {
        ignore--;
      } else {
          c = r.getCell(row, col);
          value = PROPER_CASE(TRIM(c.getValue()));
          if(dict[value] != undefined){
            c.setBackground(dict[value]);
            if(c.isPartOfMerge()){
              ignore = c.getMergedRanges()[0].getValues().length - 1;
            }
          }
      }
    }
  }
  */
  
  SpreadsheetApp.flush();
}

function uncolour() {
  var rangestr;
  
  SpreadsheetApp.getActive().getSheetByName("schedule").activate();
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  rangestr = spreadsheet.getRange("A11").getValue();
  spreadsheet.getRange(rangestr).setBackground("#ffffff");

  SpreadsheetApp.flush();
}

function outline() {
  var rangestr;
  var firstCol, count, divisor, step;
  
  SpreadsheetApp.getActive().getSheetByName("schedule").activate();
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  rangestr = spreadsheet.getRange("A11").getValue();

  count = 0;
  firstCol = spreadsheet.getRange(rangestr).getColumn();
  for(var col = 0; col < spreadsheet.getRange(rangestr).getNumColumns(); col++){
    if (spreadsheet.isColumnHiddenByUser(col + firstCol)){
      count++;
    }
  }
  
  r = spreadsheet.getRange(rangestr);
  divisor = count / 7;
  if(count % 7 == 0){
    spreadsheet.getRange(r.getCell(1,1).getA1Notation() + ":" + r.getCell(1,r.getNumColumns()).getA1Notation()).setBorder(true, null, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    spreadsheet.getRange(r.getCell(r.getNumRows(),1).getA1Notation() + ":" + r.getCell(r.getNumRows(),r.getNumColumns()).getA1Notation()).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    step = 0;
    for(var col = 0; col < r.getNumColumns(); col++){
      if((step % 3) % (3 - divisor) == 0){
        spreadsheet.getRange(r.getCell(1,col+1).getA1Notation() + ":" + r.getCell(r.getNumRows(),col+1).getA1Notation()).setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
      step++;
    }
    
    step = 1;
    for(var col = 0; col < r.getNumColumns(); col++){
      if(step == r.getNumColumns() - divisor){
        spreadsheet.getRange(r.getCell(1,col+1).getA1Notation() + ":" + r.getCell(r.getNumRows(),col+1).getA1Notation()).setBorder(null, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
      step++;
    }
  }
   
  SpreadsheetApp.flush();
}

function outline_10pm() {
  var rangestr;
  var firstCol, count, divisor, step, offset;
  
  SpreadsheetApp.getActive().getSheetByName("schedule").activate();
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  rangestr = spreadsheet.getRange("A11").getValue();

  count = 0;
  firstCol = spreadsheet.getRange(rangestr).getColumn();
  for(var col = 0; col < spreadsheet.getRange(rangestr).getNumColumns(); col++){
    if (spreadsheet.isColumnHiddenByUser(col + firstCol)){
      count++;
    }
  }
  
  r = spreadsheet.getRange(rangestr);
  offset = spreadsheet.getRange("A10").getValue();
  if(offset === ""){
    offset = 0;
  }
  divisor = count / 7;
  if(count % 7 == 0){
    spreadsheet.getRange(r.getCell(1,1).getA1Notation() + ":" + r.getCell(1,r.getNumColumns()).getA1Notation()).setBorder(true, null, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    spreadsheet.getRange(r.getCell(r.getNumRows()-offset,1).getA1Notation() + ":" + r.getCell(r.getNumRows()-offset,r.getNumColumns()).getA1Notation()).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    step = 0;
    for(var col = 0; col < r.getNumColumns(); col++){
      if((step % 3) % (3 - divisor) == 0){
        spreadsheet.getRange(r.getCell(1,col+1).getA1Notation() + ":" + r.getCell(r.getNumRows(),col+1).getA1Notation()).setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
      step++;
    }
    
    step = 1;
    for(var col = 0; col < r.getNumColumns(); col++){
      if(step == r.getNumColumns() - divisor){
        spreadsheet.getRange(r.getCell(1,col+1).getA1Notation() + ":" + r.getCell(r.getNumRows(),col+1).getA1Notation()).setBorder(null, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
      step++;
    }
  }
   
  SpreadsheetApp.flush();
}

function copy() {
  var rangestr;
  var ws;
  
  SpreadsheetApp.getActive().getSheetByName("schedule").activate();
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  rangestr = spreadsheet.getRange("A11").getValue();
  
  ws = SpreadsheetApp.getActive().getSheetByName("spare");
  r = spreadsheet.getRange(rangestr);
  ws.getRange(rangestr).clearFormat();
  ws.getRange(rangestr).breakApart();
  r.copyTo(ws.getRange(rangestr));
   
  SpreadsheetApp.flush();
}

function paste() {
  var rangestr;
  var ws;
  
  SpreadsheetApp.getActive().getSheetByName("schedule").activate();
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  rangestr = spreadsheet.getRange("A11").getValue();
  
  ws = SpreadsheetApp.getActive().getSheetByName("spare");
  r = spreadsheet.getRange(rangestr);
  spreadsheet.getRange(rangestr).clearFormat();
  spreadsheet.getRange(rangestr).breakApart();
  ws.getRange(rangestr).copyTo(r);
  
  SpreadsheetApp.flush();
}

