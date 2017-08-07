var fs = require('fs');
var path = require('path');
var xlsx2json = require('node-xlsx2json');

Array.prototype.contains = function(obj) {
  var i = this.length;
  while (i--) {
    if (this[i] === obj) {
      return true;
    }
  }
};

var Excel = { };

//Gets a column letter from a number, e.g. 1 = A, 2 = B, 27 = AA, etc.
Excel.colname = function(column) {
  var columnString = "";
  var columnNumber = column;
  while (columnNumber > 0) {
    var currentLetterNumber = (columnNumber - 1) % 26;
    var currentLetter = String.fromCharCode(currentLetterNumber + 65);
    columnString = currentLetter + columnString;
    columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
  }
  return columnString;
}

//Gets a column number from a letter, e.g. A = 1, B = 2, AA = 27, etc.
Excel.colnum = function(column) {
  var retVal = 0;
  var col = column.toUpperCase();
  for (var iChar = col.length - 1; iChar >= 0; iChar--) {
    var colPiece = col[iChar];
    var colNum = colPiece.charCodeAt(0) - 64;
    retVal = retVal + colNum * Math.floor(Math.pow(26, col.length - (iChar + 1)));
  }
  return retVal;
}

Excel.name = function(pos) {
  if (arguments.length == 0) throw new Error("Formula error");
  return Excel.colname(pos.column) + pos.row;
};

Excel.pos = function(name) {
  if (arguments.length == 0) throw new Error("Formula error");
  var m = /^([A-Z]+)([0-9]+)$/.exec(name);
  if (!m) throw new Error("Argument");
  var column = Excel.colnum(m[1]);
  var row = parseInt(m[2]);
  return { row: row, column: column };
};

Excel.range = function(range) {
  var parts = range.split(":");
  var start = Excel.pos(parts[0]);
  var stop = Excel.pos(parts[1]);
  
  var result = [ ];
  for (var r = Math.min(start.row, stop.row); r <= Math.max(start.row, stop.row); r++) {
    for (var c = Math.min(start.column, stop.column); c <= Math.max(start.column, stop.column); c++) {
      var cell = { row: r, column: c };
      cell.name = Excel.name(cell);
      result.push(cell);
    }
  }
  return result;
};

function getRange(array) {
  var rmin, rmax, cmin, cmax;
  for (var i = 0; i < array.length; i++) {
    var pos = Excel.pos(array[i].cell);
    if (i == 0) {
      rmin = rmax = pos.row;
      cmin = cmax = pos.column;
    }
    else {
      rmin = Math.min(rmin, pos.row);
      rmax = Math.max(rmax, pos.row);
      cmin = Math.min(cmin, pos.column);
      cmax = Math.max(cmax, pos.column);
    }
  }
  return Excel.name({ row: rmin, column: cmin }) + ':' + Excel.name({ row: rmax, column: cmax });
}

/** Converts XLSX file to CSV object */
function xlsx2csv(file, options, callback) {
  var table = [ ];
  
  //Shift arguments
  if (typeof options == "function" && typeof callback == "undefined") {
    callback = options;
    options = null;
  }
  
  if (typeof callback != "function") {
    callback = function(error, result) { }; //Dummy
  }
  
  //Default options
  options = options || { };
 
  try {
    if (typeof file != "string" || file.length == 0) {
      throw new Error("Please check the 'file' argument!");
    }
    
    if (!options.sheet) {
      throw new Error("Missing worksheet name! Hint: options.sheet")
    }
    
    if (!options.separator) {
      options.separator = ',';
    }
    
    if (!options.lineEnd) {
      options.lineEnd = '\n';
    }
   
    //Converts XLSX to JSON
    xlsx2json(file, options, function(error, result) {
      if (error) return callback(error);
      
      try {
        for (var i = 0; i < result.worksheets.length; i++) {
          var worksheet = result.worksheets[i];
          if (options.sheet != worksheet.name) {
            continue;
          }
        
          var range = options.range || worksheet.range || getRange(worksheet.data);
          var parts = range.split(':');
          var start = Excel.pos(parts[0]);
          var end = Excel.pos(parts[1]);
          
          table = [ ];
          for (var r = start.row; r <= end.row; r++) {
            var row = [ ];
            for (var c = start.column; c <= end.column; c++) {
              row.push(null);
            }
            table.push(row);
          }
          
          if (options.verbose) {
            console.log("Capturing " + (end.row - start.row + 1) + " rows and " + (end.column - start.column + 1) + " columns: " + range);
          }
        
          for (var j = 0; j < worksheet.data.length; j++) {
            var item = worksheet.data[j];
            var pos = Excel.pos(item.cell);
            
            if (start.row <= pos.row && pos.row <= end.row && start.column <= pos.column && pos.column <= end.column) {     
              if (item.formula || item.value) {
                if (options.data == "formula" && item.formula) {
                  table[pos.row - start.row][pos.column - start.column] = item.formula; 
                }
                else if (item.value) {
                  table[pos.row - start.row][pos.column - start.column] = options.data == "display" && typeof item.display != "undefined" ? item.display : item.value;
                }
              }
            }
          }
        }
        
        //Escape string value
        for (var i = 0; i < table.length; i++) {
          for (var j = 0; j < table[i].length; j++) {
            var item = table[i][j];
            if (typeof item == "string" && (item.indexOf(options.separator) >= 0 || item.indexOf('"') >= 0)) {
              table[i][j] = '"' + item.replace(/\"/g, '\\"') + '"'; //Quote and escape double quotes
            }
          }
        }
        
        //Build CSV string
        var csv = "";
        for (var i = 0; i < table.length; i++) {
          csv += table[i].join(options.separator) + options.lineEnd;
        }
        
        callback(null, csv);
      }
      catch (e) {
        callback(e);
      }
    });
  }
  catch (e) {
    callback(e);
  }
};

module.exports = xlsx2csv;