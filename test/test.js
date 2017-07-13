var fs = require('fs');
var path = require('path');
var xlsx2csv = require('../lib/index.js');
  
exports["Worksheet Calculate"] = function(test) {
  var file = path.resolve(__dirname, 'fixture.xlsx');
  xlsx2csv(file, { verbose: false, sheet: "Calculate", data: "formula", separator: ";" }, function(error, result) {
    if (error) {
      test.ok(false, error.message);
    }
    else {
      test.ok(typeof result == "string", "Result should be a string");
      test.ok(result.indexOf('=') >= 0, "Should contain at least one formula");
      test.ok(result.indexOf('1') == 0, "Should start with value '1'");
      test.ok(result.replace(/[^;]+/g, "").length == 6, "Should found exactly 6 separator chars");
      //console.log(result);
    }
    test.done();
  });
};

exports["Worksheet Variables"] = function(test) {
  var file = path.resolve(__dirname, 'fixture.xlsx');
  xlsx2csv(file, { verbose: false, sheet: "Variables", data: "display" }, function(error, result) {
    if (error) {
      test.ok(false, error.message);
    }
    else {
      test.ok(typeof result == "string", "Result should be a string");
      test.ok(result.indexOf('\\"') >= 0, "Should contain escaped quotes");
      test.ok(result.indexOf('xlsx2csv') == 0, "Should start with string 'xlsx2csv'");
      test.ok(result.indexOf('"a, b, and c"') > 0, "Should contain quoted values");
      //console.log(result);
    }
    test.done();
  });
};