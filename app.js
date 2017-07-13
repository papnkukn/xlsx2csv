var os = require('os');
var fs = require('fs');
var path = require('path');

var xlsx2csv = require('./lib/index.js');

//Default configuration
var config = {
  force: false,
  sheet: undefined,
  verbose: process.env.NODE_VERBOSE == "true" || process.env.NODE_VERBOSE == "1",
  source: undefined,
  target: undefined
};

//Prints help message
function help() {
  console.log("Usage:");
  console.log("  xlsx2csv [options] <input.xlsx> <output.csv>");
  console.log("");
  console.log("Options:");
  console.log("  --force             Force overwrite file");
  console.log("  --data [type]       Data type to export: formula, value or display");
  console.log("  --sheet [name]      Sheet name to export");
  console.log("  --range [A1:C3]     Range to export");
  console.log("  --separator [char]  CSV column separator, e.g. '\\t'");
  console.log("  --line-end [char]   End of line char(s), e.g. '\\r\\n'");
  console.log("  --help              Print this message");
  console.log("  --verbose           Enable detailed logging");
  console.log("  --version           Print version number");
  console.log("");
  console.log("Examples:");
  console.log("  xlsx2csv --version");
  console.log("  xlsx2csv --verbose --range A1:M30 --separator , file.xlsx");
  console.log("  xlsx2csv --data formula --sheet Sample file.xlsx file.csv");
}

//Command line interface
var args = process.argv.slice(2);
for (var i = 0; i < args.length; i++) {
  switch (args[i]) {
    case "--help":
      help();
      process.exit(0);
      break;
      
    case "-f":
    case "--force":
      config.force = true;
      break;
      
    case "--sheet":
      config.sheet = args[++i];
      break;
      
    case "--range":
      config.range = args[++i]; //"A1:Z99"
      break;
      
    case "--data":
      config.data = args[++i]; //"formula" or "value" or "display"
      break;
      
    case "--separator":
      config.separator = args[++i].replace("\\t", "\t");
      break;
      
    case "--line-end":
      config.lineEnd = args[++i].replace("\\r", "\r").replace("\\n", "\n");
      break;
      
    case "--verbose":
      config.verbose = true;
      break;
    
    case "-v":    
    case "--version":
      console.log(require('./package.json').version);
      process.exit(0);
      break;
      
    default:
      if (args[i].indexOf('-') == 0) {
        console.error("Unknown command line argument: " + args[i]);
        process.exit(2);
      }
      else if (!config.source) {
        config.source = args[i];
      }
      else if (!config.target) {
        config.target = args[i];
      }
      else {
        console.error("Too many arguments: " + args[i]);
        process.exit(2);
      }
      break;
  }
}

//Check if the source file argument is defined
if (!config.source) {
  console.error("Source file not defined!");
  process.exit(2);
}

//Check if the source file exists
if (!fs.existsSync(config.source)) {
  console.error("File not found: " + config.source);
  process.exit(2);
}

//Check if output file is ready to overwrite
if (config.target && !config.force && fs.existsSync(config.target)) {
  console.error("File already exists: " + config.source);
  process.exit(3);
}

//Convert XLSX to CSV
var options = {
  verbose: config.verbose,
  data: config.data,
  sheet: config.sheet,
  range: config.range,
  separator: config.separator,
  lineEnd: config.lineEnd
};

xlsx2csv(config.source, options, function(error, result) {
  if (error) {
    console.error(error);
    process.exit(1);
    return;
  }
  
  if (!config.target) {
    //Print to console
    console.log(result);
    if (config.verbose) {
      console.log("Done!");
    }
  }
  else {
    //Write to file
    fs.writeFileSync(config.target, result);
    console.log("Done!");
  }
  process.exit(0);
});

/*
setTimeout(function() {
  console.log("Timeout!");
  process.exit(1);
}, 60000);
*/