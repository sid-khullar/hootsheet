function createConsolidatedSheet() {

  // Configuration
  Logger.log ("Initialising...");

  var spreadsheetID = "Spreadsheet ID here";
  var configBeginRow = 6;
  var configSheetName = "Config";
  var configSheetCol = 2;
  var configMsgCol = 3;
  var configDateCol = 6;
  var configTagCol = 7;

  // Begin reading configuration
  // traverse column containing sheet names, get configuration data
  // and store, to avoid too many nesting levels 
  Logger.log ("Reading configuration sheet...");

  var hootSheet = SpreadsheetApp.openById(spreadsheetID);
  SpreadsheetApp.setActiveSpreadsheet(hootSheet);
  var configSheet = hootSheet.getSheetByName(configSheetName);
  if (configSheet == null) {
    Logger.log ("ERR: Cannot find configuration worksheet.")
    return
  }
  
  var readConfig = true;
  var readRow = configBeginRow;

  var config_array = []
  while (readConfig){
    
    // read sheet name from range
    var sourceSheet = configSheet.getRange(readRow, configSheetCol).getValue();

    // exit loop when blank cell / value is encountered
    if (sourceSheet == "") {
      readConfig = false;
      break;
    }

    // message, date, tag
    var numMessages = configSheet.getRange(readRow, configMsgCol).getValue();
    var msgDate = configSheet.getRange(readRow, configDateCol).getValue();
    var targetSheet = configSheet.getRange(readRow, configTagCol).getValue();

    // check if all of the variables have values
    if (numMessages == "" || msgDate == "" || targetSheet == "") {
      Logger.log ("ERR: Sheet, Messages, Date and Tag columns must not contain empty cells.");
      return
    }

    // assemble into dictionary
    var config_dict = new Object();
    config_dict.source_sheet = sourceSheet;
    config_dict.num_messages = numMessages;
    config_dict.start_date = msgDate;
    config_dict.target_sheet = targetSheet;

    // add dict to array
    config_array.push (config_dict);

    // increment row
    readRow++;
  
  }

  // we now have an array of dictionaries
  // each element pointing to a specific message configuration.
  // visit each sheet, get the specified number of messages
  // from the specified date onwards and build a set of messages
  // to write to the specified sheet.
  
  // but first, check if the specified sheets exist
  // both source and target, as well as make a list of
  // target sheets
  Logger.log ("Checking source and target sheets...");
  var arrayLength = config_array.length;
  var target_sheets = [];
  for (var i = 0; i < arrayLength; i++) {
    one_config = config_array [i];

    var source = hootSheet.getSheetByName (one_config.source_sheet);
    var target = hootSheet.getSheetByName (one_config.target_sheet);
    if (source == null) {
      Logger.log (`ERR: Sheet '${one_config.source_sheet}' does not exist.`);
      return;
    }
    if (target == null) {
      Logger.log (`ERR: Sheet '${one_config.target_sheet}' does not exist.`);
      return;
    }

    // add target to list of targets, if it
    // isn't already there
    if (!target_sheets.includes (one_config.target_sheet)) {
      target_sheets.push (one_config.target_sheet);
    }

  }

  // checked that the target sheets exist.
  // now visit each specified sheet and find the specified date
  // and pick up the specified number of messages from there.
  Logger.log ("Reading messages from source sheets...");
  var message_list = [];
  for (var i = 0; i < arrayLength; i++) {
    one_config = config_array [i];

    var message_sheet = hootSheet.getSheetByName (one_config.source_sheet);
    var seek_date = one_config.start_date;
    var num_messages = one_config.num_messages;
    var target_sheet = one_config.target_sheet;

    var start_row = 1;
    var date_col = 1;
    var mesg_col = 2;
    var link_col = 3;

    // traverse all the messages in a sheet, looking for
    // the start date. When found, begin counter and begin
    // adding to list
    var traverse = true;
    var start_counter = false;
    var counter = 0;

    // Logger.log (`  processing ${one_config.source_sheet}`);
    while (traverse) {

      var msg_date = message_sheet.getRange(start_row, date_col).getValue();
      var msg_text = message_sheet.getRange(start_row, mesg_col).getValue();
      var msg_link = message_sheet.getRange(start_row, link_col).getValue();

      var date_1 = new Date (seek_date).toDateString();
      var date_2 = new Date (msg_date).toDateString();

      // found the specified date
      if (date_1 == date_2) {
        start_counter = true
      }

      // if the counter has begun, start 
      // storing messages
      if (start_counter) {
        
        if (counter < num_messages) {
          
          if (msg_date == "" || msg_text == "") {
            Logger.log (`ERR: Empty data found in specified range in ${one_config.source_sheet}`);
            return;
          }

          var message_dict = new Object();
          message_dict.target_sheet = target_sheet;
          message_dict.msg_date = msg_date;
          message_dict.msg_text = msg_text;
          message_dict.msg_link = msg_link;
          
          message_list.push (message_dict)
          
          start_row++
          counter++;

        } else {

          // reached specified message limit.
          // stop counter and stop traversing this sheet
          start_counter = false
          traverse = false

        }
        
      }

    }

  }

  Logger.log ("Writing messages to target sheets...")
  // We now have an array of dictionaries, each one with a messag
  // to be written to a specific sheet. We'll travese the original
  // config dictionary and write these to the target sheets.
  num_targets = target_sheets.length;
  for (var i = 0; i < num_targets; i++) {
    
    var target = hootSheet.getSheetByName (target_sheets [i]);
    target.clear();

    start_row = 1
    var date_col = 1;
    var mesg_col = 2;
    var link_col = 3;

    total_messages = message_list.length
    // Logger.log (`  processing ${target_sheets [i]}`);
    for (var m = 0; m < total_messages; m++) {
      
      one_message = message_list [m];

      if (target_sheets [i] == one_message.target_sheet) {
        
        // get range and write date
        range = target.getRange (start_row, date_col);
        range.setValue (one_message.msg_date);

        // get range and write message text
        range = target.getRange (start_row, mesg_col);
        range.setValue (one_message.msg_text);
        // Logger.log (one_message.msg_text)

        // get range and write link 
        range = target.getRange (start_row, link_col);
        range.setValue (one_message.msg_link);

        start_row++;

      }

      // sort this sheet
      range = target.getRange (`A1:C${start_row}`)
      range.sort (1)

      // export as CSV
      // // append ".csv" extension to the sheet name
      // fileName = target.getName() + ".csv";
      // // convert all available sheet data to csv format
      // var csvFile = convertRangeToCsvFile_(fileName, sheet);
      // // create a file in the Docs List with the given name and the csv data
      // var file = folder.createFile(fileName, csvFile);
      // //File downlaod
      // var downloadURL = file.getDownloadUrl().slice(0, -8);
      // showurl(downloadURL);

    }

  }

}