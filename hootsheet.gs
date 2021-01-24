function displayStatus (statusConfig, statusMessage) {

  var workBook = SpreadsheetApp.openById(statusConfig.wbID);
  SpreadsheetApp.setActiveSpreadsheet(workBook);
  var configSheet = workBook.getSheetByName(statusConfig.wsName);

  range = configSheet.getRange (statusConfig.statusRow, statusConfig.statusCol);
  range.setValue (statusMessage);

}

function createConsolidatedSheet() {

  // Configuration
  Logger.log ("Initialising...");

  var spreadsheetID = "XXXXX";
  var configBeginRow = 6;
  var configSheetName = "XXXXX";
  var configSheetCol = 2;
  var configMsgCol = 3;
  var configDateCol = 6;
  var configTagCol = 7;

  var statusConfig = new Object();
  statusConfig.wbID = spreadsheetID;
  statusConfig.wsName = configSheetName;
  statusConfig.statusRow = 2;
  statusConfig.statusCol = 9;

  // Begin reading configuration
  // traverse column containing sheet names, get configuration data
  // and store, to avoid too many nesting levels 
  displayStatus (statusConfig, "Reading configuration sheet...");
  Logger.log ("Reading configuration sheet...");

  var hootSheet = SpreadsheetApp.openById(spreadsheetID);
  SpreadsheetApp.setActiveSpreadsheet(hootSheet);
  var configSheet = hootSheet.getSheetByName(configSheetName);
  if (configSheet == null) {
    displayStatus (statusConfig, "ERR: Cannot find configuration worksheet.");
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
      displayStatus (statusConfig, "ERR: Sheet, Messages, Date and Tag columns must not contain empty cells.");
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
  displayStatus (statusConfig, "Checking source and target sheets...");
  Logger.log ("Checking source and target sheets...");
  var arrayLength = config_array.length;
  var target_sheets = [];
  for (var i = 0; i < arrayLength; i++) {
    one_config = config_array [i];

    var source = hootSheet.getSheetByName (one_config.source_sheet);
    var target = hootSheet.getSheetByName (one_config.target_sheet);
    if (source == null) {
      displayStatus (statusConfig, `ERR: Sheet '${one_config.source_sheet}' does not exist.`);
      Logger.log (`ERR: Sheet '${one_config.source_sheet}' does not exist.`);
      return;
    }
    if (target == null) {
      displayStatus (statusConfig, `ERR: Sheet '${one_config.target_sheet}' does not exist.`);
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
  displayStatus (statusConfig, "Reading messages from source sheets...");
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

    displayStatus (statusConfig, `  processing ${one_config.source_sheet}`);
    Logger.log (`  processing ${one_config.source_sheet}`);
    while (traverse) {

      var msg_date = message_sheet.getRange(start_row, date_col).getValue();
      var msg_text = message_sheet.getRange(start_row, mesg_col).getValue();
      var msg_link = message_sheet.getRange(start_row, link_col).getValue();

      var date_1 = new Date(Date (seek_date).toLocaleString("en-US", {timeZone: "Asia/Kolkata"}));
      var date_2 = new Date(Date (msg_date).toLocaleString("en-US", {timeZone: "Asia/Kolkata"}));

      // var date_1 = new Date (date_1);
      // var date_2 = new Date (date_2);

      dateStr_1 = `${date_1.getFullYear()}/${date_1.getMonth()}/${date_1.getDate()}`
      dateStr_2 = `${date_2.getFullYear()}/${date_2.getMonth()}/${date_2.getDate()}`

      if (one_config.source_sheet == "Quotes") {
        Logger.log (`${dateStr_1}, ${dateStr_2}`);
      }

      // found the specified date
      if (dateStr_1 == dateStr_2) {
        start_counter = true
      }

      if (msg_text == "") {
        traverse = false
      }

      // if the counter has begun, start 
      // storing messages
      if (start_counter) {
        
        if (counter < num_messages) {
          
          if (msg_date == "" || msg_text == "") {
            displayStatus (statusConfig, `ERR: Empty data found in specified range in ${one_config.source_sheet}`);
            Logger.log (`ERR: Empty data found in specified range in ${one_config.source_sheet}`);
            return;
          }

          var message_dict = new Object();
          message_dict.target_sheet = target_sheet;
          message_dict.msg_date = msg_date;
          message_dict.msg_text = msg_text;
          message_dict.msg_link = msg_link;
          
          message_list.push (message_dict)
        
          // increment counter
          counter++;

        } else {

          // reached specified message limit.
          // stop counter and stop traversing this sheet
          start_counter = false
          traverse = false

        }
        
      }

      // increment row
      start_row++

    }

  }

  displayStatus (statusConfig, "Writing messages to target sheets...");
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

    }

  }

  displayStatus (statusConfig, "Completed.");

}
