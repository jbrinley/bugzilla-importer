/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
      name : "Quick Import",
      functionName: "doImport"
    },
    {
      name : "Full Import",
      functionName: "doFullImport"
    }
  ];
  sheet.addMenu("Bugzilla", entries);
};

function doImport() {
  var now = Date.now();
  var timelimit = now + (1000*60*4); //  do all we can in four minutes
  var db = ScriptDb.getMyDb();
  var importer = new BugzillaImporter();
  var newQuery = true;
  var response, numBugs, bug_id;
  
  var abort = db.count({
    type: 'flag',
    name: 'working',
    value: db.greaterThan(Date.now()-(60*10*1000))
  });
  if ( abort > 0 ) {
    return; // this process is already running. let's not interfere
  }
  
  var working = db.save({
    type: 'flag',
    name: 'working',
    value: Date.now()
  });
  
  var result = db.query({
    type: 'flag',
    name: 'to_process'
  });
  
  if ( result.hasNext() ) {
    var data = result.next();
    if ( data.bugs.length > 0 ) {
      newQuery = false;
    }
  } else {
    var data = {
      type: 'flag',
      name: 'to_process',
      bugs: []
    };
  }
  
  if ( newQuery ) {
    var last_update_result = db.query({
      type: 'flag',
      name: 'last_update'
    });
    if ( last_update_result.hasNext() ) {
      var last_update = last_update_result.next();
    } else {
      var last_update = {
        type: 'flag',
        name: 'last_update',
        value: 0
      };
    }
    data.bugs = importer.doQuery(last_update.value);
    db.save(data);
    last_update.value = now;
    db.save(last_update);
  }
  
  numBugs = data.bugs.length;
  
  for ( var i = 0 ; i < numBugs && Date.now() < timelimit ; i++ ) {
    bug_id = data.bugs.shift();
    importer.update(bug_id);
  }
  db.save(data);
  
  db.remove(working);
  
  if ( data.bugs.length > 0 ) {
    Browser.msgBox(
      'Import Incomplete',
      'The import ran successfully, but did not have time to finish. Please run the Quick Import again. It will pick up where it left off.',
      Browser.Buttons.OK
    );
  }
  return;
};

function doFullImport() {
  var db = ScriptDb.getMyDb();
  var last_update_result = db.query({
    type: 'flag',
    name: 'last_update'
  });
  while ( last_update_result.hasNext() ) {
    db.remove(last_update_result.next());
  }
  var to_process_result = db.query({
    type: 'flag',
    name: 'to_process'
  });
  while ( to_process_result.hasNext() ) {
    db.remove(to_process_result.next());
  }
  
  var fields = new BugzillaFields();
  fields.flushCache();
  
  doImport();
}

function doDebug() {
  var query = {
    product: 'Boot2Gecko',
    last_change_time: Utilities.formatDate(new Date(Date.now()-1000*60*60*24*14), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'"),
    include_fields: [ 'id' ],
    limit: 100
  };
  var query_url = 'https://bugzilla.mozilla.org/jsonrpc.cgi?method=Bug.search&params='+encodeURIComponent(JSON.stringify([query]));
  var response = UrlFetchApp.fetch(query_url);
  response = JSON.parse(response);
  return;
};

function BugzillaFields() {
  this.db = null;
  this.getDB();
  this.maybeRefreshCache();
};

BugzillaFields.prototype.getName = function( field_display_name ) {
  var result = this.db.query({
    type: 'field',
    display_name: field_display_name
  });
  if ( result.hasNext() ) {
    return result.next().name;
  } else {
    return '';
  }
};

BugzillaFields.prototype.getValues = function( field_name ) {
  var result = this.db.query({
    type: 'field',
    name: field_name
  });
  if ( result.hasNext() ) {
    return result.next().values;
  } else {
    return [];
  }
}

BugzillaFields.prototype.getDB = function() {
  this.db = ScriptDb.getMyDb();
};

/**
 * Refresh the field cache once per week
 */
BugzillaFields.prototype.maybeRefreshCache = function() {
  var count = this.db.count({
    type: 'flag',
    name: 'fields_updated',
    value: this.db.greaterThan(Date.now()-(60*60*24*7*1000))
  });
  if ( count < 1 ) {
    this.importFields();
  }
};

/**
 * Import all fields from Bugzilla and store in a local cache
 */
BugzillaFields.prototype.importFields = function() {
  this.flushCache();
  
  var query_url = 'https://bugzilla.mozilla.org/jsonrpc.cgi?method=Bug.fields';
  var response = UrlFetchApp.fetch(query_url);
  response = JSON.parse(response);
  
  if ( response.error != null ) {
    return false;
  }
  if ( response.result.fields.length < 1 ) {
    return false;
  }
  
  for ( var i = 0 ; i < response.result.fields.length ; i++ ) {
    var field = response.result.fields[i];
    if ( !(field.type == 2 || field.type == 3) || !field.hasOwnProperty('values') || field.values.length > 1000 ) {
      continue;
    }
    var object = {
      type: 'field',
      name: field.name,
      display_name: field.display_name,
      id: field.id,
      values: []
    };
    for ( var j = 0 ; j < field.values.length ; j++ ) {
      if ( field.values[j].name != '---' ) {
        object.values.push(field.values[j].name);
      }
    }
    this.db.save(object);
  }
  
  this.db.save({
    type: 'flag',
    name: 'fields_updated',
    value: Date.now()
  });
};

/**
 * Flush all fields from the cache
 */
BugzillaFields.prototype.flushCache = function() {
  while (true) {
    var result = this.db.query({type: 'field'}); // get everything, up to limit
    if (result.getSize() == 0) {
      break;
    }
    while (result.hasNext()) {
      this.db.remove(result.next());
    }
  }
  
  while (true) {
    var result = this.db.query({type: 'flag', name: 'fields_updated'}); // get everything, up to limit
    if (result.getSize() == 0) {
      break;
    }
    while (result.hasNext()) {
      this.db.remove(result.next());
    }
  }
};

function BugzillaImporter() {
  this.dataSheet = null;
  this.map = {}; // a map of bug IDs to rows
  this.rowHeaders = {}; // a map of headers to columns
  this.headerColumns = {}; // a map of columns to headers
  this.config = null;
  this.currentDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  
  this.readConfig();
  this.readHeaders();
  this.buildMap();
};

BugzillaImporter.prototype.readConfig = function() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName("Settings");
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  
  this.config = {};
  
  for ( var i = 0 ; i < numRows ; i++ ) {
    var row = values[i];
    if ( row[1] != '' ) {
      this.config[row[0].toLowerCase()] = row[1];
    }
  }
};

BugzillaImporter.prototype.readHeaders = function() {
  var sheet = this.getDataSheet();
  var range = sheet.getRange( 1, 1, 1, sheet.getLastColumn() );
  var values = range.getValues()[0];
  for ( var i in values ) {
    this.rowHeaders[values[i]] = i;
    this.headerColumns[i] = values[i];
  }
};

BugzillaImporter.prototype.getQuery = function() {
  var query = {};
  
  for ( var i in this.config ) {
    if ( i == 'product' ) {
      query.product = this.parseConfigValue(this.config[i]);
      continue;
    }
    if ( i == 'component' ) {
      query.component = this.parseConfigValue(this.config[i]);
    }
    if ( i == 'whiteboard' ) {
      query.whiteboard = this.parseConfigValue(this.config[i]);
    }
  }
  
  return query;
};

BugzillaImporter.prototype.getFlags = function() {
  var output = [];
  for ( var i in this.config ) {
    if ( i = 'flag' ) {
      fields = new BugzillaFields();
      var flags = this.parseConfigValue(this.config[i]);
      if ( typeof(flags) != 'undefined' ) {
        for ( var i = 0 ; i < flags.length ; i++ ) {
          var field_name = fields.getName(flags[i]);
          if ( field_name != '' ) {
            var values = fields.getValues(field_name);
            if ( values.length > 0 ) {
              output.push({name: field_name, values: values});
            }
          }
        }
      }
      break;
    }
  }
  
  return output;
}

BugzillaImporter.prototype.parseConfigValue = function( value ) {
  if ( value == '' || typeof(value) == 'undefined' ) {
    return [];
  }
  return value.split(/,\s*/);
};

BugzillaImporter.prototype.doQuery = function( last_update ) {
  var uniques = {}; // store IDs of found bugs, so that we return an array of unique bugs
  var consolidated_bugs = [];
  var query = this.getQuery();
  if ( query == false ) {
    return [];
  }
  
  query.include_fields = [ 'id' ];
  if ( last_update > 0 ) {
    var last_update_date = new Date(last_update);
    query.last_change_time = Utilities.formatDate(last_update_date, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  }

  // if we have flags, we need to loop through them, doing multiple queries, as we can't do an OR search with this API
  var flags = this.getFlags();
  if ( flags.length > 0 ) {
    for ( var j = 0 ; j < flags.length ; j++ ) {
      // clone the query
      var flag_query = JSON.parse(JSON.stringify(query));
      flag_query[flags[j].name] = flags[j].values;
      var query_url = 'https://bugzilla.mozilla.org/jsonrpc.cgi?method=Bug.search&params='+encodeURIComponent(JSON.stringify([flag_query]));
      var response = UrlFetchApp.fetch(query_url);
      response = JSON.parse(response);
      if ( response.error != null ) {
        return [];
      }
      var bugs = response.result.bugs;
      for ( var i = 0 ; i < bugs.length ; i++ ) {
        if ( !uniques.hasOwnProperty(bugs[i].id) ) {
          uniques[bugs[i].id] = 1;
          consolidated_bugs.push(bugs[i].id);
        }
      }
    }
  } else {
    var query_url = 'https://bugzilla.mozilla.org/jsonrpc.cgi?method=Bug.search&params='+encodeURIComponent(JSON.stringify([query]));
    var response = UrlFetchApp.fetch(query_url);
    response = JSON.parse(response);
    if ( response.error != null ) {
      return [];
    }
    var bugs = response.result.bugs;
    for ( var i = 0 ; i < bugs.length ; i++ ) {
      if ( !uniques.hasOwnProperty(bugs[i].id) ) {
        uniques[bugs[i].id] = 1;
        consolidated_bugs.push(bugs[i].id);
      }
    }
  }
  
  var count = consolidated_bugs.length;
  return consolidated_bugs;
};

BugzillaImporter.prototype.buildMap = function() {
  var sheet = this.getDataSheet();
  var lastRow = sheet.getLastRow();
  if ( lastRow < 2 ) {
    return;
  }
  var rows = sheet.getRange(2,1,lastRow-1,1);
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  
  for ( var i = 0 ; i < numRows ; i++ ) {
    var row = values[i];
    this.map[row[0]] = i+2;
  }
};

BugzillaImporter.prototype.getDataSheet = function() {
  if ( this.dataSheet == null ) {
    
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    this.dataSheet = doc.getSheetByName("Data");
  }
  return this.dataSheet;
};

BugzillaImporter.prototype.update = function( bug ) {
  var bug = this.getBug(bug); // get a full bug object
  var sheet = this.getDataSheet();
  if ( bug.id in this.map ) {
    var range = sheet.getRange( this.map[bug.id], 1, 1, sheet.getLastColumn() );
    var before = range.getValues()[0];
    var after = this.populateRow( JSON.parse(JSON.stringify(before)), bug ); // clone the object before passing it
    var changes = [];
    for ( var i in before ) {
      if ( typeof(before[i]) == 'string' && before[i] != after[i] ) { // don't try to compare on non strings (mostly will be Date objects)
        changes.push(this.headerColumns[i]);
      }
    }
    if ( changes.length > 0 ) {
      if ( after[this.rowHeaders['update log']] != '' ) {
        after[this.rowHeaders['update log']] += "\n";
      }
      after[this.rowHeaders['update log']] = after[this.rowHeaders['update log']] + this.currentDate + ": " + changes.join(', ');
      range.setValues([after]);
    }
  } else {
    var values = new Array(sheet.getLastColumn()); // create an array with the required number of elements
    values[0] = bug.id; // ID is always in first column
    values = this.populateRow(values, bug);
    values[this.rowHeaders['update log']] = this.currentDate + ": new";
    sheet.appendRow(values);
    this.map[bug.id] = sheet.getLastRow();
  }
};

BugzillaImporter.prototype.populateRow = function( row, bug ) {
  for ( var field in this.rowHeaders ) {
    if ( field == 'id' ) { continue; }
    
    if ( field == 'summary' ) {
      row[this.rowHeaders[field]] = bug.summary;
      continue;
    }
    
    // special handling of the UX-* whiteboard flags
    if ( field == 'whiteboard:UX' ) {
      var match = bug.whiteboard.match(/\bUX-([\w\?\+]+)/);
      if ( Array.isArray(match) && match.length > 1 ) {
        row[this.rowHeaders[field]] = match[1];
      } else {
        row[this.rowHeaders[field]] = '';
      }
      continue;
    }
    
    // any other whiteboard flags get a simple Yes/No match
    if ( field.indexOf('whiteboard:') == 0 ) {
      if ( bug.whiteboard.match(new RegExp("\b"+field.substring(11)+"\b")) ) {
        row[this.rowHeaders[field]] = 'Y';
      } else {
        row[this.rowHeaders[field]] = 'N';
      }
      continue;
    }
    
    // does the bug have an attached patch?
    if ( field == 'patch' ) {
      if ( this.hasPatch(bug) ) {
        row[this.rowHeaders[field]] = 'Y';
      } else {
        row[this.rowHeaders[field]] = 'N';
      }
      continue;
    }
    
    if ( field.indexOf('patch:') == 0 ) {
      var field_name = field.substring(6).replace('-', '_').replace(' ', '_');
      row[this.rowHeaders[field]] = this.getPatchFlagValue(bug, field_name);
    }
        
    if ( field.indexOf('flag:') == 0 ) {
      var field_name = 'cf_' + field.substring(5).replace('-', '_').replace(' ', '_');
      if ( bug.hasOwnProperty(field_name) && bug[field_name] != '---' ) {
        row[this.rowHeaders[field]] = bug[field_name];
      } else {
        row[this.rowHeaders[field]] = '';
      }
      continue;
    }
    
    if ( field == 'updated' ) {
      var date_parts = bug.last_change_time.substr(0,10).split('-');
      var time_parts = bug.last_change_time.substr(11,8).split(':');

      row[this.rowHeaders[field]] = new Date(Date.UTC(date_parts[0], date_parts[1]-1, date_parts[2], time_parts[0], time_parts[1], time_parts[2]));
      continue;
    }
    
    if ( field.indexOf('status:') == 0 ) {
      if ( field.indexOf('>') > 0 ) {
        var looking_for = field.substring(7).split('>');
        if ( bug.status.toLowerCase() == looking_for[0].toLowerCase() && bug.resolution.toLowerCase() == looking_for[1].toLowerCase() ) {
          row[this.rowHeaders[field]] = 'Y';
        } else {
          row[this.rowHeaders[field]] = '';
        }
      } else {
        var looking_for = field.substring(7);
        if ( bug.status.toLowerCase() == looking_for.toLowerCase() ) {
          row[this.rowHeaders[field]] = 'Y';
        } else {
          row[this.rowHeaders[field]] = '';
        }
      }
      continue;
    }
    
    
  }
  for ( var i = 0 ; i < row.length ; i++ ) {
    if ( typeof(row[i]) == 'undefined' ) {
      row[i] = '';
    }
  }
  return row;
};

BugzillaImporter.prototype.getBug = function( id ) {
  var iterate = function(obj) {
   for (var property in obj) {
    if (obj.hasOwnProperty(property)) {
     if (typeof obj[property] == "object")
      iterate(obj[property]);
     else
      Logger.log(property + "   " + obj[property]);
    }
   }
  }
  var params = {
    ids: id.toString()
  };
  
  var response = UrlFetchApp.fetch('https://bugzilla.mozilla.org/jsonrpc.cgi?method=Bug.get&params='+encodeURIComponent(JSON.stringify([params])));
  response = JSON.parse(response);
  if ( response.error != null ) {
    return {};
  }
  if ( response.result.bugs.length < 1 ) {
    return {};
  }
  var bug = response.result.bugs[0];
  bug.id = bug.id.toString();
  bug.attachments = [];
  
  var att_response = UrlFetchApp.fetch('https://bugzilla.mozilla.org/jsonrpc.cgi?method=Bug.attachments&params='+encodeURIComponent(JSON.stringify([{ids: [bug.id], exclude_fields: ['data']}])));
  att_response = JSON.parse(att_response);
  
  if ( att_response.error == null && att_response.result.bugs.hasOwnProperty(bug.id) ) {
    bug.attachments = att_response.result.bugs[bug.id];
  }
  return bug;
};

/**
 * Get an attachment for the bug marked as a patch
 * @return object|false
 */
BugzillaImporter.prototype.getPatch = function( bug ) {
  for ( var i = 0 ; i < bug.attachments.length ; i++ ) {
    if ( bug.attachments[i].is_patch == 1 ) {
      return bug.attachments[i];
    }
  }
  return false;
};

/**
 * Does the bug have an attachment marked as a patch
 * @return bool
 */
BugzillaImporter.prototype.hasPatch = function( bug ) {
  var patch = this.getPatch(bug);
  return patch !== false;
};

/**
 * Get the status of a patch's flag
 * @return string
 */
BugzillaImporter.prototype.getPatchFlagValue = function( bug, flag ) {
  var patch = this.getPatch(bug);
  if ( patch === false ) {
    return '';
  }
  for ( var i = 0 ; i < patch.flags.length ; i++ ) {
    if ( patch.flags[i].name == flag ) {
      return patch.flags[i].status;
    }
  }
  return '';
};
