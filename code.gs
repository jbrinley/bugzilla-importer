function onOpen() {
  BugzillaImporter.onOpen();
};

function doImport() {
  var db = ScriptDb.getMyDb();
  BugzillaImporter.doImport(db);
};

function doFullImport() {
  var db = ScriptDb.getMyDb();
  BugzillaImporter.doFullImport(db);
}

