// Script pulls a list of all files and folders in a drive based on folder ID
// Source: https://gist.github.com/abubelinha/c797c4b9c5f0da28e351de20ab7f433c
// YT link: https://www.youtube.com/watch?v=8OghJCld6DY&lc=Ugwh6jG5kAWh8W_uS4l4AaABAg.9WmuPNiDPCK9cxSgNFc-Aa



function onOpen() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('List Files/Folders')
    .addItem('List All Parent/Child Files and Folders', 'listFilesAndFolders')
    .addItem('Duplicate Active Tab (multiple)', 'duplicateActiveTab')
    .addItem('Set Folders Formatting to Bold', 'copyConditional')
    .addToUi();
};

//For my purposes I needed to duplicate the current tab X times and was annoyed by doing it manually :D only use then when on the current tab
function duplicateActiveTab() {
  var duplicateActiveTabCount = Browser.inputBox('How many tabs?', Browser.Buttons.OK_CANCEL);
  var currentTabCount = 0;
  for (var counter = 0; counter < duplicateActiveTabCount; counter = counter + 1) {
    SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet()
    SpreadsheetApp.getActiveSheet().setName("name"  + counter);
  }  
}


// This was something I did as I was doing a migration from a shared folder to a shared drive and wanted to highlight the ones that were folders.
function copyConditional(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("A2:L");
  var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$C2=\"Folder\"")
      .setBackground("#F3F3F3")
      .setBold(true)
      .setRanges([range])
      .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}


function listFilesAndFolders(){
  var folderId = Browser.inputBox('Enter folder ID', Browser.Buttons.OK_CANCEL);
  if (folderId === "") {
    Browser.msgBox('Folder ID is invalid');
    return;
  }
  getFolderTree(folderId, true); 
};

// Get Folder Tree
function getFolderTree(folderId, listAll) {
  try {
    // Get folder by id
    var parentFolder = DriveApp.getFolderById(folderId);
    
    // Initialise the sheet
    var file, data, sheet = SpreadsheetApp.getActiveSheet();
    // Clear just the contents (lets you keep formatting)
    sheet.clearContents();
    sheet.appendRow(["Full Path", "Name","Type" ,"Date", "Last Updated", "Description","Owner Email","Owner Name","Owner Avatar"]);
    
    // Get files and folders
    getParentFiles(parentFolder.getName(), parentFolder, data, sheet, listAll);
    // Note: When Grabbing JUST your main parent DIR it is quickest to just comment out the getChildFolders() part.
    getChildFolders(parentFolder.getName(), parentFolder, data, sheet, listAll);

    // Freeze Row 1 and sort by Full Path
    sheet.setFrozenRows(1);
    sheet.sort(1); 
  } catch (e) {
    Logger.log(e.toString());
  }
};


// getParentFiles() (called on line 66)
function getParentFiles(parentName, parent, data, sheet, listAll) {
    // Get files in parent folder
    // NOTE: ONLY RUN THIS ONCE FOR MAIN ROOT FOLDER AND COPY RESULTS TO OTHER TAB THEN COMMENT OUT RE-RUN script on Sub folders
    var parentFiles = parent.getFiles();
    while (listAll & parentFiles.hasNext()) {
      var parentFile = parentFiles.next();
      data = [
        parentName,
        "=hyperlink(\""+ parentFile.getUrl() + "\",\"" + parentFile.getName().replace(/\"+/g,"")+ "\")",
        //parentFile.getName() // Get File Name
        parentFile.getMimeType(),
        //"File",
        parentFile.getDateCreated(),
        // parentFile.getUrl(),
        parentFile.getLastUpdated(),
        parentFile.getDescription(),
        parentFile.getOwner().getEmail(),
        parentFile.getOwner().getName(),
        "=IMAGE(\"" + parentFile.getOwner().getPhotoUrl() + "\")"
      ];
      //Write
      sheet.appendRow(data);
    }
    sheet.setName(parentName);
};

// Get the list of files and folders and their metadata in recursive mode
function getChildFolders(parentName, parent, data, sheet, listAll) {
  var childFolders = parent.getFolders();

  // List folders inside the folder
  while (childFolders.hasNext()) {
    var childFolder = childFolders.next();
    var folderId = childFolder.getId();
    data = [ 
      parentName + "/" + childFolder.getName(),
      "=hyperlink(\""+ childFolder.getUrl() + "\",\"" + childFolder.getName().replace(/\"+/g,"")+ "\")",
      //childFolder.getName(),
      "Folder",
      childFolder.getDateCreated(),
      // childFolder.getUrl(),
      childFolder.getLastUpdated(),
      childFolder.getDescription(),
      childFolder.getOwner().getEmail(),
      childFolder.getOwner().getName(),
      "=IMAGE(\"" + childFolder.getOwner().getPhotoUrl() + "\")"
    ];
    // Write
    sheet.appendRow(data);
    

    // List files inside the child folders
    //NOTE: ONLY RUN THIS FOR CHILD FOLDERS OFF THE ROOT. If you run this in the root, you will get everything.
    var files = childFolder.getFiles();
    while (listAll & files.hasNext()) {
      var childFile = files.next();
      data = [ 
        parentName + "/" + childFolder.getName(),
        "=hyperlink(\""+ childFile.getUrl() + "\",\"" + childFile.getName().replace(/\"+/g,"")+ "\")",
        //childFile.getName(),
        //"Files",
        childFile.getMimeType(),
        childFile.getDateCreated(),
        // childFile.getUrl(),
        childFile.getLastUpdated(),
        childFile.getDescription(),
        childFile.getOwner().getEmail(),
        childFile.getOwner().getName(),
        "=IMAGE(\"" + childFile.getOwner().getPhotoUrl() + "\")"
      ];
      // Write
      sheet.appendRow(data);
    }


    // Recursive call of the subfolder
    getChildFolders(parentName + "/" + childFolder.getName(), childFolder, data, sheet, listAll);
     
  }
};


// Custom Formulas
// =sheetnames() --> List out all tabnames in a sheet
function sheetnames() {
  var out = new Array()
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) out.push( [ sheets[i].getName() ] )
  return out 
}




