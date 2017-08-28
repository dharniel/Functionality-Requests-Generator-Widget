var spreadSheet = SpreadsheetApp.getActive();
var scriptProperties = PropertiesService.getScriptProperties();

/**
* Initializes a 'Functionality Reuquests' menu in the Google Sheets UI of the active spreadsheet
*/

function onOpen() {
    SpreadsheetApp.getUi().createMenu('Functionality Requests').addItem('Functionality Requests Generator',
         'display').addToUi();
}

/**
* Display the Sidebar Addon
*/

function display() {
    var html = HtmlService.createHtmlOutputFromFile('Sidebar').setWidth(200).setTitle('Generate Functionality Requests');
    SpreadsheetApp.getUi().showSidebar(html);
}

/** 
* Generate alerts based on status id entry
* @params id The status id to generate alert with
*/

function alert(id) {
    switch (id) {
        case 0:
            SpreadsheetApp.getUi().alert('Files and Folders have been generated successfully');
            break;
        case 1:
            SpreadsheetApp.getUi().alert('Sheet has been rendered completely');
            break;
        default:
            break;
    }
}

/** 
* Searches Google Drive for all the files and folders that contain reports
*/

function getAllFoldersAndFiles() {
    SpreadsheetApp.getActiveSpreadsheet().toast("Files and Folders are being generated", "Task Status", 4);
    var folders = [];
    var files = [];
    var foldersCollection = "";
    var content = [folders, files];
    //FolderIterator
    var driveFolders = DriveApp.searchFolders("title contains 'CSD Weekly Reports' and trashed = false"); 
    while (driveFolders.hasNext()) {
      //Google Folder Object
        var folder = driveFolders.next(); 
        var folderName = folder.getName();
        folders.push(folderName);
        foldersCollection += folderName + ";";
       //File Iterator
        var folderFiles = folder.getFiles();
        var fileArray = []; 
        while (folderFiles.hasNext()) {
            var file = folderFiles.next(); 
            var fileName = file.getName();
            fileArray.push(fileName);
        }
        files.push(fileArray);
    };
    return content;
}

/**
* Based on the parameters, requests contained in the files in the queryFolder or the requests in the queryFile 
* are generated onto a sheet
* @param selectionType This determines which category of requests are to be generated - 'All' requests, 
* requests in a 'Folder', or requests in a 'File'
* @param generateType This determines if the requests are generaed on a 'new' sheet or the 'current' sheet 
* within the spreadsheet
* @param queryFolder The folder containing weekly reports files
* @param queryFile A weekly report file
*/

function generateRequestsSheet(selectionType, generateType, queryFolder, queryFile) {
    var activeSheet;
    process: {
        if (generateType == "default" || generateType == "new") {
            var newName;
            var ui = SpreadsheetApp.getUi();
            var response = ui.prompt('Enter a unique name for the new sheet', '', ui.ButtonSet.OK_CANCEL);
            if (response.getSelectedButton() == ui.Button.OK) {
                newName = response.getResponseText();
            } else if (response.getSelectedButton() == ui.Button.CANCEL) {
                break process;
            }
            activeSheet = spreadSheet.insertSheet(newName); //Edit Name;
            activeSheet.appendRow(["DATE", "REQUESTS"]);
        } else if (generateType == "current") {
            activeSheet = spreadSheet.getActiveSheet();
            activeSheet.appendRow(["DATE", "REQUESTS"]);
        }
        
        SpreadsheetApp.getActiveSpreadsheet().toast("Sheet is being generated", "Task Status", 4.5);
      
        switch (selectionType) {
            case "All":
                var folderArray = [];
                var folders = DriveApp.searchFolders("title contains 'CSD Weekly Reports' and trashed = false");
                if (folders.hasNext()) {
                    var folderName = folders.next().getName();
                    folderArray.push(folderName);
                }
                for (var i = 0; i < foldersArray.length; i++) {
                    foldersGenerator(foldersArray[i], activeSheet);
                }
                break;
            case "Folders":
                foldersGenerator(queryFolder, activeSheet);
                break;
            case "Files":
                filesGenerator(queryFile, activeSheet);
                break;
        }
        alert(1);
    }
}

/**
* Helper function that generates individual rows containing date and functionality requests
* @param query The name of the file to be queried for the requests
* @param sheet The selected sheet to generate data on
*/

function filesGenerator(query, sheet) {
    var content = generateRequestsText(query);
    var item = [extractDate(query), content];
    sheet.appendRow(item);
}

/**
* Helper function that generates functionality requests from files in a folder
* @param query The name of the folder to be queried for the requests
* @param sheet The selected sheet to generate data on
*/

function foldersGenerator(query, sheet) {
    var folders = DriveApp.searchFolders("title contains " + "'" + query + "'" + " and trashed = false");
    if (folders.hasNext()) {
        var folderId = folders.next().getId();
    }
    var folder = DriveApp.getFolderById(folderId);
    var folderFiles = folder.getFiles();
    while (folderFiles.hasNext()) {
        var file = folderFiles.next();
        if (file.getMimeType() == "application/vnd.google-apps.document") { //Ensures that file is a document
            var queryValue = file.getName();
            filesGenerator(queryValue, sheet);
        }
    }
}

/**
* Helper function that extracts the requests from a specific file, based on the uniform structure of the file
* contents
* @param fileName The name of the file containing functionality requests
*/

function generateRequestsText(fileName) {
    var requestText;
    var keyStart = "Functionality Requests";
    var keyEnd = "The Coming Week";
    var files = DriveApp.searchFiles("title contains " + "'" + fileName + "'" + " and trashed = false");
    if (files.hasNext()) {
        var docId = files.next().getId();
    }
    var docBody = DocumentApp.openById(docId).getBody();
    var docTextString = docBody.getText();
    var mainContent = docTextString.slice(docTextString.indexOf(keyStart) + keyStart.length,
         docTextString.indexOf(keyEnd));
    requestText = mainContent.trim();
    if (requestText.length > 0) {
        return requestText;
    } else {
        return "NIL";
    }
}

/**
* Helper function to extract the date from a file name based on the uniform structure of the file name
* @param item the name of the file
*/

function extractDate(item) {
    var elementI = item.match(/(\d){6}/)[0];
    var dateString = elementI.slice(0, 2) + "/" + elementI.slice(2, 4) + "/" + elementI.slice(4);
    return dateString;
}