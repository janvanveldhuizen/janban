'use strict';

var outlookApp;
var outlookNS;

function checkBrowser() {
    var isBrowserSupported
    if (window.external !== undefined && window.external.OutlookApplication !== undefined) {
        isBrowserSupported = true;
        outlookApp = window.external.OutlookApplication;
        outlookNS = outlookApp.GetNameSpace("MAPI");
    } else {
        try {
            isBrowserSupported = true;
            outlookApp = new ActiveXObject("Outlook.Application");
            outlookNS = outlookApp.GetNameSpace("MAPI");
        }
        catch (e) {
            isBrowserSupported = false;
        }
    }
    return isBrowserSupported;
}

function getOutlookCategories() {
    var i;
    var catNames = [];
    var catColors = [];
    var categories = outlookNS.Categories;
    var count = outlookNS.Categories.Count;
    catNames.length = count;
    catColors.length = count;
    for (i = 1; i <= count; i++) {
        catNames[i - 1] = categories(i).Name;
        catColors[i - 1] = categories(i).Color;
    };
    return { names: catNames, colors: catColors };
}

function getTaskFolder(folderpath) {
    if (folderpath === undefined || folderpath === '') {
        // if folder path is not defined, return main Tasks folder
        var folder = outlookNS.GetDefaultFolder(13);
    } else {
        // if folder path is defined then find it, create it if it doesn't exist yet
        try {
            var folder = outlookNS.GetDefaultFolder(13).Folders(folderpath);
        }
        catch (e) {
            outlookNS.GetDefaultFolder(13).Folders.Add(folderpath);
            var folder = outlookNS.GetDefaultFolder(13).Folders(folderpath);
        }
    }
    return folder;
}

function getJournalFolder(){
    return outlookNS.GetDefaultFolder(11);
}

function getTask(id){
    return outlookNS.GetItemFromID(id);
}

function newMailItem(){
    return outlookApp.CreateItem(0);
}

function newJournalItem(){
    return outlookApp.CreateItem(4);
}
