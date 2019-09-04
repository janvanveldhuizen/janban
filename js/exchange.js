'use strict';

var outlookApp;
var outlookNS;

const SENSITIVITY = { olNormal: 0, olPrivate: 2 };

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

function getOutlookTodayHomePageFolder() {
    return outlookNS.GetDefaultFolder(13).Parent.WebViewUrl.replace('\\kanban.html', '')
}

function getOutlookVersion() {
    return outlookApp.version;
}

function getTaskFolder(folderName) {
    if (folderName === undefined || folderName === '') {
        // if folder path is not defined, return main Tasks folder
        var folder = outlookNS.GetDefaultFolder(13);
    } else {
        // if folder path is defined then find it, create it if it doesn't exist yet
        try {
            var folder = outlookNS.GetDefaultFolder(13).Folders(folderName);
        }
        catch (e) {
            outlookNS.GetDefaultFolder(13).Folders.Add(folderName);
            var folder = outlookNS.GetDefaultFolder(13).Folders(folderName);
        }
    }
    return folder;
}

function test() {
    try {
        // var folderAccount = outlookNS.Folders.item("jan.van.veldhuizen@onguard.com");
        var folderAccount = outlookNS.Folders.item("TasksAndCalendar");
        // alert(folderAccount.Folders.item("Tasks"))
        alert(folderAccount.Folders.Count)

        var count = folderAccount.Folders.Count;
        for (var i = 1; i <= count; i++) {
            if (folderAccount.Folders(i).DefaultItemType == 3) {
                alert(folderAccount.Folders(i))
                alert(folderAccount.Folders(i).DefaultItemType)
            }
        };

    } catch (error) {
        alert(error)
    }
}

function getTaskFolder2(folderName) {
    if (folderName === undefined || folderName === '') {
        // if folder path is not defined, return main Tasks folder
        var folder = outlookNS.GetDefaultFolder(13);
    } else {
        // if folder path is defined then find it, create it if it doesn't exist yet
        try {
            var folder = outlookNS.GetDefaultFolder(13).Folders(folderName);
        }
        catch (e) {
            outlookNS.GetDefaultFolder(13).Folders.Add(folderName);
            var folder = outlookNS.GetDefaultFolder(13).Folders(folderName);
        }
    }
    return folder;
}

function getJournalFolder() {
    return outlookNS.GetDefaultFolder(11);
}

function getTaskItems(folderName) {
    return getTaskFolder(folderName).Items;
}

function getTaskItem(id) {
    return outlookNS.GetItemFromID(id);
}

function newMailItem() {
    return outlookApp.CreateItem(0);
}

function newJournalItem() {
    return outlookApp.CreateItem(4);
}

function getJournalItem(subject) {
    var folder = getJournalFolder();
    var configItems = folder.Items.Restrict('[Subject] = "' + subject + '"');
    if (configItems.Count > 0) {
        var configItem = configItems(1);
        if (configItem.Body) {
            return configItem.Body;
        }
    }
    return null;
}

function getPureJournalItem(subject) {
    var folder = getJournalFolder();
    var configItems = folder.Items.Restrict('[Subject] = "' + subject + '"');
    if (configItems.Count > 0) {
        var configItem = configItems(1);
        return configItem;
    }
    return null;
}

function saveJournalItem(subject, body) {
    var folder = getJournalFolder();
    var configItems = folder.Items.Restrict('[Subject] = "' + subject + '"');
    if (configItems.Count == 0) {
        var configItem = newJournalItem();
        configItem.Subject = subject;
    }
    else {
        configItem = configItems(1);
    }
    configItem.Body = body;
    configItem.Save();
}

function getUserEmailAddress() {
    try {
        return outlookNS.Accounts.Item(1).SmtpAddress;
    } catch (error) {
        return 'address-unknown';
    }
}

function getUserName() {
    try {
        return outlookApp.Session.CurrentUser.Name;
    } catch (error) {
        return 'name-unknown';
    }
}

function getUserProperty(item, prop) {
    var userprop = item.UserProperties(prop);
    var value = '';
    if (userprop != null) {
        value = userprop.Value;
    }
    return value;
};

