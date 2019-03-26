'use strict';

var tbApp = angular.module('taskboardApp', ['ui.sortable']);

tbApp.controller('taskboardController', function ($scope, $filter, $http) {

    var applMode;
    var outlookCategories;

    const VERSION_URL = 'http://janware.nl/gitlab/version';

    const APP_MODE = 0;
    const CONFIG_MODE = 1;
    const HELP_MODE = 2;

    const STATE_ID = "KanbanState";
    const CONFIG_ID = "KanbanConfig";

    $scope.privacyFilter = 
        { all:     { text : "Both", value : 0 },
          private: { text : "Private", value: 1 },
          public:  { text : "Work", value: 2 }
        };
    $scope.display_message = false;


    $scope.switchToAppMode = function () {
        applMode = APP_MODE;
    }

    $scope.switchToConfigMode = function () {
        applMode = CONFIG_MODE;
    }

    $scope.switchToHelpMode = function () {
        applMode = HELP_MODE;
    }

    $scope.inAppMode = function () {
        return applMode === APP_MODE;
    }

    $scope.inConfigMode = function () {
        return applMode === CONFIG_MODE;
    }

    $scope.inHelpMode = function () {
        return applMode === HELP_MODE;
    }

    $scope.init = function () {

        $scope.isBrowserSupported = checkBrowser();
        if (!$scope.isBrowserSupported)
        {
            return;
        }

        $scope.switchToAppMode();
        getConfig();
        getState();
        getVersion();
        
        outlookCategories = getOutlookCategories();
        $scope.initTasks();

        $scope.folders = 
            { count: 0 };
        if ($scope.config.BACKLOG_FOLDER.ACTIVE) $scope.folders.count++;
        if ($scope.config.NEXT_FOLDER.ACTIVE) $scope.folders.count++;
        if ($scope.config.INPROGRESS_FOLDER.ACTIVE) $scope.folders.count++;
        if ($scope.config.WAITING_FOLDER.ACTIVE) $scope.folders.count++;
        if ($scope.config.COMPLETED_FOLDER.ACTIVE) $scope.folders.count++;


        // ui-sortable options and events
        $scope.sortableOptions = {
            connectWith: '.tasklist',
            items: 'li',
            opacity: 0.5,
            cursor: 'move',
            containment: 'document',

            stop: function (e, ui) {
                // locate the target folder in outlook
                // ui.item.sortable.droptarget[0].id represents the id of the target list
                if (ui.item.sortable.droptarget) { // check if it is dropped on a valid target
                    if (($scope.config.INPROGRESS_FOLDER.LIMIT !== 0 && e.target.id !== 'inprogressList' && ui.item.sortable.droptarget.attr('id') === 'inprogressList' && $scope.inprogressTasks.length > $scope.config.INPROGRESS_FOLDER.LIMIT) ||
                    ($scope.config.NEXT_FOLDER.LIMIT !== 0 && e.target.id !== 'nextList' && ui.item.sortable.droptarget.attr('id') === 'nextList' && $scope.nextTasks.length > $scope.config.NEXT_FOLDER.LIMIT) ||
                    ($scope.config.WAITING_FOLDER.LIMIT !== 0 && e.target.id !== 'waitingList' && ui.item.sortable.droptarget.attr('id') === 'waitingList' && $scope.waitingTasks.length > $scope.config.WAITING_FOLDER.LIMIT)) {
                        $scope.initTasks();
                        ui.item.sortable.cancel();
                } else {
                    switch (ui.item.sortable.droptarget[0].id) {
                        case 'backlogList':
                            var tasksfolder = getTaskFolder($scope.config.BACKLOG_FOLDER.NAME);
                            var newstatus = $scope.config.STATUS.NOT_STARTED.VALUE;
                            break;
                        case 'nextList':
                            var tasksfolder = getTaskFolder($scope.config.NEXT_FOLDER.NAME);
                            var newstatus = $scope.config.STATUS.NOT_STARTED.VALUE;
                            break;
                        case 'inprogressList':
                            var tasksfolder = getTaskFolder($scope.config.INPROGRESS_FOLDER.NAME);
                            var newstatus = $scope.config.STATUS.IN_PROGRESS.VALUE;
                            break;
                        case 'waitingList':
                            var tasksfolder = getTaskFolder($scope.config.WAITING_FOLDER.NAME);
                            var newstatus = $scope.config.STATUS.WAITING.VALUE;
                            break;
                        case 'completedList':
                            var tasksfolder = getTaskFolder($scope.config.COMPLETED_FOLDER.NAME);
                            var newstatus = $scope.config.STATUS.COMPLETED.VALUE;
                            break;
                    };

                    // locate the task in outlook namespace by using unique entry id
                    var taskitem = getTaskItem(ui.item.sortable.model.entryID);
                    var itemChanged = false;

                    // set new status, if different
                    if (taskitem.Status != newstatus) {
                        taskitem.Status = newstatus;
                        taskitem.Save();
                        itemChanged = true;
                        ui.item.sortable.model.status = taskStatusText(newstatus);
                        ui.item.sortable.model.completeddate = new Date(taskitem.DateCompleted)
                    }

                    // ensure the task is not moving into same folder
                    if (taskitem.Parent.Name != tasksfolder.Name) {
                        // move the task item
                        taskitem = taskitem.Move(tasksfolder);
                        itemChanged = true;

                        // update entryID with new one (entryIDs get changed after move)
                        // https://msdn.microsoft.com/en-us/library/office/ff868618.aspx
                        ui.item.sortable.model.entryID = taskitem.EntryID;
                    }

                    if (itemChanged) {
                        $scope.initTasks();
                    }
                }}
            }
        };

        // watch search filter and apply it
        $scope.$watchGroup(['filter.search', 'filter.private'], function (newValues, oldValues) {
            var search = newValues[0];
            $scope.applyFilters();
            saveState();
        });
    };

    $scope.submitConfig = function (editedConfig) {
        var delta = DeepDiff.diff(editedConfig, $scope.configRaw);
        if (delta){
            try {
                var newConfig = JSON.parse(JSON.minify(editedConfig));
                $scope.config = newConfig;
                saveConfig();
                $scope.init();
            }
            catch (e) {
                alert("I am afraid there is something wrong with the json structure of your configuration data. Please correct it.");
                return;
            }
        }
        $scope.switchToAppMode();
    }

    // borrowed from http://stackoverflow.com/a/30446887/942100
    var fieldSorter = function (fields) {
        return function (a, b) {
            return fields
                .map(function (o) {
                    var dir = 1;
                    if (o[0] === '-') {
                        dir = -1;
                        o = o.substring(1);
                    }
                    var propOfA = a[o];
                    var propOfB = b[o];

                    //string comparisons shall be case insensitive
                    if (typeof propOfA === "string") {
                        propOfA = propOfA.toUpperCase();
                        propOfB = propOfB.toUpperCase();
                    }

                    if (propOfA > propOfB) return dir;
                    if (propOfA < propOfB) return -(dir);
                    return 0;
                }
                ).reduce(function firstNonZeroValue(p, n) {
                    return p ? p : n;
                }, 0
                );
        };
    };

    var getTasksFromOutlook = function (path, sort, folderStatus) {
        var i, array = [];
        var tasks = getTaskItems(path);

        var count = tasks.Count;
        for (i = 1; i <= count; i++) {
            var task = tasks(i);
            if (task.Status == folderStatus) {
                array.push({
                    entryID: task.EntryID,
                    subject: task.Subject,
                    priority: task.Importance,
                    startdate: new Date(task.StartDate),
                    duedate: new Date(task.DueDate),
                    sensitivity: task.Sensitivity,
                    categories: getCategoryStyles(task.Categories),
                    notes: taskBodyNotes(task.Body, $scope.config.TASKNOTE_MAXLEN),
                    status: taskStatusText(task.Status),
                    oneNoteTaskID: getUserProperty(tasks(i), "OneNoteTaskID"),
                    oneNoteURL: getUserProperty(tasks(i), "OneNoteURL"),
                    completeddate: new Date(task.DateCompleted),
                    percent: task.PercentComplete,
                    owner: task.Owner,
                    totalwork: task.TotalWork,
                });
            }
        };

        // sort tasks
        var sortKeys;
        if (sort === undefined) { sortKeys = ["-priority"]; }
        else { sortKeys = sort.split(","); }

        var sortedTasks = array.sort(fieldSorter(sortKeys));

        return sortedTasks;
    };

    $scope.openOneNoteURL = function (url) {
        window.event.returnValue = false;
        if (navigator.msLaunchUri) {
            navigator.msLaunchUri(url);
        } else {
            window.open(url, "_blank").close();
        }
        return nfalse;
    }

    $scope.initTasks = function () {
        // get tasks from each outlook folder and populate model data
        $scope.backlogTasks = getTasksFromOutlook($scope.config.BACKLOG_FOLDER.NAME, $scope.config.BACKLOG_FOLDER.SORT, $scope.config.STATUS.NOT_STARTED.VALUE);
        $scope.inprogressTasks = getTasksFromOutlook($scope.config.INPROGRESS_FOLDER.NAME, $scope.config.INPROGRESS_FOLDER.SORT, $scope.config.STATUS.IN_PROGRESS.VALUE);
        $scope.nextTasks = getTasksFromOutlook($scope.config.NEXT_FOLDER.NAME, $scope.config.NEXT_FOLDER.SORT, $scope.config.STATUS.NOT_STARTED.VALUE);
        $scope.waitingTasks = getTasksFromOutlook($scope.config.WAITING_FOLDER.NAME, $scope.config.WAITING_FOLDER.SORT, $scope.config.STATUS.WAITING.VALUE);
        $scope.completedTasks = getTasksFromOutlook($scope.config.COMPLETED_FOLDER.NAME, $scope.config.COMPLETED_FOLDER.SORT, $scope.config.STATUS.COMPLETED.VALUE);

        // copy the lists as the initial filter    
        $scope.filteredBacklogTasks = $scope.backlogTasks;
        $scope.filteredInprogressTasks = $scope.inprogressTasks;
        $scope.filteredNextTasks = $scope.nextTasks;
        $scope.filteredWaitingTasks = $scope.waitingTasks;
        $scope.filteredCompletedTasks = $scope.completedTasks;

        // then apply the current filters for search and sensitivity
        $scope.applyFilters();

        // clean up Completed Tasks
        if ($scope.config.COMPLETED.ACTION == 'ARCHIVE' || $scope.config.COMPLETED.ACTION == 'DELETE') {
            var i;
            var tasks = $scope.completedTasks;
            var count = tasks.length;
            for (i = 0; i < count; i++) {
                var days = Date.daysBetween(tasks[i].completeddate, new Date());
                if (days > $scope.config.COMPLETED.AFTER_X_DAYS) {
                    if ($scope.config.COMPLETED.ACTION == 'ARCHIVE') {
                        $scope.archiveTask(tasks[i], $scope.completedTasks, $scope.filteredCompletedTasks);
                    }
                    if ($scope.config.COMPLETED.ACTION == 'DELETE') {
                        $scope.deleteTask(tasks[i], $scope.completedTasks, $scope.filteredCompletedTasks, false);
                    }
                };
            };
        };

        // move tasks with start date today to the Next folder
        if ($scope.config.AUTO_START_TASKS) {
            var i;
            var movedTask = false;
            var tasks = $scope.backlogTasks;
            var count = tasks.length;
            for (i = 0; i < count; i++) {
                if (tasks[i].startdate.getFullYear() != 4501) {
                    var seconds = Date.secondsBetween(tasks[i].startdate, new Date());
                    if (seconds >= 0) {
                        var taskitem = getTaskItem(tasks[i].entryID);
                        taskitem.Move(getTaskFolder($scope.config.NEXT_FOLDER.NAME));
                        movedTask = true;
                    }
                };
            };
            if (movedTask) {
                $scope.backlogTasks = getTasksFromOutlook($scope.config.BACKLOG_FOLDER.NAME, $scope.config.BACKLOG_FOLDER.SORT, $scope.config.STATUS.NOT_STARTED.VALUE);
                $scope.nextTasks = getTasksFromOutlook($scope.config.NEXT_FOLDER.NAME, $scope.config.NEXT_FOLDER.SORT, $scope.config.STATUS.NOT_STARTED.VALUE);
                $scope.filteredBacklogTasks = $scope.backlogTasks;
                $scope.filteredNextTasks = $scope.nextTasks;
            }
        };
    }

    $scope.applyFilters = function () {
        if ($scope.filter.search.length > 0) {
            $scope.filteredBacklogTasks = $filter('filter')($scope.backlogTasks, $scope.filter.search);
            $scope.filteredNextTasks = $filter('filter')($scope.nextTasks, $scope.filter.search);
            $scope.filteredInprogressTasks = $filter('filter')($scope.inprogressTasks, $scope.filter.search);
            $scope.filteredWaitingTasks = $filter('filter')($scope.waitingTasks, $scope.filter.search);
            $scope.filteredCompletedTasks = $filter('filter')($scope.completedTasks, $scope.filter.search);
        }
        else {
            $scope.filteredBacklogTasks = $scope.backlogTasks;
            $scope.filteredInprogressTasks = $scope.inprogressTasks;
            $scope.filteredNextTasks = $scope.nextTasks;
            $scope.filteredWaitingTasks = $scope.waitingTasks;
            $scope.filteredCompletedTasks = $scope.completedTasks;
        }

        // I think this can be written shorter, but for now it works
        var sensitivityFilter;
        if ($scope.filter.private != $scope.privacyFilter.all.value) {
            if ($scope.filter.private == $scope.privacyFilter.private.value) { sensitivityFilter = SENSITIVITY.olPrivate; }
            if ($scope.filter.private == $scope.privacyFilter.public.value) { sensitivityFilter = SENSITIVITY.olNormal; }
            $scope.filteredBacklogTasks = $filter('filter')($scope.filteredBacklogTasks, function (task) { return task.sensitivity == sensitivityFilter });
            $scope.filteredNextTasks = $filter('filter')($scope.filteredNextTasks, function (task) { return task.sensitivity == sensitivityFilter });
            $scope.filteredInprogressTasks = $filter('filter')($scope.filteredInprogressTasks, function (task) { return task.sensitivity == sensitivityFilter });
            $scope.filteredWaitingTasks = $filter('filter')($scope.filteredWaitingTasks, function (task) { return task.sensitivity == sensitivityFilter });
            $scope.filteredCompletedTasks = $filter('filter')($scope.filteredCompletedTasks, function (task) { return task.sensitivity == sensitivityFilter });
        }

        // filter backlog on start date
        if ($scope.config.BACKLOG_FOLDER.FILTER_ON_START_DATE) {
            $scope.filteredBacklogTasks = $filter('filter')($scope.filteredBacklogTasks, function (task) {
                if (task.startdate.getFullYear() != 4501) {
                    var days = Date.daysBetween(task.startdate, new Date());
                    return days >= 0;
                }
                else return true; // always show tasks not having start date
            });
        };

        // filter completed tasks if the HIDE options is configured
        if ($scope.config.COMPLETED.ACTION == 'HIDE') {
            $scope.filteredCompletedTasks = $filter('filter')($scope.filteredCompletedTasks, function (task) {
                var days = Date.daysBetween(task.completeddate, new Date());
                return days < $scope.config.COMPLETED.AFTER_X_DAYS;
            });
        }
    }


    // this is only a proof-of-concept single page report in a draft email for weekly report
    // it will be improved later on
    $scope.createReport = function () {
        var i, array = [];
        var mailItem, mailBody;
        mailItem = newMailItem();
        mailItem.Subject = "Status Report";
        mailItem.BodyFormat = 2;

        mailBody = "<style>";
        mailBody += "body { font-family: Calibri; font-size:11.0pt; } ";
        //mailBody += " h3 { font-size: 11pt; text-decoration: underline; } ";
        mailBody += " </style>";
        mailBody += "<body>";

        // COMPLETED ITEMS
        if ($scope.config.COMPLETED_FOLDER.REPORT.DISPLAY) {
            var tasks = getTaskFolder($scope.config.COMPLETED_FOLDER.NAME).Items.Restrict("[Complete] = true And Not ([Sensitivity] = 2)");
            tasks.Sort("[Importance][Status]", true);
            mailBody += "<h3>" + $scope.config.COMPLETED_FOLDER.TITLE + "</h3>";
            mailBody += "<ul>";
            var count = tasks.Count;
            for (i = 1; i <= count; i++) {
                mailBody += "<li>"
                if (tasks(i).Categories !== "") { mailBody += "[" + tasks(i).Categories + "] "; }
                mailBody += "<strong>" + tasks(i).Subject + "</strong>" + " - <i>" + taskStatusText(tasks(i).Status) + "</i>";
                if ($scope.config.COMPLETED_FOLDER.DISPLAY_PROPERTIES.TOTALWORK) { mailBody += " - " + tasks(i).TotalWork + " mn "; }
                if (tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
                if (tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
                var dueDate = new Date(tasks(i).DueDate);
                if (moment(dueDate).isValid && moment(dueDate).year() != 4501) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
                if (taskBodyNotes(tasks(i).Body, 10000)) { mailBody += "<br>" + "<font color=gray>" + taskBodyNotes(tasks(i).Body, 10000) + "</font>"; }
                mailBody += "</li>";
            }
            mailBody += "</ul>";
        }

        // INPROGRESS ITEMS
        if ($scope.config.INPROGRESS_FOLDER.REPORT.DISPLAY) {
            var tasks = getTaskFolder($scope.config.INPROGRESS_FOLDER.NAME).Items.Restrict("[Status] = 1 And Not ([Sensitivity] = 2)");
            tasks.Sort("[Importance][Status]", true);
            mailBody += "<h3>" + $scope.config.INPROGRESS_FOLDER.TITLE + "</h3>";
            mailBody += "<ul>";
            var count = tasks.Count;
            for (i = 1; i <= count; i++) {
                mailBody += "<li>"
                if (tasks(i).Categories !== "") { mailBody += "[" + tasks(i).Categories + "] "; }
                mailBody += "<strong>" + tasks(i).Subject + "</strong>" + " - <i>" + taskStatusText(tasks(i).Status) + "</i>";
                if ($scope.config.INPROGRESS_FOLDER.DISPLAY_PROPERTIES.TOTALWORK) { mailBody += " - " + tasks(i).TotalWork + " mn "; }
                if (tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
                if (tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
                var dueDate = new Date(tasks(i).DueDate);
                if (moment(dueDate).isValid && moment(dueDate).year() != 4501) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
                if (taskBodyNotes(tasks(i).Body, 10000)) { mailBody += "<br>" + "<font color=gray>" + taskBodyNotes(tasks(i).Body, 10000) + "</font>"; }
                mailBody += "</li>";
            }
            mailBody += "</ul>";
        }

        // NEXT ITEMS
        if ($scope.config.NEXT_FOLDER.REPORT.DISPLAY) {
            var tasks = getTaskFolder($scope.config.NEXT_FOLDER.NAME).Items.Restrict("[Status] = 0 And Not ([Sensitivity] = 2)");
            tasks.Sort("[Importance][Status]", true);
            mailBody += "<h3>" + $scope.config.NEXT_FOLDER.TITLE + "</h3>";
            mailBody += "<ul>";
            var count = tasks.Count;
            for (i = 1; i <= count; i++) {
                mailBody += "<li>"
                if (tasks(i).Categories !== "") { mailBody += "[" + tasks(i).Categories + "] "; }
                mailBody += "<strong>" + tasks(i).Subject + "</strong>" + " - <i>" + taskStatusText(tasks(i).Status) + "</i>";
                if ($scope.config.NEXT_FOLDER.DISPLAY_PROPERTIES.TOTALWORK) { mailBody += " - " + tasks(i).TotalWork + " mn "; }
                if (tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
                if (tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
                var dueDate = new Date(tasks(i).DueDate);
                if (moment(dueDate).isValid && moment(dueDate).year() != 4501) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
                if (taskBodyNotes(tasks(i).Body, 10000)) { mailBody += "<br>" + "<font color=gray>" + taskBodyNotes(tasks(i).Body, 10000) + "</font>"; }
                mailBody += "</li>";
            }
            mailBody += "</ul>";
        }

        // WAITING ITEMS
        if ($scope.config.WAITING_FOLDER.REPORT.DISPLAY) {
            var tasks = getTaskFolder($scope.config.WAITING_FOLDER.NAME).Items.Restrict("[Status] = 3 And Not ([Sensitivity] = 2)");
            tasks.Sort("[Importance][Status]", true);
            mailBody += "<h3>" + $scope.config.WAITING_FOLDER.TITLE + "</h3>";
            mailBody += "<ul>";
            var count = tasks.Count;
            for (i = 1; i <= count; i++) {
                mailBody += "<li>"
                if (tasks(i).Categories !== "") { mailBody += "[" + tasks(i).Categories + "] "; }
                mailBody += "<strong>" + tasks(i).Subject + "</strong>" + " - <i>" + taskStatusText(tasks(i).Status) + "</i>";
                if ($scope.config.WAITING_FOLDER.DISPLAY_PROPERTIES.TOTALWORK) { mailBody += " - " + tasks(i).TotalWork + " mn "; }
                if (tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
                if (tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
                var dueDate = new Date(tasks(i).DueDate);
                if (moment(dueDate).isValid && moment(dueDate).year() != 4501) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
                if (taskBodyNotes(tasks(i).Body, 10000)) { mailBody += "<br>" + "<font color=gray>" + taskBodyNotes(tasks(i).Body, 10000) + "</font>"; }
                mailBody += "</li>";
            }
            mailBody += "</ul>";
        }

        // BACKLOG ITEMS
        if ($scope.config.BACKLOG_FOLDER.REPORT.DISPLAY) {
            var tasks = getTaskFolder($scope.config.BACKLOG_FOLDER.NAME).Items.Restrict("[Status] = 0 And Not ([Sensitivity] = 2)");
            tasks.Sort("[Importance][Status]", true);
            mailBody += "<h3>" + $scope.config.BACKLOG_FOLDER.TITLE + "</h3>";
            mailBody += "<ul>";
            var count = tasks.Count;
            for (i = 1; i <= count; i++) {
                mailBody += "<li>"
                if (tasks(i).Categories !== "") { mailBody += "[" + tasks(i).Categories + "] "; }
                mailBody += "<strong>" + tasks(i).Subject + "</strong>" + " - <i>" + taskStatusText(tasks(i).Status) + "</i>";
                if ($scope.config.BACKLOG_FOLDER.DISPLAY_PROPERTIES.TOTALWORK) { mailBody += " - " + tasks(i).TotalWork + " mn "; }
                if (tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
                if (tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
                var dueDate = new Date(tasks(i).DueDate);
                if (moment(dueDate).isValid && moment(dueDate).year() != 4501) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
                if (taskBodyNotes(tasks(i).Body, 10000)) { mailBody += "<br>" + "<font color=gray>" + taskBodyNotes(tasks(i).Body, 10000) + "</font>"; }
                mailBody += "</li>";
            }
            mailBody += "</ul>";
        }

        mailBody += "</body>"

        // include report content to the mail body
        mailItem.HTMLBody = mailBody;

        // only display the draft email
        mailItem.Display();

    };

    var taskBodyNotes = function (str, limit) {
        // remove empty lines, cut off text if length > limit
        str = str.replace(/^(?=\n)$|^\s*|\s*$|\n\n+/gm, '');
        str = str.replace('\r\n', '<br>');
        if (str.length > limit) {
            str = str.substring(0, str.lastIndexOf(' ', limit));
            if (limit != 0) { str = str + "..." }
        };
        return str;
    };

    var taskStatusText = function (status) {
        if (status == $scope.config.STATUS.NOT_STARTED.VALUE) { return $scope.config.STATUS.NOT_STARTED.TEXT; }
        if (status == $scope.config.STATUS.IN_PROGRESS.VALUE) { return $scope.config.STATUS.IN_PROGRESS.TEXT; }
        if (status == $scope.config.STATUS.WAITING.VALUE) { return $scope.config.STATUS.WAITING.TEXT; }
        if (status == $scope.config.STATUS.COMPLETED.VALUE) { return $scope.config.STATUS.COMPLETED.TEXT; }
        return '';
    };

    // create a new task under target folder
    $scope.addTask = function (target) {
        // set the parent folder to target defined
        switch (target) {
            case 'backlog':
                var tasksfolder = getTaskFolder($scope.config.BACKLOG_FOLDER.NAME);
                break;
            case 'inprogress':
                var tasksfolder = getTaskFolder($scope.config.INPROGRESS_FOLDER.NAME);
                break;
            case 'next':
                var tasksfolder = getTaskFolder($scope.config.NEXT_FOLDER.NAME);
                break;
            case 'waiting':
                var tasksfolder = getTaskFolder($scope.config.WAITING_FOLDER.NAME);
                break;
        };
        // create a new task item object in outlook
        var taskitem = tasksfolder.Items.Add();

        // set sensitivity according to the current filter
        if ($scope.filter.private == $scope.privacyFilter.private.value) {
            taskitem.Sensitivity = SENSITIVITY.olPrivate;
        }

        // display outlook task item window
        taskitem.Display();

        if ($scope.config.AUTO_UPDATE) {
            saveState();

            // bind to taskitem write event on outlook and reload the page after the task is saved
            eval("function taskitem::Write (bStat) {window.location.reload();  return true;}");
        }

        // for anyone wondering about this weird double colon syntax:
        // Office is using IE11 to launch custom apps.
        // This syntax is used in IE to bind events. 
        //(https://msdn.microsoft.com/en-us/library/ms974564.aspx?f=255&MSPPError=-2147217396)
        //
        // by using eval we can avoid any error message until it is actually executed by Microsofts scripting engine
    }

    // opens up task item in outlook
    // refreshes the taskboard page when task item window closed
    $scope.editTask = function (item) {
        var taskitem = getTaskItem(item.entryID);
        taskitem.Display();
        if ($scope.config.AUTO_UPDATE) {
            saveState();
            // bind to taskitem write event on outlook and reload the page after the task is saved
            eval("function taskitem::Write (bStat) {window.location.reload(); return true;}");
            // bind to taskitem beforedelete event on outlook and reload the page after the task is deleted
            eval("function taskitem::BeforeDelete (bStat) {window.location.reload(); return true;}");
        }
    };

    // deletes the task item in both outlook and model data
    $scope.deleteTask = function (item, sourceArray, filteredSourceArray, bAskConfirmation) {
        var doDelete = true;
        if (bAskConfirmation) {
            doDelete = window.confirm('Are you absolutely sure you want to delete this item?');
        }
        if (doDelete) {
            // locate and delete the outlook task
            var taskitem = getTaskItem(item.entryID);
            taskitem.Delete();

            // locate and remove the item from the models
            removeItemFromArray(item, sourceArray);
            removeItemFromArray(item, filteredSourceArray);
        };
    };

    // moves the task item to the archive folder and marks it as complete
    // also removes it from the model data
    $scope.archiveTask = function (item, sourceArray, filteredSourceArray) {
        // locate the task in outlook namespace by using unique entry id
        var taskitem = getTaskItem(item.entryID);

        // move the task to the archive folder first (if it is not already in)
        var archivefolder = getTaskFolder($scope.config.ARCHIVE_FOLDER.NAME);
        if (taskitem.Parent.Name != archivefolder.Name) {
            taskitem = taskitem.Move(archivefolder);
        };

        // locate and remove the item from the models
        removeItemFromArray(item, sourceArray);
        removeItemFromArray(item, filteredSourceArray);
    };

    var removeItemFromArray = function (item, array) {
        var index = array.indexOf(item);
        if (index != -1) { array.splice(index, 1); }
    };

    // checks whether the task date is overdue or today
    // returns class based on the result
    $scope.isOverdue = function (strdate) {
        var dateobj = new Date(strdate).setHours(0, 0, 0, 0);
        var today = new Date().setHours(0, 0, 0, 0);
        return { 'task-overdue': dateobj < today, 'task-today': dateobj == today };
    };
    
    $scope.getFooterStyle = function (categories) {
        if ($scope.config.USE_CATEGORY_COLOR_FOOTERS) {
            if ((categories !== '') && $scope.config.USE_CATEGORY_COLORS) {
                // Copy category style
                if (categories.length == 1) {
                    if (categories[0] == undefined) return undefined;
                    return categories[0].style;
                }
                // Make multi-category tasks light gray
                else {
                    var lightGray = '#dfdfdf';
                    return { "background-color": lightGray, color: getContrastYIQ(lightGray) };
                }           
            }
        }
        return;
    };

    Date.daysBetween = function (date1, date2) {
        //Get 1 day in milliseconds
        var one_day = 1000 * 60 * 60 * 24;

        // Convert both dates to milliseconds
        var date1_ms = date1.getTime();
        var date2_ms = date2.getTime();

        // Calculate the difference in milliseconds
        var difference_ms = date2_ms - date1_ms;

        // Convert back to days and return
        return difference_ms / one_day;
    }

    Date.secondsBetween = function (date1, date2) {
        //Get 1 second in milliseconds
        var one_second = 1000;

        // Convert both dates to milliseconds
        var date1_ms = date1.getTime();
        var date2_ms = date2.getTime();

        // Calculate the difference in milliseconds
        var difference_ms = date2_ms - date1_ms;

        // Convert back to seconds and return
        return difference_ms / one_second;
    }

    const DEFAULT_CONFIG =  {
           "BACKLOG_FOLDER": {
                "ACTIVE": true,
                "NAME": "",
                "TITLE": "BACKLOG",
                "LIMIT": 0,
                "SORT": "duedate,-priority",
                "DISPLAY_PROPERTIES": {
                    "OWNER": false,
                    "PERCENT": false,
                    "TOTALWORK": false
                },
                "FILTER_ON_START_DATE": true,
                "REPORT": {
                    "DISPLAY": true
                }
            },
            "NEXT_FOLDER": {
                "ACTIVE": true,
                "NAME": "Kanban",
                "TITLE": "NEXT",
                "LIMIT": 20,
                "SORT": "duedate,-priority",
                "DISPLAY_PROPERTIES": {
                    "OWNER": false,
                    "PERCENT": false,
                    "TOTALWORK": false
                },
                "REPORT": {
                    "DISPLAY": true
                }
            },
            "INPROGRESS_FOLDER": {
                "ACTIVE": true,
                "NAME": "Kanban",
                "TITLE": "IN PROGRESS",
                "LIMIT": 5,
                "SORT": "-priority",
                "DISPLAY_PROPERTIES": {
                    "OWNER": false,
                    "PERCENT": false,
                    "TOTALWORK": false
                },
                "REPORT": {
                    "DISPLAY": true
                }
            },
            "WAITING_FOLDER": {
                "ACTIVE": true,
                "NAME": "Kanban",
                "TITLE": "WAITING",
                "LIMIT": 0,
                "SORT": "-priority",
                "DISPLAY_PROPERTIES": {
                    "OWNER": false,
                    "PERCENT": false,
                    "TOTALWORK": false
                },
                "REPORT": {
                    "DISPLAY": true
                }
            },
            "COMPLETED_FOLDER": {
                "ACTIVE": true,
                "NAME": "Kanban",
                "TITLE": "COMPLETED",
                "LIMIT": 0,
                "SORT": "-completeddate,-priority,subject",
                "DISPLAY_PROPERTIES": {
                    "OWNER": false,
                    "PERCENT": false,
                    "TOTALWORK": false
                },
                "REPORT": {
                    "DISPLAY": true
                },
                "EDITABLE": true
            },
            "ARCHIVE_FOLDER": {
                "NAME": "Completed"
            },
            "TASKNOTE_MAXLEN": 100,
            "DATE_FORMAT": "dd-MMM",
            "USE_CATEGORY_COLORS": true,
            "USE_CATEGORY_COLOR_FOOTERS": false,
            "SAVE_STATE": true,
            "STATUS": {
                "NOT_STARTED": {
                    "VALUE": 0,
                    "TEXT": "Not Started"
                },
                "IN_PROGRESS": {
                    "VALUE": 1,
                    "TEXT": "In Progress"
                },
                "WAITING": {
                    "VALUE": 3,
                    "TEXT": "Waiting For Someone Else"
                },
                "COMPLETED": {
                    "VALUE": 2,
                    "TEXT": "Completed"
                }
            },
            "COMPLETED": {
                "AFTER_X_DAYS": 7,
                "ACTION": "ARCHIVE"
            },
            "AUTO_UPDATE": true,
            "AUTO_START_TASKS": false,
            "VERSION": ""
        }
    
    var getState = function () {
        var state = { "private": 0, "search": "" }; // default state

        if ($scope.config.SAVE_STATE) {
            var stateRaw = getJournalItem(STATE_ID);
            if (stateRaw !== null){
                state = JSON.parse(stateRaw);
            }
        }

        // handle backwards compatibility
        if (state.private === true) state.private = $scope.privacyFilter.private.value;
        if (state.private === false) state.private = $scope.privacyFilter.public.value;

        $scope.orgState = state;
        $scope.filter = 
            {   private: state.private,
                search:  state.search         
            };
    }

    var saveState = function () {
        if ($scope.config.SAVE_STATE) {
            var currState = { "private": $scope.filter.private, "search": $scope.filter.search };
            if (DeepDiff.diff($scope.orgState, currState)) {
                saveJournalItem(STATE_ID, JSON.stringify(currState));
            }
        }
    }

    var getConfig = function () {
        // $scope.orgConfig = {};
        $scope.configRaw = getJournalItem(CONFIG_ID);
        if ($scope.configRaw !== null){
            try {
                $scope.config = JSON.parse(JSON.minify($scope.configRaw));
            }
            catch (e) {
                alert("I am afraid there is something wrong with the json structure of your configuration data. Please correct it.");
                $scope.switchToConfigMode();
                return;
            }

            // Newer versions of the app can have new config entries (or removed)
            try {
                var delta = DeepDiff.diff($scope.config, DEFAULT_CONFIG);
                if (delta) {
                    var isUpdated = false;
                    delta.forEach(function (change) {
                        if (change.kind === 'N' || change.kind === 'D') {
                            DeepDiff.applyChange($scope.config, DEFAULT_CONFIG, change);
                            isUpdated = true;
                        }
                    });
                    if (isUpdated) {
                        saveConfig();
                    }
                }
                    
            } catch (error) {
                alert(error)
            }
    
        }
        else {
            $scope.config = DEFAULT_CONFIG;
            saveConfig();
        }
    }

    var saveConfig = function () {
        saveJournalItem(CONFIG_ID, JSON.stringify($scope.config, null, 2));
    }

    var getVersion = function () {
        $http.get(VERSION_URL)
            .then(function(response) {
                $scope.version_number = response.data;
                $scope.version_number = $scope.version_number.replace(/\n|\r/g, "");
                checkVersion();
            });
    };

    var checkVersion = function () {
        if ($scope.config.VERSION != $scope.version_number) {
            if ($scope.config.VERSION == '') {
                $scope.config.VERSION = $scope.version_number;
            }
            else {
                $scope.display_message = true;
            }
            saveConfig();
        }
    };

    var getCategoryStyles = function (csvCategories) {

        const colorArray = [ 
            '#E7A1A2', '#F9BA89', '#F7DD8F', '#FCFA90', '#78D168', '#9FDCC9', '#C6D2B0', '#9DB7E8', '#B5A1E2', 
            '#daaec2', '#dad9dc', '#6b7994', '#bfbfbf', '#6f6f6f', '#4f4f4f', '#c11a25', '#e2620d', '#c79930', 
            '#b9b300', '#368f2b', '#329b7a', '#778b45', '#2858a5', '#5c3fa3', '#93446b'
        ];
    
        var getColor = function (category) {
            var c = outlookCategories.names.indexOf(category);
            var i = outlookCategories.colors[c];        
            if (i == -1) {
                return '#4f4f4f';
            }
            else {
                return colorArray[i-1];
            }
        }
        
        function getContrastYIQ(hexcolor) {
            if (hexcolor == undefined) {
                return 'black';
            }
            var r = parseInt(hexcolor.substr(1, 2), 16);
            var g = parseInt(hexcolor.substr(3, 2), 16);
            var b = parseInt(hexcolor.substr(5, 2), 16);
            var yiq = ((r * 299) + (g * 587) + (b * 114)) / 1000;
            return (yiq >= 128) ? 'black' : 'white';
        }
    
        var i;
        var catStyles = [];
        var categories = csvCategories.split(/[;,]+/);
        catStyles.length = categories.length;
        for (i = 0; i < categories.length; i++) {
            categories[i] = categories[i].trim();
            if (categories[i].length > 0) {
                if ($scope.config.USE_CATEGORY_COLORS) {
                    catStyles[i] = {
                        label: categories[i], style: { "background-color": getColor(categories[i]), color: getContrastYIQ(getColor(categories[i])) }
                    }
                }
                else {
                    catStyles[i] = {
                        label: categories[i], style: { color: "black" }
                    };
                }
            }
        }
        return catStyles;
    }
});