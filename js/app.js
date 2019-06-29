'use strict';

var tbApp = angular.module('taskboardApp', ['ui.sortable']);

tbApp.controller('taskboardController', function ($scope, $filter, $http) {

    var applMode;
    var outlookCategories;

    const APP_MODE = 0;
    const CONFIG_MODE = 1;
    const HELP_MODE = 2;

    const STATE_ID = "KanbanState";
    const CONFIG_ID = "KanbanConfig";
    const LOG_ID = "KanbanErrorLog";

    const BACKLOG = 0;
    const SPRINT = 1;
    const DOING = 2;
    const WAITING = 3;
    const DONE = 4;
    const SOMEDAY = 5;

    const MAX_LOG_ENTRIES = 500;

    $scope.includeConfig = true;
    $scope.includeState = true;
    $scope.includeLog = false;

    $scope.categories = ["<All Categories>", "<No Category>"];
    $scope.privacyFilter =
        {
            all: { text: "Both", value: "0" },
            private: { text: "Private", value: "1" },
            public: { text: "Work", value: "2" }
        };
    $scope.display_message = false;

    $scope.taskFolders = [
        { type: 0 },
        { type: 1 },
        { type: 2 },
        { type: 3 },
        { type: 4 }
    ];

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
        setUrls();

        $scope.isBrowserSupported = checkBrowser();
        if (!$scope.isBrowserSupported) {
            return;
        }

        $scope.switchToAppMode();

        // watch search filter and apply it
        $scope.$watchGroup(['filter.search', 'filter.private', 'filter.category'], function (newValues, oldValues) {
            var search = newValues[0];
            $scope.applyFilters();
            saveState();
        });

        // ui-sortable options and events
        $scope.sortableOptions = {
            connectWith: '.tasklist',
            items: 'li',
            opacity: 0.5,
            cursor: 'move',
            containment: 'document',

            stop: function (e, ui) {
                try {
                    // locate the target folder in outlook
                    // ui.item.sortable.droptarget[0].id represents the id of the target list
                    if (ui.item.sortable.droptarget) { // check if it is dropped on a valid target
                        if (($scope.config.INPROGRESS_FOLDER.LIMIT !== 0 && e.target.id !== ('folder-' + DOING) && ui.item.sortable.droptarget.attr('id') === ('folder-' + DOING) && $scope.taskFolders[DOING].tasks.length >= $scope.config.INPROGRESS_FOLDER.LIMIT) ||
                            ($scope.config.NEXT_FOLDER.LIMIT !== 0 && ('folder-' + SPRINT) && ui.item.sortable.droptarget.attr('id') === ('folder-' + SPRINT) && $scope.taskFolders[SPRINT].tasks.length >= $scope.config.NEXT_FOLDER.LIMIT) ||
                            ($scope.config.WAITING_FOLDER.LIMIT !== 0 && e.target.id !== ('folder-' + WAITING) && ui.item.sortable.droptarget.attr('id') === ('folder-' + WAITING) && $scope.taskFolders[WAITING].tasks.length >= $scope.config.WAITING_FOLDER.LIMIT)) {
                            writeLog('Drag and drop canceled because of limit reached. From ' + e.target.id + ' to ' + ui.item.sortable.droptarget.attr('id'));
                            alert('Sorry, you reached the defined limit of this folder')
                            $scope.initTasks();
                            ui.item.sortable.cancel();
                        } else {
                            //TODO dit kan korter
                            switch (ui.item.sortable.droptarget[0].id) {
                                case 'folder-' + BACKLOG:
                                    var tasksfolder = getTaskFolder($scope.config.BACKLOG_FOLDER.NAME);
                                    var newstatus = $scope.config.STATUS.NOT_STARTED.VALUE;
                                    break;
                                case 'folder-' + SPRINT:
                                    var tasksfolder = getTaskFolder($scope.config.NEXT_FOLDER.NAME);
                                    var newstatus = $scope.config.STATUS.NOT_STARTED.VALUE;
                                    break;
                                case 'folder-' + DOING:
                                    var tasksfolder = getTaskFolder($scope.config.INPROGRESS_FOLDER.NAME);
                                    var newstatus = $scope.config.STATUS.IN_PROGRESS.VALUE;
                                    break;
                                case 'folder-' + WAITING:
                                    var tasksfolder = getTaskFolder($scope.config.WAITING_FOLDER.NAME);
                                    var newstatus = $scope.config.STATUS.WAITING.VALUE;
                                    break;
                                case 'folder-' + DONE:
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
                        }
                    }
                } catch (error) {
                    writeLog('drag and drop: ' + error)
                }

            }
        };

        getConfig();
        getState();
        getVersion();
        pingUsage();

        outlookCategories = getOutlookCategories();
        outlookCategories.names.forEach(function (name) {
            $scope.categories.push(name);
        });
        $scope.categories = $scope.categories.sort();
        applyConfig();
        $scope.displayFolderCount = 0;
        $scope.taskFolders.forEach(function (folder) {
            if (folder.display) $scope.displayFolderCount++;
        });

        $scope.initTasks();
    };

    $scope.submitConfig = function () {
        try {
            saveConfig();
            $scope.init();
            $scope.switchToAppMode();
        } catch (error) {
            writeLog('submitConfig: ' + error)
        }
    }

    // borrowed from http://stackoverflow.com/a/30446887/942100
    var fieldSorter = function (fields) {
        try {
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
        } catch (error) {
            writeLog('fieldSorter: ' + error)
        }
    };

    var getTasksFromOutlook = function (path, sort, folderStatus) {
        try {
            var i, array = [];
            var tasks = getTaskItems(path);

            var count = tasks.Count;
            for (i = 1; i <= count; i++) {
                var task = tasks(i);
                if (task.Status == folderStatus || folderStatus == -1) {
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
        } catch (error) {
            writeLog('getTasksFromOutlook: ' + error)
        }
    };

    $scope.openOneNoteURL = function (url) {
        try {
            window.event.returnValue = false;
            if (navigator.msLaunchUri) {
                navigator.msLaunchUri(url);
            } else {
                window.open(url, "_blank").close();
            }
            return false;
        } catch (error) {
            writeLog('openOneNoteURL: ' + error)
        }
    }

    $scope.initTasks = function () {
        try {
            // get tasks from each outlook folder and populate model data
            $scope.taskFolders.forEach(function (taskFolder) {
                taskFolder.tasks = getTasksFromOutlook(taskFolder.name, taskFolder.sort, taskFolder.initialStatus);
                taskFolder.filteredTasks = taskFolder.tasks;
            });

            // then apply the current filters for search and sensitivity
            $scope.applyFilters();

            // clean up Completed Tasks
            if ($scope.config.COMPLETED.ACTION == 'ARCHIVE' || $scope.config.COMPLETED.ACTION == 'DELETE') {
                var i;
                var tasks = $scope.taskFolders[DONE].tasks;
                var count = tasks.length;
                for (i = 0; i < count; i++) {
                    try {
                        var days = Date.daysBetween(tasks[i].completeddate, new Date());
                        if (days > $scope.config.COMPLETED.AFTER_X_DAYS) {
                            if ($scope.config.COMPLETED.ACTION == 'ARCHIVE') {
                                $scope.archiveTask(tasks[i], $scope.taskFolders[DONE].tasks, $scope.taskFolders[DONE].filteredTasks);
                            }
                            if ($scope.config.COMPLETED.ACTION == 'DELETE') {
                                $scope.deleteTask(tasks[i], $scope.taskFolders[DONE].tasks, $scope.taskFolders[DONE].filteredTasks, false);
                            }
                        };
                    } catch (error) {
                        // ignore errors at this point. 
                    }
                };
            };

            // move tasks that do not have status New to the Next folder
            if (true) {
                var i;
                var movedTask = false;
                var tasks = $scope.taskFolders[BACKLOG].tasks;
                var count = tasks.length;
                for (i = 0; i < count; i++) {
                    if (tasks[i].status != $scope.config.STATUS.NOT_STARTED.TEXT) {
                        var taskitem = getTaskItem(tasks[i].entryID);
                        taskitem.Move(getTaskFolder($scope.taskFolders[SPRINT].name));
                        movedTask = true;
                    }
                };
                if (movedTask) {
                    // TODO: why read all the task when onlya few items are moved
                    // Read all tasks again
                    $scope.taskFolders.forEach(function (taskFolder) {
                        taskFolder.tasks = getTasksFromOutlook(taskFolder.name, taskFolder.sort, taskFolder.initialStatus);
                        taskFolder.filteredTasks = taskFolder.tasks;
                    });
                }
            }

            // move tasks with start date today to the Next folder
            if ($scope.config.AUTO_START_TASKS) {
                var i;
                var movedTask = false;
                var tasks = $scope.taskFolders[BACKLOG].tasks;
                var count = tasks.length;
                for (i = 0; i < count; i++) {
                    if (tasks[i].startdate.getFullYear() != 4501) {
                        var seconds = Date.secondsBetween(tasks[i].startdate, new Date());
                        if (seconds >= 0) {
                            var taskitem = getTaskItem(tasks[i].entryID);
                            taskitem.Move(getTaskFolder($scope.taskFolders[SPRINT].name));
                            movedTask = true;
                        }
                    };
                };
                if (movedTask) {
                    // TODO: why read all the task when onlya few items are moved
                    // Read all tasks again
                    $scope.taskFolders.forEach(function (taskFolder) {
                        taskFolder.tasks = getTasksFromOutlook(taskFolder.name, taskFolder.sort, taskFolder.initialStatus);
                        taskFolder.filteredTasks = taskFolder.tasks;
                    });
                }
            }

            // move tasks with past due date to the Next folder
            if ($scope.config.AUTO_START_DUE_TASKS) {
                var i;
                var movedTask = false;
                var tasks = $scope.taskFolders[BACKLOG].tasks;
                var count = tasks.length;
                for (i = 0; i < count; i++) {
                    if (tasks[i].duedate.getFullYear() != 4501) {
                        var seconds = Date.secondsBetween(tasks[i].duedate, new Date());
                        if (seconds >= 0) {
                            var taskitem = getTaskItem(tasks[i].entryID);
                            taskitem.Move(getTaskFolder($scope.taskFolders[SPRINT].name));
                            movedTask = true;
                        }
                    };
                };
                if (movedTask) {
                    // TODO: why read all the task when onlya few items are moved
                    // Read all tasks again
                    $scope.taskFolders.forEach(function (taskFolder) {
                        taskFolder.tasks = getTasksFromOutlook(taskFolder.name, taskFolder.sort, taskFolder.initialStatus);
                        taskFolder.filteredTasks = taskFolder.tasks;
                    });
                }
            }

            // move tasks with start date in future back to the Backlog folder
            if (true) {
                var i;
                var movedTask = false;
                var tasks = $scope.taskFolders[SPRINT].tasks;
                var count = tasks.length;
                for (i = 0; i < count; i++) {
                    if (tasks[i].startdate.getFullYear() != 4501) {
                        var seconds = Date.secondsBetween(new Date(), tasks[i].startdate);
                        if (seconds >= 0) {
                            var taskitem = getTaskItem(tasks[i].entryID);
                            taskitem.Move(getTaskFolder($scope.taskFolders[BACKLOG].name));
                            movedTask = true;
                        }
                    };
                };
                if (movedTask) {
                    // TODO: why read all the task when onlya few items are moved
                    // Read all tasks again
                    $scope.taskFolders.forEach(function (taskFolder) {
                        taskFolder.tasks = getTasksFromOutlook(taskFolder.name, taskFolder.sort, taskFolder.initialStatus);
                        taskFolder.filteredTasks = taskFolder.tasks;
                    });
                }
            }
        } catch (error) {
            writeLog('initTasks: ' + error)
        }
    };

    function var_dump(object, returnString) {
        var returning = '';
        for (var element in object) {
            var elem = object[element];
            if (typeof elem == 'object') {
                elem = var_dump(object[element], true);
            }
            returning += element + ': ' + elem + '\n';
        }
        if (returning == '') {
            returning = 'Empty object';
        }
        if (returnString === true) {
            return returning;
        }
        alert(returning);
    }

    $scope.applyFilters = function () {
        try {
            if ($scope.filter.search.length > 0) {
                $scope.taskFolders.forEach(function (taskFolder) {
                    taskFolder.filteredTasks = $filter('filter')(taskFolder.tasks, $scope.filter.search);
                });
            }
            else {
                $scope.taskFolders.forEach(function (taskFolder) {
                    taskFolder.filteredTasks = taskFolder.tasks;
                });
            }

            if ($scope.filter.category != "<All Categories>") {
                if ($scope.filter.category == "<No Category>") {
                    $scope.taskFolders.forEach(function (taskFolder) {
                        taskFolder.filteredTasks = $filter('filter')(taskFolder.filteredTasks, function (task) {
                            return task.categories == '';
                        });
                    });
                }
                else {
                    $scope.taskFolders.forEach(function (taskFolder) {
                        taskFolder.filteredTasks = $filter('filter')(taskFolder.filteredTasks, function (task) {
                            if (task.categories == '') {
                                return false;
                            }
                            else {
                                for (var i = 0; i < task.categories.length; i++) {
                                    var cat = task.categories[i];
                                    if (cat.label == $scope.filter.category) {
                                        return true;
                                    }
                                }
                                return false;
                            }
                        });
                    });
                }
            }

            // I think this can be written shorter, but for now it works
            var sensitivityFilter;
            if ($scope.filter.private != $scope.privacyFilter.all.value) {
                if ($scope.filter.private == $scope.privacyFilter.private.value) { sensitivityFilter = SENSITIVITY.olPrivate; }
                if ($scope.filter.private == $scope.privacyFilter.public.value) { sensitivityFilter = SENSITIVITY.olNormal; }
                $scope.taskFolders.forEach(function (taskFolder) {
                    taskFolder.filteredTasks = $filter('filter')(taskFolder.filteredTasks, function (task) { return task.sensitivity == sensitivityFilter });
                });
            }

            // filter on start date
            $scope.taskFolders.forEach(function (taskFolder) {
                if (taskFolder.filterOnStartDate === true) {
                    taskFolder.filteredTasks = $filter('filter')(taskFolder.filteredTasks, function (task) {
                        if (task.startdate.getFullYear() != 4501) {
                            var days = Date.daysBetween(task.startdate, new Date());
                            return days >= 0;
                        }
                        else return true; // always show tasks not having start date
                    });
                }
            });

            // filter completed tasks if the HIDE options is configured
            if ($scope.config.COMPLETED.ACTION == 'HIDE') {
                $scope.taskFolders[DONE].filteredTasks = $filter('filter')($scope.taskFolders[DONE].filteredTasks, function (task) {
                    var days = Date.daysBetween(task.completeddate, new Date());
                    return days < $scope.config.COMPLETED.AFTER_X_DAYS;
                });
            }
        } catch (error) {
            writeLog('applyFilters: ' + error)
        }
    }

    $scope.sendFeedback = function (includeConfig, includeState, includeLog) {
        try {
            var mailItem = newMailItem();
            mailItem.Subject = "JanBan version " + $scope.version + " Feedback (Outlook version: " + getOutlookVersion() + ")";
            mailItem.To = "janban@papasmurf.nl";
            mailItem.BodyFormat = 2;
            if (includeConfig) {
                mailItem.Attachments.Add(getPureJournalItem(CONFIG_ID));
            }
            if (includeState) {
                mailItem.Attachments.Add(getPureJournalItem(STATE_ID));
            }
            if (includeLog) {
                mailItem.Attachments.Add(getPureJournalItem(LOG_ID));
            }
            mailItem.Display();
        } catch (error) {
            writeLog('sendFeedback: ' + error)
        }
    }

    // this is only a proof-of-concept single page report in a draft email for weekly report
    // it will be improved later on
    $scope.createReport = function () {
        try {
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
        } catch (error) {
            writeLog('createReport: ' + error)
        }
    };

    var taskBodyNotes = function (str, limit) {
        try {
            // remove empty lines, cut off text if length > limit
            str = str.replace(/^(?=\n)$|^\s*|\s*$|\n\n+/gm, '');
            str = str.replace('\r\n', '<br>');
            if (str.length > limit) {
                str = str.substring(0, str.lastIndexOf(' ', limit));
                if (limit != 0) { str = str + "..." }
            };
            return str;
        } catch (error) {
            writeLog('taskBodyNotes: ' + error)
        }
    };

    var taskStatusText = function (status) {
        try {
            if (status == $scope.config.STATUS.NOT_STARTED.VALUE) { return $scope.config.STATUS.NOT_STARTED.TEXT; }
            if (status == $scope.config.STATUS.IN_PROGRESS.VALUE) { return $scope.config.STATUS.IN_PROGRESS.TEXT; }
            if (status == $scope.config.STATUS.WAITING.VALUE) { return $scope.config.STATUS.WAITING.TEXT; }
            if (status == $scope.config.STATUS.COMPLETED.VALUE) { return $scope.config.STATUS.COMPLETED.TEXT; }
            return '';
        } catch (error) {
            writeLog('taskStatusText: ' + error)
        }
    };

    // create a new task under target folder
    $scope.addTask = function (target) {
        try {
            // set the parent folder to target defined
            switch (target) {
                case BACKLOG:
                    var tasksfolder = getTaskFolder($scope.taskFolders[BACKLOG].name);
                    break;
                case SPRINT:
                    var tasksfolder = getTaskFolder($scope.taskFolders[SPRINT].name);
                    break;
                case DOING:
                    var tasksfolder = getTaskFolder($scope.taskFolders[DOING].name);
                    break;
                case WAITING:
                    var tasksfolder = getTaskFolder($scope.taskFolders[WAITING].name);
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
        } catch (error) {
            writeLog('addTask: ' + error)
        }
    }

    // opens up task item in outlook
    // refreshes the taskboard page when task item window closed
    $scope.editTask = function (item) {
        try {
            if (item.status == $scope.config.STATUS.COMPLETED.TEXT) return;
            var taskitem = getTaskItem(item.entryID);
            taskitem.Display();
            if ($scope.config.AUTO_UPDATE) {
                saveState();
                // bind to taskitem write event on outlook and reload the page after the task is saved
                eval("function taskitem::Write (bStat) {window.location.reload(); return true;}");
                // bind to taskitem beforedelete event on outlook and reload the page after the task is deleted
                eval("function taskitem::BeforeDelete (bStat) {window.location.reload(); return true;}");
            }
        } catch (error) {
            writeLog('editTask: ' + error)
        }
    };

    // deletes the task item in both outlook and model data
    $scope.deleteTask = function (item, sourceArray, filteredSourceArray, bAskConfirmation) {
        try {
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
        } catch (error) {
            writeLog('deleteTask: ' + error)
        }
    };

    // moves the task item to the archive folder and marks it as complete
    // also removes it from the model data
    $scope.archiveTask = function (item, sourceArray, filteredSourceArray) {
        try {
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
        } catch (error) {
            writeLog('archiveTask: ' + error)
        }
    };

    var removeItemFromArray = function (item, array) {
        try {
            var index = array.indexOf(item);
            if (index != -1) { array.splice(index, 1); }
        } catch (error) {
            writeLog('removeItemFromArray: ' + error)
        }
    };

    // checks whether the task date is overdue or today
    // returns class based on the result
    $scope.isOverdue = function (strdate) {
        try {
            var dateobj = new Date(strdate).setHours(0, 0, 0, 0);
            var today = new Date().setHours(0, 0, 0, 0);
            return { 'task-overdue': dateobj < today, 'task-today': dateobj == today };
        } catch (error) {
            writeLog('isOverdue: ' + error)
        }
    };

    $scope.getFooterStyle = function (categories) {
        try {
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
        } catch (error) {
            writeLog('getFooterStyle: ' + error)
        }
    };

    Date.daysBetween = function (date1, date2) {
        try {
            //Get 1 day in milliseconds
            var one_day = 1000 * 60 * 60 * 24;

            // Convert both dates to milliseconds
            var date1_ms = date1.getTime();
            var date2_ms = date2.getTime();

            // Calculate the difference in milliseconds
            var difference_ms = date2_ms - date1_ms;

            // Convert back to days and return
            return difference_ms / one_day;
        } catch (error) {
            writeLog('Date.daysbetween: ' + error)
        }
    }

    Date.secondsBetween = function (date1, date2) {
        try {
            //Get 1 second in milliseconds
            var one_second = 1000;

            // Convert both dates to milliseconds
            var date1_ms = date1.getTime();
            var date2_ms = date2.getTime();

            // Calculate the difference in milliseconds
            var difference_ms = date2_ms - date1_ms;

            // Convert back to seconds and return
            return difference_ms / one_second;
        } catch (error) {
            writeLog('Date.secondsBetween: ' + error)
        }
    }

    var applyConfig = function () {
        try {
            // $scope.taskFolders[SOMEDAY].type = BACKLOG;
            // $scope.taskFolders[SOMEDAY].initialStatus = $scope.config.STATUS.NOT_STARTED.VALUE;
            // $scope.taskFolders[SOMEDAY].display = $scope.config.BACKLOG_FOLDER.ACTIVE;
            // $scope.taskFolders[SOMEDAY].name = $scope.config.BACKLOG_FOLDER.NAME;
            // $scope.taskFolders[SOMEDAY].title = $scope.config.BACKLOG_FOLDER.TITLE;

            // $scope.taskFolders[SOMEDAY].limit = $scope.config.BACKLOG_FOLDER.LIMIT;
            // $scope.taskFolders[SOMEDAY].sort = $scope.config.BACKLOG_FOLDER.SORT;

            $scope.taskFolders[BACKLOG].type = BACKLOG;
            $scope.taskFolders[BACKLOG].initialStatus = $scope.config.STATUS.NOT_STARTED.VALUE;
            $scope.taskFolders[BACKLOG].initialStatus = -1;
            $scope.taskFolders[BACKLOG].display = $scope.config.BACKLOG_FOLDER.ACTIVE;
            $scope.taskFolders[BACKLOG].name = $scope.config.BACKLOG_FOLDER.NAME;
            $scope.taskFolders[BACKLOG].title = $scope.config.BACKLOG_FOLDER.TITLE;
            $scope.taskFolders[BACKLOG].limit = $scope.config.BACKLOG_FOLDER.LIMIT;
            $scope.taskFolders[BACKLOG].sort = $scope.config.BACKLOG_FOLDER.SORT;
            $scope.taskFolders[BACKLOG].displayOwner = $scope.config.BACKLOG_FOLDER.DISPLAY_PROPERTIES.OWNER;
            $scope.taskFolders[BACKLOG].displayPercent = $scope.config.BACKLOG_FOLDER.DISPLAY_PROPERTIES.PERCENT;
            $scope.taskFolders[BACKLOG].displayTotalWork = $scope.config.BACKLOG_FOLDER.DISPLAY_PROPERTIES.TOTALWORK;
            $scope.taskFolders[BACKLOG].filterOnStartDate = $scope.config.BACKLOG_FOLDER.FILTER_ON_START_DATE;
            $scope.taskFolders[BACKLOG].displayInReport = $scope.config.BACKLOG_FOLDER.REPORT.DISPLAY;
            $scope.taskFolders[BACKLOG].allowAdd = true;
            $scope.taskFolders[BACKLOG].allowEdit = true;

            $scope.taskFolders[SPRINT].type = SPRINT;
            $scope.taskFolders[SPRINT].initialStatus = $scope.config.STATUS.NOT_STARTED.VALUE;
            $scope.taskFolders[SPRINT].display = $scope.config.NEXT_FOLDER.ACTIVE;
            $scope.taskFolders[SPRINT].name = $scope.config.NEXT_FOLDER.NAME;
            $scope.taskFolders[SPRINT].title = $scope.config.NEXT_FOLDER.TITLE;
            $scope.taskFolders[SPRINT].limit = $scope.config.NEXT_FOLDER.LIMIT;
            $scope.taskFolders[SPRINT].sort = $scope.config.NEXT_FOLDER.SORT;
            $scope.taskFolders[SPRINT].displayOwner = $scope.config.NEXT_FOLDER.DISPLAY_PROPERTIES.OWNER;
            $scope.taskFolders[SPRINT].displayPercent = $scope.config.NEXT_FOLDER.DISPLAY_PROPERTIES.PERCENT;
            $scope.taskFolders[SPRINT].displayTotalWork = $scope.config.NEXT_FOLDER.DISPLAY_PROPERTIES.TOTALWORK;
            $scope.taskFolders[SPRINT].filterOnStartDate = $scope.config.NEXT_FOLDER.FILTER_ON_START_DATE;
            $scope.taskFolders[SPRINT].displayInReport = $scope.config.NEXT_FOLDER.REPORT.DISPLAY;
            $scope.taskFolders[SPRINT].allowAdd = true;
            $scope.taskFolders[SPRINT].allowEdit = true;

            $scope.taskFolders[DOING].type = DOING;
            $scope.taskFolders[DOING].initialStatus = $scope.config.STATUS.IN_PROGRESS.VALUE;
            $scope.taskFolders[DOING].display = $scope.config.INPROGRESS_FOLDER.ACTIVE;
            $scope.taskFolders[DOING].name = $scope.config.INPROGRESS_FOLDER.NAME;
            $scope.taskFolders[DOING].title = $scope.config.INPROGRESS_FOLDER.TITLE;
            $scope.taskFolders[DOING].limit = $scope.config.INPROGRESS_FOLDER.LIMIT;
            $scope.taskFolders[DOING].sort = $scope.config.INPROGRESS_FOLDER.SORT;
            $scope.taskFolders[DOING].displayOwner = $scope.config.INPROGRESS_FOLDER.DISPLAY_PROPERTIES.OWNER;
            $scope.taskFolders[DOING].displayPercent = $scope.config.INPROGRESS_FOLDER.DISPLAY_PROPERTIES.PERCENT;
            $scope.taskFolders[DOING].displayTotalWork = $scope.config.INPROGRESS_FOLDER.DISPLAY_PROPERTIES.TOTALWORK;
            $scope.taskFolders[DOING].filterOnStartDate = $scope.config.INPROGRESS_FOLDER.FILTER_ON_START_DATE;
            $scope.taskFolders[DOING].displayInReport = $scope.config.INPROGRESS_FOLDER.REPORT.DISPLAY;
            $scope.taskFolders[DOING].allowAdd = false;
            $scope.taskFolders[DOING].allowEdit = true;

            $scope.taskFolders[WAITING].type = WAITING;
            $scope.taskFolders[WAITING].initialStatus = $scope.config.STATUS.WAITING.VALUE;
            $scope.taskFolders[WAITING].display = $scope.config.WAITING_FOLDER.ACTIVE;
            $scope.taskFolders[WAITING].name = $scope.config.WAITING_FOLDER.NAME;
            $scope.taskFolders[WAITING].title = $scope.config.WAITING_FOLDER.TITLE;
            $scope.taskFolders[WAITING].limit = $scope.config.WAITING_FOLDER.LIMIT;
            $scope.taskFolders[WAITING].sort = $scope.config.WAITING_FOLDER.SORT;
            $scope.taskFolders[WAITING].displayOwner = $scope.config.WAITING_FOLDER.DISPLAY_PROPERTIES.OWNER;
            $scope.taskFolders[WAITING].displayPercent = $scope.config.WAITING_FOLDER.DISPLAY_PROPERTIES.PERCENT;
            $scope.taskFolders[WAITING].displayTotalWork = $scope.config.WAITING_FOLDER.DISPLAY_PROPERTIES.TOTALWORK;
            $scope.taskFolders[WAITING].filterOnStartDate = $scope.config.WAITING_FOLDER.FILTER_ON_START_DATE;
            $scope.taskFolders[WAITING].displayInReport = $scope.config.WAITING_FOLDER.REPORT.DISPLAY;
            $scope.taskFolders[WAITING].allowAdd = false;
            $scope.taskFolders[WAITING].allowEdit = true;

            $scope.taskFolders[DONE].type = DONE;
            $scope.taskFolders[DONE].initialStatus = $scope.config.STATUS.COMPLETED.VALUE;
            $scope.taskFolders[DONE].display = $scope.config.COMPLETED_FOLDER.ACTIVE;
            $scope.taskFolders[DONE].name = $scope.config.COMPLETED_FOLDER.NAME;
            $scope.taskFolders[DONE].title = $scope.config.COMPLETED_FOLDER.TITLE;
            $scope.taskFolders[DONE].limit = $scope.config.COMPLETED_FOLDER.LIMIT;
            $scope.taskFolders[DONE].sort = $scope.config.COMPLETED_FOLDER.SORT;
            $scope.taskFolders[DONE].displayOwner = $scope.config.COMPLETED_FOLDER.DISPLAY_PROPERTIES.OWNER;
            $scope.taskFolders[DONE].displayPercent = $scope.config.COMPLETED_FOLDER.DISPLAY_PROPERTIES.PERCENT;
            $scope.taskFolders[DONE].displayTotalWork = $scope.config.COMPLETED_FOLDER.DISPLAY_PROPERTIES.TOTALWORK;
            $scope.taskFolders[DONE].filterOnStartDate = $scope.config.COMPLETED_FOLDER.FILTER_ON_START_DATE;
            $scope.taskFolders[DONE].displayInReport = $scope.config.COMPLETED_FOLDER.REPORT.DISPLAY;
            $scope.taskFolders[DONE].allowAdd = false;
            $scope.taskFolders[DONE].allowEdit = false;
        } catch (error) {
            writeLog('applyConfig: ' + error)
        }
    };

    const DEFAULT_CONFIG = {
        "BACKLOG_FOLDER": {
            "TYPE": BACKLOG,
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
            "TYPE": "SPRINT",
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
            "FILTER_ON_START_DATE": undefined,
            "REPORT": {
                "DISPLAY": true
            }
        },
        "INPROGRESS_FOLDER": {
            "TYPE": "DOING",
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
            "FILTER_ON_START_DATE": undefined,
            "REPORT": {
                "DISPLAY": true
            }
        },
        "WAITING_FOLDER": {
            "TYPE": "WAITING",
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
            "FILTER_ON_START_DATE": undefined,
            "REPORT": {
                "DISPLAY": true
            }
        },
        "COMPLETED_FOLDER": {
            "TYPE": "DONE",
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
            "FILTER_ON_START_DATE": undefined,
            "REPORT": {
                "DISPLAY": true
            },
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
        "AUTO_START_TASKS": true,
        "AUTO_START_DUE_TASKS": false,
        "PING_BACK": true,
        "LAST_PING": new Date(2000, 1, 1),
        "LOG_ERRORS": false
    }

    var getState = function () {
        try {
            var state = { "private": "0", "search": "", "category": "<All Categories>" }; // default state

            if ($scope.config.SAVE_STATE) {
                var stateRaw = getJournalItem(STATE_ID);
                if (stateRaw !== null) {
                    state = JSON.parse(stateRaw);
                }
                else {
                    saveJournalItem(STATE_ID, JSON.stringify(state, null, 2));
                }
            }

            // handle backwards compatibility
            if (state.private === true) state.private = $scope.privacyFilter.private.value;
            if (state.private === false) state.private = $scope.privacyFilter.public.value;
            if (state.category === undefined) state.category = '<All Categories>';

            $scope.prevState = state;
            $scope.filter =
                {
                    private: state.private,
                    search: state.search,
                    category: state.category
                };

        } catch (error) {
            writeLog('getState: ' + error)
        }
    }

    var saveState = function () {
        try {
            if ($scope.config.SAVE_STATE) {
                var currState = { "private": $scope.filter.private, "search": $scope.filter.search, "category": $scope.filter.category };
                if (DeepDiff.diff($scope.prevState, currState)) {
                    saveJournalItem(STATE_ID, JSON.stringify(currState, null, 2));
                    $scope.prevState = currState;
                }
            }
        } catch (error) {
            writeLog('saveState: ' + error)
        }
    }

    var getConfig = function () {
        try {
            $scope.previousConfig = null;
            $scope.configRaw = getJournalItem(CONFIG_ID);
            if ($scope.configRaw !== null) {
                try {
                    $scope.config = JSON.parse(JSON.minify($scope.configRaw));
                }
                catch (e) {
                    alert("I am afraid there is something wrong with the json structure of your configuration data. Please correct it.");
                    writeLog('getConfig JSON parse error: ' + e)
                    $scope.switchToConfigMode();
                    return;
                }
                updateConfig();
                migrateConfig();
                $scope.includeLog = $scope.config.LOG_ERRORS;
            }
            else {
                $scope.config = DEFAULT_CONFIG;
                saveConfig();
            }
        } catch (error) {
            writeLog('getConfig: ' + error)
        }
    }

    var saveConfig = function () {
        try {
            saveJournalItem(CONFIG_ID, JSON.stringify($scope.config, null, 2));
            $scope.includeLog = $scope.config.LOG_ERRORS;
        } catch (error) {
            writeLog('saveConfig: ' + error)
        }
    }

    var updateConfig = function () {
        try {
            // Check for added or removed key entries in the config
            var delta = DeepDiff.diff($scope.config, DEFAULT_CONFIG);
            if (delta) {
                var isUpdated = false;
                $scope.previousConfig = $scope.config;
                delta.forEach(function (change) {
                    if (change.kind === 'N' || change.kind === 'D') {
                        DeepDiff.applyChange($scope.config, DEFAULT_CONFIG, change);
                        isUpdated = true;
                    }
                });
                if (isUpdated) {
                    saveConfig();
                    // as long as we need configraw...
                    $scope.configRaw = getJournalItem(CONFIG_ID);
                }
            }
        } catch (error) {
            alert("updateConfig: " + error)
        }
    }

    var migrateConfig = function () {
        try {
            var isChanged = false;
            // some older configs have no folder name for the Kanban task lanes
            if ($scope.config.NEXT_FOLDER.NAME == "") {
                $scope.config.NEXT_FOLDER.NAME = "Kanban";
                $scope.taskFolders[SPRINT].name = "Kanban";
                isChanged = true;
            }
            if ($scope.config.INPROGRESS_FOLDER.NAME == "") {
                $scope.config.INPROGRESS_FOLDER.NAME = "Kanban";
                $scope.taskFolders[DOING].name = "Kanban";
                isChanged = true;
            }
            if ($scope.config.WAITING_FOLDER.NAME == "") {
                $scope.config.WAITING_FOLDER.NAME = "Kanban";
                $scope.taskFolders[WAITING].name = "Kanban";
                isChanged = true;
            }
            if ($scope.config.COMPLETED_FOLDER.NAME == "") {
                $scope.config.COMPLETED_FOLDER.NAME = "Kanban";
                $scope.taskFolders[DONE].name = "Kanban";
                isChanged = true;
            }
            if (isChanged) {
                saveConfig();
                // as long as we need configraw...
                $scope.configRaw = getJournalItem(CONFIG_ID);
            }
        }
        catch (error) {
            writeLog('migrateConfig: ' + error)
        }
    }

    var writeLog = function (message) {
        try {
            var doLog = false;
            if ($scope.config == undefined) {
                doLog = true;
            }
            else {
                doLog = $scope.config.LOG_ERRORS;
            }
            if (doLog) {
                var now = new Date();
                var datetimeString = now.getFullYear() + '-' + now.getMonth() + '-' + now.getDate() + ' ' + now.getHours() + ':' + now.getMinutes();
                message = datetimeString + "  " + message;
                var logRaw = getJournalItem(LOG_ID);
                var log = [];
                if (logRaw !== null) {
                    log = JSON.parse(logRaw);
                }
                log.unshift(message);
                if (log.length > MAX_LOG_ENTRIES) {
                    log.pop();
                }
                saveJournalItem(LOG_ID, JSON.stringify(log, null, 2));
            }
        } catch (error) {
            alert('Unexpected writeLog error. Please make a screenprint and send it to janban@papasmurf.nl.\n\r ' + error)
        }
    }

    var setUrls = function () {
        // These variables are replaced in the build pipeline
        $scope.VERSION_URL = '#VERSION#';
        $scope.DOWNLOAD_URL = '#DOWNLOAD#';
        $scope.WHATSNEW_URL = '#WHATSNEW#';
        $scope.PING_URL = '#PINGBACK#';
        $scope.version = VERSION;

    }

    var pingUsage = function () {
        try {
            if ($scope.config.PING_BACK) {
                // monitor the usage of the app
                if (Date.daysBetween(new Date($scope.config.LAST_PING), new Date()) > 1) {
                    var url = $scope.PING_URL.replace('{{email}}', escape(getUserEmailAddress()));
                    url = url.replace('{{name}}', escape(getUserName()));
                    url = url.replace('{{version}}', escape($scope.version));
                    try {
                        $http.post(url, { headers: { 'Cache-Control': 'no-cache', 'Pragma': 'no-cache' } });
                    } catch (error) {
                        if (url.indexOf('jan.van.veldhuizen') > -1) {
                            alert(error)
                        }
                        // other end users shouldn't be bothered with issues when the post is not successful
                    }
                    $scope.config.LAST_PING = new Date();
                    saveConfig();
                }
            }
        } catch (error) {
            writeLog('pingUsage: ' + error)
        }
    };

    var getVersion = function () {
        try {
            $http.get($scope.VERSION_URL, { headers: { 'Cache-Control': 'no-cache', 'Pragma': 'no-cache' } })
                .then(function (response) {
                    $scope.version_number = response.data;
                    $scope.version_number = $scope.version_number.replace(/\n|\r/g, "");
                    checkVersion();
                });
        } catch (error) {
            writeLog('getVersion: ' + error)
        }
    };

    var checkVersion = function () {
        try {
            if ($scope.version != $scope.version_number) {
                $scope.display_message = true;
            }
        } catch (error) {
            writeLog('checkVersion: ' + error)
        }
    };

    var getCategoryStyles = function (csvCategories) {

        const colorArray = [
            '#E7A1A2', '#F9BA89', '#F7DD8F', '#FCFA90', '#78D168', '#9FDCC9', '#C6D2B0', '#9DB7E8', '#B5A1E2',
            '#daaec2', '#dad9dc', '#6b7994', '#bfbfbf', '#6f6f6f', '#4f4f4f', '#c11a25', '#e2620d', '#c79930',
            '#b9b300', '#368f2b', '#329b7a', '#778b45', '#2858a5', '#5c3fa3', '#93446b'
        ];

        var getColor = function (category) {
            try {
                var c = outlookCategories.names.indexOf(category);
                var i = outlookCategories.colors[c];
                if (i == -1) {
                    return '#4f4f4f';
                }
                else {
                    return colorArray[i - 1];
                }
            } catch (error) {
                writeLog('getColor: ' + error);
            }
        }

        try {
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
        } catch (error) {
            writeLog('getCategoryStyles: ' + error);
        }
    }

    function getContrastYIQ(hexcolor) {
        try {
            if (hexcolor == undefined) {
                return 'black';
            }
            var r = parseInt(hexcolor.substr(1, 2), 16);
            var g = parseInt(hexcolor.substr(3, 2), 16);
            var b = parseInt(hexcolor.substr(5, 2), 16);
            var yiq = ((r * 299) + (g * 587) + (b * 114)) / 1000;
            return (yiq >= 128) ? 'black' : 'white';
        } catch (error) {
            writeLog('getContrastYIQ: ' + message);
        }
    }
});