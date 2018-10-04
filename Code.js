function getActiveTaskListID(targetTitle) {
	var lists = Tasks.Tasklists.list().getItems();
	for (var i in lists) {
		if (targetTitle == lists[i].getTitle()) {
			return lists[i].getId();
		}
	}
	return null;
}

function getTasks() {
	// Load tasks from Works
	var id = getActiveTaskListID(PropertiesService.getScriptProperties().getProperty('TASKLIST_NAME'));

	if (!id) {
		Logger.log("Tasklist not found");
		return;
	} else {
		Logger.log("Tasklist found");
	}

	var optionArgs = {
		maxResults: 1000,
		showCompleted: true,
		showDeleted: false,
		showHidden: true
	};
	return Tasks.Tasks.list(id, optionArgs).getItems();
}

function composeCollectionFromTasks(tasks) {
	var tasksCollection = [];
	for (var i in tasks) {
		// A task line is composed of the following 6 cells:
		// title, status, Due Date, Completed Date, Completed Time, Notes
		taskline = {
			"title": tasks[i].getTitle(),
			"status": tasks[i].getStatus(),
			"due": tasks[i].getDue(),
			"completed": tasks[i].getCompleted(),
			"performance": 0,
			"notes": tasks[i].getNotes()
		};

		// Handle task with due day
		if (null != taskline["due"]) {
			dueDate = new Date(taskline["due"]);
			dueDate.setHours(0);
			dueDate.setMinutes(0);
			dueDate.setDate(dueDate.getDate()+1);
			taskline["due"] = dueDate;
		} else {
			taskline["due"] = "";
		}

		// Handle completed task
		if (null != taskline["completed"]) {
			completedDate = new Date(taskline["completed"]);
			taskline["completed"] = completedDate;
		} else {
			taskline["completed"] = "";
		}

		// Handle notes
		if (null == taskline["notes"]) {
			taskline["notes"] = "";
		}

		// Calculate performance.
		if (null != tasks[i].getCompleted() && null != tasks[i].getDue()) {
			dueDate = new Date(tasks[i].getDue());
			dueDate.setHours(0);
			dueDate.setMinutes(0);
			completedDate = new Date(tasks[i].getCompleted());
			completedDate.setHours(0);
			completedDate.setMinutes(0);
			taskline["performance"] = Math.round((completedDate.getTime() - dueDate.getTime())/(24*60*60*1000))-1;
		} else {
			taskline["performance"] = "";
		}

		tasksCollection.push(taskline);
		Logger.log(taskline);
	}
	return tasksCollection;
}

function refillSheet(tasks, sheet) {
	// Clean original sheet and refill
	var lastRow = sheet.getLastRow();
	if (lastRow > 1) {
		sheet.getRange("A2:E"+lastRow).clearContent();
	}
	if (tasks.length > 0) {
		tasksoutput = tasks.map(function(task) {
			return [task["title"], task["due"], task["completed"], task["performance"], task["notes"]];
		});
		sheet.getRange("A2:E"+(tasks.length+1)).setValues(tasksoutput);
	}
}

function task2sheet() {
	var tasks = getTasks();

	var tasksCollection = composeCollectionFromTasks(tasks);

	var completedTasks = tasksCollection.filter(function(value, index) {
		return value["status"] == "completed";
	});
	var ongoingTasks = tasksCollection.filter(function(value, index) {
		return value["status"] != "completed";
	});
	// set spreadsheet
	var sheets = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
	var ongoingSheet = sheets.getSheetByName("Ongoing");
	var completedSheet = sheets.getSheetByName("Completed");

	refillSheet(completedTasks, completedSheet);
	refillSheet(ongoingTasks, ongoingSheet);

	// clear completed tasks from TASKS app
	//Tasks.Tasks.clear(id);
}

function sendDailyReport() {
	var tasks = getTasks();

	var tasksCollection = composeCollectionFromTasks(tasks);

	var ongoingTasks = tasksCollection.filter(function(value, index) {
		return value["status"] != "completed";
	});

	tasksoutput = ongoingTasks.map(function(task) {
		return [task["title"], task["due"], task["notes"]];
	});
	html = HtmlService.createTemplateFromFile('Mail_Template');
	html.data = tasksoutput;
	MailApp.sendEmail({
		to: PropertiesService.getScriptProperties().getProperty('RECEIVER_EMAIL'),
		subject: "[TODO List] All ongoing tasks",
		htmlBody: html.evaluate().getContent()});
}
