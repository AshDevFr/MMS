var EMAIL_SENT = "EMAIL_SENT";
var COLUMN_NAME_RESULT = "MERGE_STATUS"

function onOpen() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var items = [
        {name: 'Strat Merge', functionName: 'menuMerge'},
        null,
        {name: 'About', functionName: 'menuAbout'}
    ];
    doc.addMenu('MMS', items);
}

function menuAbout() {
    Browser.msgBox('Mail Merging System - Version 0.1');
}

function menuMerge() {
    var me = Session.getActiveUser().getEmail();
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getActiveSheet();
    var app = UiApp.createApplication()
        .setTitle('Mail Merging System')
        .setWidth(300)
        .setHeight(150);

    var panel = app.createVerticalPanel();
    var grid = app.createGrid(5, 2)
        .setWidth(300)
        .setHeight(150);

    var labelDraft = app.createLabel('Draft : ');
    var listBoxDraft = app.createListBox(false).setName('draftId');
    listBoxDraft.setVisibleItemCount(1);

    var drafts = GmailApp.getDraftMessages();
    for (i in drafts) {
        listBoxDraft.addItem(drafts[i].getSubject(), drafts[i].getId());
    }
    grid.setWidget(0, 0, labelDraft);
    grid.setWidget(0, 1, listBoxDraft);


    var labelSenderName = app.createLabel('Sender Name : ');
    var textBoxSenderName = app.createTextBox().setName("senderName");
    var self = ContactsApp.getContact(me);
    if (self) {
        textBoxSenderName.setValue(self.getFullName() || self.getGivenName());
    }
    grid.setWidget(1, 0, labelSenderName);
    grid.setWidget(1, 1, textBoxSenderName);


    var labelSenderMail = app.createLabel('Sender Mail : ');
    var listBoxSenderMail = app.createListBox(false).setName("senderMail");
    var aliases = GmailApp.getAliases();
    listBoxSenderMail.addItem(me);
    for (i in aliases) {
        listBoxSenderMail.addItem(aliases[i]);
    }

    grid.setWidget(2, 0, labelSenderMail);
    grid.setWidget(2, 1, listBoxSenderMail);


    var labelRecipientColumn = app.createLabel('Recipient Column : ');
    var listBoxRecipientColumn = app.createListBox(false).setName("recipientColumn");
    var columns = sheet.getDataRange().getValues()[0];
    for (i in columns) {
        listBoxRecipientColumn.addItem(columns[i], i);
    }

    grid.setWidget(3, 0, labelRecipientColumn);
    grid.setWidget(3, 1, listBoxRecipientColumn);


    var buttonCancel = app.createButton('Cancel');
    var handlerCancel = app.createServerHandler('clickCancel');
    buttonCancel.addClickHandler(handlerCancel);
    grid.setWidget(4, 0, buttonCancel);

    var buttonNext = app.createButton('Next');
    var handlerNext = app.createServerHandler('clickNext');
    handlerNext.addCallbackElement(listBoxDraft);
    handlerNext.addCallbackElement(textBoxSenderName);
    handlerNext.addCallbackElement(listBoxSenderMail);
    handlerNext.addCallbackElement(listBoxRecipientColumn);
    buttonNext.addClickHandler(handlerNext);
    grid.setWidget(4, 1, buttonNext);

    panel.add(grid);
    app.add(panel);
    doc.show(app);
}


function clickNext(eventInfo) {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var app = UiApp.getActiveApplication();
    var sheet = doc.getActiveSheet();

    // Logger.log('Selected Draft: ' + eventInfo.parameter.draftId);
    // Logger.log('Sender Name: ' + eventInfo.parameter.senderName);
    // Logger.log('Sender Mail: ' + eventInfo.parameter.senderMail);
    // Logger.log('Column Recipient: ' + eventInfo.parameter.recipientColumn);

    var mail = GmailApp.getMessageById(eventInfo.parameter.draftId);
    var body = mail.getBody();
    // Logger.log('Subject : ' + mail.getSubject());
    // Logger.log('Body : ' + mail.getBody());

    var datas = sheet.getDataRange().getValues();
    var colStatus = datas[0].indexOf(COLUMN_NAME_RESULT);
    if (colStatus == -1) {
        colStatus = sheet.getLastColumn();
        sheet.getRange(1, parseInt(colStatus, 10) + 1).setValue(COLUMN_NAME_RESULT);
    }

    for (i in datas) {
        var colNumber = parseInt(i, 10) + 1;
        var colStatusNumber = parseInt(colStatus, 10) + 1;
        if (i > 0 && sheet.getRange(colNumber, colStatusNumber).getValue() != EMAIL_SENT) {
            var text = body;
            for (j in datas[i]) {
                if (j != eventInfo.parameter.recipientColumn && j != colStatus) {
                    text = text.replace("&lt;&lt;" + datas[0][j] + "&gt;&gt;", datas[i][j]);
                }
            }
            sendEmail(datas[i][eventInfo.parameter.recipientColumn], eventInfo.parameter.senderMail, eventInfo.parameter.senderName, mail.getSubject(), text);
            sheet.getRange(colNumber, colStatusNumber).setValue(EMAIL_SENT);
        }
    }
    app.close();
    return app;
}

function clickCancel() {
    var app = UiApp.getActiveApplication();
    app.close();
    return app;
}

function sendEmail(to, from, fromName, subject, body) {
    var options = {};
    options.from = from;
    options.htmlBody = body;
    if (fromName && fromName != '') {
        options.name = fromName;
    }
    GmailApp.sendEmail(to, subject, '', options);
}