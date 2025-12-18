let dialog; // Global variable to hold the dialog object

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Automatically open the floating form when the add-in starts
        openFloatingForm();
    }
});

function openFloatingForm() {
    // Open the form as a floating dialog
    // Height and Width are percentages of the screen
    Office.context.ui.displayDialogAsync(
        'https://markhunter.github.io/excel-form/index.html', 
        { height: 45, width: 30, displayInIframe: true }, 
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Dialog failed: " + asyncResult.error.message);
                return;
            }
            dialog = asyncResult.value;
            // Listen for messages coming BACK from the floating form
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, processFormData);
        }
    );
}

async function processFormData(arg) {
    // 1. Receive data from the floating window
    const studentData = JSON.parse(arg.message);
    
    // 2. Close the floating window once data is received
    dialog.close();

    // 3. Write the data to Excel (same logic as before)
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange("A:E").getUsedRangeOrNullObject();
        range.load("rowCount");
        await context.sync();

        let nextRow = range.isNullObject ? 0 : range.rowCount;
        const newRow = sheet.getRangeByIndexes(nextRow, 0, 1, 5);
        newRow.values = [[
            studentData.fName, 
            studentData.lName, 
            `${studentData.fName} ${studentData.lName}`, 
            studentData.phone, 
            studentData.prog
        ]];
        await context.sync();
    });
}
