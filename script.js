let dialog;

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Automatically open the floating dialog when the add-in is opened
        openDialog();
    }
});

function openDialog() {
    // Opens index.html as a floating window
    Office.context.ui.displayDialogAsync(
        'https://markhunter.github.io/excel-form/index.html',
        { height: 50, width: 30, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
                return;
            }
            dialog = asyncResult.value;
            // Listen for the "Submit" signal from the floating window
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
        }
    );
}

async function processMessage(arg) {
    // Parse the data sent from the floating form
    const student = JSON.parse(arg.message);
    
    // Close the floating window after submission
    dialog.close();

    // Write the data to the active Excel sheet
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange("A:E").getUsedRangeOrNullObject();
        range.load("rowCount");
        await context.sync();

        let nextRow = range.isNullObject ? 0 : range.rowCount;
        const newRow = sheet.getRangeByIndexes(nextRow, 0, 1, 5);
        
        newRow.values = [[
            student.fName, 
            student.lName, 
            student.fName + " " + student.lName, 
            student.phone, 
            student.prog
        ]];
        
        await context.sync();
    });
}
