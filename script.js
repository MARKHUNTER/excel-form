Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("submit").onclick = addData;
    }
});

async function addData() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // Get inputs
            const fName = document.getElementById("firstName").value;
            const lName = document.getElementById("lastName").value;
            const phone = document.getElementById("phone").value;
            const prog = document.getElementById("program").value;
            const fullName = `${fName} ${lName}`;

            // Check if sheet is protected and unprotect if necessary
            // sheet.getProtection().unprotect(""); 

            // Find last row
            const range = sheet.getRange("A:E").getUsedRangeOrNullObject();
            range.load("rowCount");
            await context.sync();

            let nextRow = 0;
            if (!range.isNullObject) {
                nextRow = range.rowCount;
            }

            // Write data
            const newRowRange = sheet.getRangeByIndexes(nextRow, 0, 1, 5);
            newRowRange.values = [[fName, lName, fullName, phone, prog]];
            
            await context.sync();
            console.log("Success!");

            // Clear form
            document.querySelectorAll("input").forEach(i => i.value = "");
        });
    } catch (error) {
        console.error("Error: " + error);
    }
}
