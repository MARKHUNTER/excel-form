Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("submit").onclick = addStudent;
    }
});

async function addStudent() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            const fName = document.getElementById("firstName").value;
            const lName = document.getElementById("lastName").value;
            const phone = document.getElementById("phone").value;
            const prog = document.getElementById("program").value;
            const fullName = `${fName} ${lName}`;

            const range = sheet.getRange("A:E").getUsedRangeOrNullObject();
            range.load("rowCount");
            await context.sync();

            let nextRow = range.isNullObject ? 0 : range.rowCount;

            const newRow = sheet.getRangeByIndexes(nextRow, 0, 1, 5);
            newRow.values = [[fName, lName, fullName, phone, prog]];
            
            await context.sync();
            
            // Clear inputs
            ["firstName", "lastName", "phone", "program"].forEach(id => {
                document.getElementById(id).value = "";
            });
        });
    } catch (error) {
        console.error(error);
    }
}
