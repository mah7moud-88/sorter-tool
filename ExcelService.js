/* global Excel */


export async function sortExcelAccounts() {


    await Excel.run(async (context) => {


        const sheet =
            context.workbook.worksheets.getActiveWorksheet();



        const correctRange =
            sheet.getRange("A1:A1000");


        const accountRange =
            sheet.getRange("B1:B1000");


        const valueRange =
            sheet.getRange("C1:C1000");



        correctRange.load("values");

        accountRange.load("values");

        valueRange.load("values");



        await context.sync();



        const correctAccounts =
            correctRange.values
                .flat()
                .filter(x => x !== "" && x !== null)
                .map(String);



        const accounts =
            accountRange.values
                .flat()
                .filter(x => x !== "" && x !== null)
                .map(String);



        const values =
            valueRange.values
                .flat();



        const accountMap = {};



        accounts.forEach((account, index) => {


            accountMap[account] =
                values[index];


        });



        const output = [];



        correctAccounts.forEach(account => {


            output.push([
                account,
                accountMap[account] ?? ""
            ]);


        });



        const resultRange =
            sheet.getRange(
                `D1:E${output.length}`
            );



        resultRange.values =
            output;



        await context.sync();


    });


}