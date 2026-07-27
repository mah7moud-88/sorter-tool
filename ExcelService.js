/* global Excel */

export async function sortExcelAccounts(
    correctColumn,
    accountColumn,
    valueColumn
) {

    await Excel.run(async (context) => {

        const sheet =
            context.workbook.worksheets.getActiveWorksheet();


        const correctRange =
            sheet.getRange(
                `${correctColumn}1:${correctColumn}1000`
            );


        const accountRange =
            sheet.getRange(
                `${accountColumn}1:${accountColumn}1000`
            );


        const valueRange =
            sheet.getRange(
                `${valueColumn}1:${valueColumn}1000`
            );


        correctRange.load("values");
        accountRange.load("values");
        valueRange.load("values");


        await context.sync();



        // تجاهل أول صف (عنوان رقم الحساب)
        const correctAccounts =
            correctRange.values
                .flat()
                .slice(1)
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



        // تنظيف النتائج القديمة فقط بداية من الصف الثاني
        // حتى لا يتم مسح العناوين في D1 و E1 و F1
        sheet.getRange("D2:F1000")
            .clear(Excel.ClearApplyTo.contents);



        // كتابة النتائج في D و E بداية من الصف الثاني
        const resultRange =
            sheet.getRange(
                `D2:E${output.length + 1}`
            );


        resultRange.values = output;



        // وضع معادلة المقارنة في F بداية من الصف الثاني
        const formulas = [];


        for (let i = 0; i < output.length; i++) {

            const row = i + 2;


            formulas.push([

                `=IF(A${row}=D${row},TRUE,FALSE)`

            ]);

        }



        const fRange =
            sheet.getRange(
                `F2:F${formulas.length + 1}`
            );


        fRange.formulas = formulas;



        await context.sync();


    });

}
