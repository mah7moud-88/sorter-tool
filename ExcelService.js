/* global Excel */


export async function sortExcelAccounts(
    correctColumn,
    accountColumn,
    valueColumn,
    outputAccount = "D",
    outputValue = "E"
) {


    await Excel.run(async (context) => {


        const sheet =
            context.workbook.worksheets.getActiveWorksheet();



        // قراءة العناوين

        const headerRange =
            sheet.getRange(
                `${valueColumn}1`
            );


        headerRange.load("values");


        await context.sync();



        // وضع عنوان القيمة في E1

        const outputHeader =
            sheet.getRange(
                `${outputValue}1`
            );


        outputHeader.values = [
            [
                headerRange.values[0][0]
            ]
        ];





        // قراءة البيانات بداية من الصف الثاني

        const correctRange =
            sheet.getRange(
                `${correctColumn}2:${correctColumn}1000`
            );


        const accountRange =
            sheet.getRange(
                `${accountColumn}2:${accountColumn}1000`
            );


        const valueRange =
            sheet.getRange(
                `${valueColumn}2:${valueColumn}1000`
            );



        correctRange.load("values");

        accountRange.load("values");

        valueRange.load("values");



        await context.sync();





        const correctAccounts =
            correctRange.values
            .flat()
            .filter(
                x => x !== "" && x !== null
            )
            .map(
                x => String(x).trim()
            );




        const accounts =
            accountRange.values
            .flat()
            .filter(
                x => x !== "" && x !== null
            )
            .map(
                x => String(x).trim()
            );




        const values =
            valueRange.values
            .flat();







        // ربط الحساب بالقيمة

        const accountMap = {};



        accounts.forEach((account, index) => {


            accountMap[account] =
                values[index];


        });








        const resultAccounts = [];

        const resultValues = [];





        // ترتيب الحسابات حسب العمود المرجعي

        correctAccounts.forEach(account => {



            resultAccounts.push([
                account
            ]);



            resultValues.push([
                accountMap[account] ?? ""
            ]);



        });








        // كتابة الحسابات في D من الصف الثاني

        const accountOutputRange =
            sheet.getRange(
                `${outputAccount}2:${outputAccount}${resultAccounts.length + 1}`
            );


        accountOutputRange.values =
            resultAccounts;






        // كتابة القيم في E من الصف الثاني

        const valueOutputRange =
            sheet.getRange(
                `${outputValue}2:${outputValue}${resultValues.length + 1}`
            );


        valueOutputRange.values =
            resultValues;





        await context.sync();



    });


}
