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




        const correctAccounts =
            correctRange.values
            .flat()
            .filter(
                x => x !== "" && x !== null
            )
            .map(String);




        const accounts =
            accountRange.values
            .flat()
            .filter(
                x => x !== "" && x !== null
            )
            .map(String);




        const values =
            valueRange.values
            .flat();





        const accountMap = {};



        accounts.forEach((account,index)=>{


            accountMap[account] =
                values[index];


        });






        const output = [];



        correctAccounts.forEach(account=>{


            output.push([

                account,

                accountMap[account] ?? ""

            ]);


        });







        // الناتج يبدأ من العمود التالي بعد الأعمدة المستخدمة
        const resultRange =
            sheet.getRange(
                `D1:E${output.length}`
            );



        resultRange.values =
            output;



        await context.sync();



    });


}