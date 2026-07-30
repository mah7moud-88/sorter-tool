/* global Excel */


/**
 * جلب أسماء الأعمدة من الصف الأول
 */
export async function getExcelColumns() {

    return await Excel.run(async (context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();

        const usedRange = sheet.getUsedRange();

        usedRange.load([
            "values",
            "columnCount"
        ]);

        await context.sync();


        const columns = [];

        const headers = usedRange.values[0];


        for (let i = 0; i < usedRange.columnCount; i++) {

            const header = headers[i];


            if (header !== "" && header !== null) {

                columns.push({
                    name: String(header),
                    column: getColumnLetter(i + 1)
                });

            }

        }


        return columns;

    });
}



/**
 * تحويل رقم العمود إلى حرف Excel
 */
function getColumnLetter(number) {

    let letter = "";

    while (number > 0) {

        let remainder = (number - 1) % 26;

        letter =
            String.fromCharCode(65 + remainder) +
            letter;

        number =
            Math.floor((number - remainder) / 26);

    }

    return letter;

}




/**
 * ترتيب الحسابات وربط القيمة بالحساب
 */
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



        // الحسابات كنص للحفاظ على الصفر الأول

        correctRange.load("text");

        accountRange.load("text");

        valueRange.load("values");



        await context.sync();



        // الحسابات الصحيحة

        const correctAccounts = [];


        for (
            let i = 1;
            i < correctRange.text.length;
            i++
        ) {


            const account =
                correctRange.text[i][0].trim();



            if (account !== "") {

                correctAccounts.push(account);

            }

        }




        // خريطة الحساب -> القيمة

        const accountMap = new Map();



        for (
            let i = 1;
            i < accountRange.text.length;
            i++
        ) {


            const account =
                accountRange.text[i][0].trim();


            const value =
                valueRange.values[i][0];



            if (account !== "") {

                accountMap.set(
                    account,
                    value
                );

            }

        }




        // تجهيز الناتج

        const output = [];



        for (const account of correctAccounts) {


            output.push([

                account,

                accountMap.get(account) ?? ""

            ]);

        }




        // تنظيف القديم

        sheet
            .getRange("D2:F1000")
            .clear(
                Excel.ClearApplyTo.contents
            );





        // كتابة الحساب والقيمة

        if (output.length > 0) {


            const resultRange =
                sheet.getRange(
                    `D2:E${output.length + 1}`
                );


            resultRange.values = output;



            // الحفاظ على أصفار الحساب

            sheet
                .getRange(
                    `D2:D${output.length + 1}`
                )
                .numberFormat = [["@"]];

        }




        // مقارنة الحسابات

        if (output.length > 0) {


            const formulas = [];


            for (
                let i = 0;
                i < output.length;
                i++
            ) {


                const row = i + 2;


                formulas.push([
                    `=A${row}=D${row}`
                ]);

            }



            sheet
                .getRange(
                    `F2:F${output.length + 1}`
                )
                .formulas = formulas;

        }



        await context.sync();


    });

}
