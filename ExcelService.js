/* global Excel */

export async function sortExcelAccounts(
    correctColumn,
    accountColumn,
    valueColumn
) {
    await Excel.run(async (context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();

        const correctRange = sheet.getRange(
            `${correctColumn}1:${correctColumn}1000`
        );

        const accountRange = sheet.getRange(
            `${accountColumn}1:${accountColumn}1000`
        );

        const valueRange = sheet.getRange(
            `${valueColumn}1:${valueColumn}1000`
        );


        // الحسابات نقرأها حسب الشكل الظاهر
        correctRange.load("text");

        accountRange.load("text");

        // القيم نقرأها كقيم فعلية
        valueRange.load("values");


        await context.sync();


        // قراءة الحسابات الصحيحة (بدون العنوان)
        const correctAccounts = [];

        for (let i = 1; i < correctRange.text.length; i++) {

            const account = correctRange.text[i][0].trim();

            if (account !== "") {
                correctAccounts.push(account);
            }
        }


        // إنشاء خريطة الحساب -> القيمة
        const accountMap = new Map();


        for (let i = 1; i < accountRange.text.length; i++) {

            const account = accountRange.text[i][0].trim();

            const value = valueRange.values[i][0];


            if (account !== "") {

                accountMap.set(account, value);

            }
        }


        // تجهيز النتائج
        const output = [];


        for (const account of correctAccounts) {

            output.push([

                account,

                accountMap.has(account)
                    ? accountMap.get(account)
                    : ""

            ]);

        }



        // تنظيف النتائج القديمة
        sheet.getRange("D2:F1000")
            .clear(Excel.ClearApplyTo.contents);



        // كتابة النتائج
        if (output.length > 0) {

            const resultRange = sheet.getRange(
                `D2:E${output.length + 1}`
            );

            resultRange.values = output;


            // معادلة المقارنة
            const formulas = [];


            for (let i = 0; i < output.length; i++) {

                const row = i + 2;

                formulas.push([
                    `=IF(A${row}=D${row},TRUE,FALSE)`
                ]);
            }


            const fRange = sheet.getRange(
                `F2:F${formulas.length + 1}`
            );

            fRange.formulas = formulas;

        }


        await context.sync();

    });
}
