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

        correctRange.load("values");
        accountRange.load("values");
        valueRange.load("values");

        await context.sync();

        // قراءة الحسابات الصحيحة
        const correctAccounts = [];

        for (let i = 1; i < correctRange.values.length; i++) {
            const account = correctRange.values[i][0];

            if (account !== "" && account !== null) {
                correctAccounts.push(String(account).trim());
            }
        }

        // إنشاء Map للحسابات والقيم
        const accountMap = new Map();

        for (let i = 1; i < accountRange.values.length; i++) {

            const account = String(accountRange.values[i][0] ?? "").trim();
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

            // معادلات المقارنة
            const formulas = [];

            for (let i = 0; i < output.length; i++) {

                const row = i + 2;

                formulas.push([
                    `=A${row}=D${row}`
                ]);
            }

            const fRange = sheet.getRange(
                `F2:F${output.length + 1}`
            );

            fRange.formulas = formulas;
        }

        await context.sync();
    });
}
