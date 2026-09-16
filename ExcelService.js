/* global Excel */

export async function sortExcelAccounts(
    correctColumn,
    accountColumn,
    valueColumn
) {
    await Excel.run(async (context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // ============================================
        // تحديد آخر صف مستخدم
        // ============================================
        const usedRange = sheet.getUsedRange();
        usedRange.load("rowCount");

        await context.sync();

        const lastRow = usedRange.rowCount;

        if (lastRow < 1) {
            return;
        }

        // ============================================
        // قراءة الأعمدة المطلوبة
        // ============================================
        const correctRange = sheet.getRange(
            `${correctColumn}1:${correctColumn}${lastRow}`
        );

        const accountRange = sheet.getRange(
            `${accountColumn}1:${accountColumn}${lastRow}`
        );

        const valueRange = sheet.getRange(
            `${valueColumn}1:${valueColumn}${lastRow}`
        );

        correctRange.load("text");
        accountRange.load("text");
        valueRange.load("values");

        await context.sync();

        // ============================================
        // قراءة الحسابات الصحيحة
        // ============================================
        const correctAccounts = [];

        for (let i = 1; i < correctRange.text.length; i++) {

            const account = String(
                correctRange.text[i][0] || ""
            ).trim();

            if (account !== "") {
                correctAccounts.push(account);
            }
        }

        // ============================================
        // إنشاء Map
        // الحساب -> القيمة
        // ============================================
        const accountMap = new Map();

        for (let i = 1; i < accountRange.text.length; i++) {

            const account = String(
                accountRange.text[i][0] || ""
            ).trim();

            const value = valueRange.values[i][0];

            if (account !== "") {
                accountMap.set(account, value);
            }
        }

        // ============================================
        // تجهيز النتائج
        // ============================================
        const output = [];

        for (const account of correctAccounts) {

            output.push([
                account,
                accountMap.has(account)
                    ? accountMap.get(account)
                    : ""
            ]);
        }

        // ============================================
        // تنظيف النتائج القديمة
        // ============================================
        const clearEndRow = Math.max(
            lastRow,
            correctAccounts.length + 1
        );

        sheet.getRange(
            `D2:F${clearEndRow}`
        ).clear(Excel.ClearApplyTo.contents);

        // ============================================
        // كتابة النتائج
        // ============================================
        if (output.length > 0) {

            const resultRange = sheet.getRange(
                `D2:E${output.length + 1}`
            );

            resultRange.values = output;
        }

        // ============================================
        // التحقق
        // ============================================
        if (output.length > 0) {

            const formulas = [];

            for (let i = 0; i < output.length; i++) {

                const row = i + 2;

                /*
                 * VALUE يحول:
                 *
                 * 001234 -> 1234
                 * 000123 -> 123
                 * 1234   -> 1234
                 *
                 * وبالتالي الأصفار الأولى لا تؤثر على المقارنة.
                 */

                formulas.push([
                    `=IF(OR(${correctColumn}${row}="",D${row}=""),FALSE,IFERROR(VALUE(${correctColumn}${row})=VALUE(D${row}),FALSE))`
                ]);
            }

            sheet.getRange(
                `F2:F${output.length + 1}`
            ).formulas = formulas;
        }

        await context.sync();
    });
}
