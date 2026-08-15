/* global Excel */

export async function sortExcelAccounts(
    correctColumn,
    accountColumn,
    valueColumn
) {
    await Excel.run(async (context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // ============================================
        // تحديد آخر صف مستخدم في الشيت تلقائيًا
        // ============================================
        const usedRange = sheet.getUsedRange();
        usedRange.load("rowCount");

        await context.sync();

        const lastRow = usedRange.rowCount;

        // لو الشيت فاضي تقريبًا
        if (lastRow < 1) {
            return;
        }

        // ============================================
        // قراءة الأعمدة المطلوبة حتى آخر صف مستخدم
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

        // الحسابات كنص للحفاظ على الصفر الأول
        correctRange.load("text");
        accountRange.load("text");

        // القيم كما هي
        valueRange.load("values");

        await context.sync();

        // ============================================
        // قراءة الحسابات الصحيحة
        // ============================================
        const correctAccounts = [];

        for (let i = 1; i < correctRange.text.length; i++) {

            const account = (correctRange.text[i][0] || "").trim();

            if (account !== "") {
                correctAccounts.push(account);
            }
        }

        // ============================================
        // إنشاء Map:
        // رقم الحساب -> القيمة
        // ============================================
        const accountMap = new Map();

        for (let i = 1; i < accountRange.text.length; i++) {

            const account = (accountRange.text[i][0] || "").trim();
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
        const outputRange = sheet.getRange(
            `D2:F${Math.max(lastRow, correctAccounts.length + 1)}`
        );

        outputRange.clear(Excel.ClearApplyTo.contents);

        // ============================================
        // كتابة النتائج
        // ============================================
        if (output.length > 0) {

            const resultRange = sheet.getRange(
                `D2:E${output.length + 1}`
            );

            // نخلي عمود الحساب Text قبل الكتابة
            const accountOutputRange = sheet.getRange(
                `D2:D${output.length + 1}`
            );

            accountOutputRange.numberFormat = [["@"]];

            resultRange.values = output;

            // تأكيد أن عمود الحساب Text
            accountOutputRange.numberFormat = [["@"]];
        }

        // ============================================
        // كتابة معادلات المقارنة
        // ============================================
        if (output.length > 0) {

            const formulas = [];

            for (let i = 0; i < output.length; i++) {

                const row = i + 2;

                formulas.push([
                    `=TRIM(${correctColumn}${row}&"")=TRIM(D${row}&"")`
                ]);
            }

            sheet.getRange(
                `F2:F${output.length + 1}`
            ).formulas = formulas;
        }

        await context.sync();
    });
}
