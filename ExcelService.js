/* global Excel, Office */

// ============================================
// تهيئة الإضافة
// ============================================
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Account Sorter Pro loaded successfully");
        
        // تحميل الشيتات والأعمدة
        loadSheetsAndColumns();
        
        // ربط زر الترتيب
        document.getElementById("sortButton").addEventListener("click", handleSort);
        
        // ربط تغيير الشيت
        document.getElementById("sheetSelect").addEventListener("change", loadColumns);
    }
});

// ============================================
// تحميل الشيتات
// ============================================
async function loadSheetsAndColumns() {
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const sheetSelect = document.getElementById("sheetSelect");
            sheetSelect.innerHTML = "";

            sheets.items.forEach((sheet) => {
                const option = document.createElement("option");
                option.value = sheet.name;
                option.textContent = sheet.name;
                sheetSelect.appendChild(option);
            });

            // اختيار الشيت النشط
            const activeSheet = context.workbook.worksheets.getActiveWorksheet();
            activeSheet.load("name");
            await context.sync();
            sheetSelect.value = activeSheet.name;

            // تحميل أعمدة الشيت النشط
            await loadColumns();
        });
    } catch (error) {
        updateStatus("خطأ في تحميل الشيتات: " + error.message, "error");
    }
}

// ============================================
// تحميل الأعمدة
// ============================================
async function loadColumns() {
    try {
        await Excel.run(async (context) => {
            const sheetName = document.getElementById("sheetSelect").value;
            const sheet = context.workbook.worksheets.getItem(sheetName);

            const usedRange = sheet.getUsedRange();
            usedRange.load("columnCount");
            await context.sync();

            const columnCount = usedRange.columnCount || 0;
            const columns = [];

            for (let i = 0; i < columnCount; i++) {
                columns.push(columnLetterFromIndex(i));
            }

            // ملء الـ selects
            fillSelect("correctColumn", columns);
            fillSelect("accountColumn", columns);
            fillSelect("valueColumn", columns);

            // ملء أعمدة الناتج
            fillSelect("outputAccount", columns);
            fillSelect("outputValue", columns);
        });
    } catch (error) {
        updateStatus("خطأ في تحميل الأعمدة: " + error.message, "error");
    }
}

// ============================================
// تحويل رقم العمود إلى حرف
// ============================================
function columnLetterFromIndex(index) {
    let letter = "";
    let i = index;
    while (i >= 0) {
        letter = String.fromCharCode((i % 26) + 65) + letter;
        i = Math.floor(i / 26) - 1;
    }
    return letter;
}

// ============================================
// ملء select بالخيارات
// ============================================
function fillSelect(id, options) {
    const select = document.getElementById(id);
    const currentValue = select.value;
    select.innerHTML = "";

    options.forEach((opt) => {
        const option = document.createElement("option");
        option.value = opt;
        option.textContent = opt;
        select.appendChild(option);
    });

    // استرجاع القيمة القديمة لو موجودة
    if (options.includes(currentValue)) {
        select.value = currentValue;
    }
}

// ============================================
// معالج زر الترتيب
// ============================================
async function handleSort() {
    const sortButton = document.getElementById("sortButton");
    sortButton.disabled = true;
    updateStatus("جاري الترتيب...", "loading");

    try {
        const correctColumn = document.getElementById("correctColumn").value;
        const accountColumn = document.getElementById("accountColumn").value;
        const valueColumn = document.getElementById("valueColumn").value;
        const outputAccount = document.getElementById("outputAccount").value;
        const outputValue = document.getElementById("outputValue").value;

        // التحقق من المدخلات
        if (!correctColumn || !accountColumn || !valueColumn) {
            updateStatus("الرجاء اختيار جميع الأعمدة المطلوبة", "error");
            sortButton.disabled = false;
            return;
        }

        await sortExcelAccounts(
            correctColumn,
            accountColumn,
            valueColumn,
            outputAccount,
            outputValue
        );

        updateStatus("✅ تم الترتيب بنجاح!", "success");
    } catch (error) {
        updateStatus("❌ خطأ: " + error.message, "error");
    } finally {
        sortButton.disabled = false;
    }
}

// ============================================
// تحديث حالة النتيجة
// ============================================
function updateStatus(message, type = "success") {
    const status = document.getElementById("status");
    status.textContent = message;
    status.className = type;
}

// ============================================
// الدالة الرئيسية للترتيب
// ============================================
export async function sortExcelAccounts(
    correctColumn,
    accountColumn,
    valueColumn,
    outputAccount,
    outputValue
) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        const usedRange = sheet.getUsedRange();
        usedRange.load("rowCount");
        await context.sync();

        const lastRow = usedRange.rowCount;

        if (lastRow < 1) {
            return;
        }

        // قراءة الأعمدة
        const correctRange = sheet.getRange(`${correctColumn}1:${correctColumn}${lastRow}`);
        const accountRange = sheet.getRange(`${accountColumn}1:${accountColumn}${lastRow}`);
        const valueRange = sheet.getRange(`${valueColumn}1:${valueColumn}${lastRow}`);

        correctRange.load("text");
        accountRange.load("text");
        valueRange.load("values");

        await context.sync();

        // قراءة الحسابات الصحيحة
        const correctAccounts = [];
        for (let i = 1; i < correctRange.text.length; i++) {
            const account = (correctRange.text[i][0] || "").trim();
            if (account !== "") {
                correctAccounts.push(account);
            }
        }

        // إنشاء Map
        const accountMap = new Map();
        for (let i = 1; i < accountRange.text.length; i++) {
            const account = (accountRange.text[i][0] || "").trim();
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
                accountMap.has(account) ? accountMap.get(account) : ""
            ]);
        }

        // تنظيف النتائج القديمة
        const outputRange = sheet.getRange(
            `${outputAccount}2:${outputValue}${Math.max(lastRow, correctAccounts.length + 1)}`
        );
        outputRange.clear(Excel.ClearApplyTo.contents);

        // كتابة النتائج
        if (output.length > 0) {
            const accountOutputRange = sheet.getRange(
                `${outputAccount}2:${outputAccount}${output.length + 1}`
            );
            accountOutputRange.numberFormat = [["@"]];

            const resultRange = sheet.getRange(
                `${outputAccount}2:${outputValue}${output.length + 1}`
            );
            resultRange.values = output;

            accountOutputRange.numberFormat = [["@"]];
        }

        // كتابة معادلات المقارنة
        if (output.length > 0) {
            const formulas = [];
            for (let i = 0; i < output.length; i++) {
                const row = i + 2;
                formulas.push([
                    `=TRIM(${correctColumn}${row}&"")=TRIM(${outputAccount}${row}&"")`
                ]);
            }

            // كتابة المعادلة في العمود اللي بعد الناتج
            const formulaColumn = columnLetterFromIndex(
                columnIndexFromLetter(outputValue) + 1
            );
            sheet.getRange(`${formulaColumn}2:${formulaColumn}${output.length + 1}`).formulas = formulas;
        }

        await context.sync();
    });
}

// ============================================
// تحويل حرف العمود إلى رقم
// ============================================
function columnIndexFromLetter(letter) {
    let index = 0;
    for (let i = 0; i < letter.length; i++) {
        index = index * 26 + (letter.charCodeAt(i) - 64);
    }
    return index - 1;
}
