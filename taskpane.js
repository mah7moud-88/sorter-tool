/* global Excel, Office, document */

import { sortExcelAccounts } from "./ExcelService.js";


Office.onReady((info) => {

    if (info.host === Office.HostType.Excel) {


        const button =
            document.getElementById("sortButton");


        button.onclick = async () => {

            const status =
                document.getElementById("status");


            try {

                status.innerText =
                    "⏳ جاري الترتيب...";


                const correctColumn =
                    document.getElementById("correctColumn").value;


                const accountColumn =
                    document.getElementById("accountColumn").value;


                const valueColumn =
                    document.getElementById("valueColumn").value;


                await sortExcelAccounts(
                    correctColumn,
                    accountColumn,
                    valueColumn
                );


                status.innerText =
                    "✅ تم الترتيب بنجاح";


            }

            catch(error){

                console.error(error);


                status.innerText =
                    "❌ " + error.message;

            }

        };



        loadSheets();

        loadExcelData();



        document
        .getElementById("sheetSelect")
        .addEventListener(
            "change",
            loadSheets
        );


    }

});





// ===============================
// تحميل أسماء الشيتات
// ===============================

async function loadSheets(){


    try {


        await Excel.run(async(context)=>{


            const sheets =
                context.workbook.worksheets;


            sheets.load("items/name");


            await context.sync();



            const sheetSelect =
                document.getElementById("sheetSelect");



            sheetSelect.innerHTML = "";



            sheets.items.forEach(sheet=>{


                const option =
                    document.createElement("option");


                option.text =
                    sheet.name;


                option.value =
                    sheet.name;


                sheetSelect.appendChild(option);


            });


        });


    }

    catch(error){

        console.error(error);

    }

}





// ===============================
// تحميل الأعمدة
// ===============================

async function loadExcelData(){


    try{


        await Excel.run(async(context)=>{


            const sheet =
                context.workbook
                .worksheets
                .getActiveWorksheet();



            const range =
                sheet.getUsedRange();


            range.load("values");


            await context.sync();



            const data =
                range.values;



            if(!data || data.length===0)
                return;



            fillColumnLists(
                data[0]
            );


        });


    }

    catch(error){


        console.error(error);


        document.getElementById("status").innerText =
        "❌ خطأ في قراءة البيانات";


    }

}





// ===============================
// تعبئة الأعمدة
// ===============================

function fillColumnLists(columns){


    [
        "correctColumn",
        "accountColumn",
        "valueColumn"

    ].forEach(id=>{


        const select =
            document.getElementById(id);



        select.innerHTML="";



        columns.forEach((column,index)=>{


            const option =
                document.createElement("option");



            const letter =
                String.fromCharCode(
                    65 + index
                );



            option.text =
                `${letter} - ${column || "عمود"}`;


            option.value =
                letter;



            select.appendChild(option);


        });


    });


}