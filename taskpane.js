/* global Excel, Office, document */


import { sortExcelAccounts } from "./ExcelService";





Office.onReady((info) => {



    if (info.host === Office.HostType.Excel) {



        document
            .getElementById("sortButton")
            .onclick = sortExcelAccounts;




        loadSheets();



        loadExcelData();




        document
            .getElementById("sheetSelect")
            .addEventListener(
                "click",
                loadSheets
            );



    }



});









// ===============================
// تحميل أسماء الشيتات
// ===============================


async function loadSheets() {



    try {



        await Excel.run(async (context) => {



            const sheets =
                context.workbook.worksheets;



            sheets.load(
                "items/name"
            );



            await context.sync();






            const sheetSelect =
                document.getElementById(
                    "sheetSelect"
                );



            sheetSelect.innerHTML = "";






            sheets.items.forEach((item) => {



                const option =
                    document.createElement(
                        "option"
                    );



                option.text =
                    item.name;



                option.value =
                    item.name;



                sheetSelect.appendChild(
                    option
                );



            });




        });



    }

    catch(error){



        console.error(error);



    }



}









// ===============================
// تحميل بيانات الشيت والأعمدة
// ===============================


async function loadExcelData() {



    try {



        await Excel.run(async(context)=>{



            const sheet =
                context.workbook
                .worksheets
                .getActiveWorksheet();





            const usedRange =
                sheet.getUsedRange();




            usedRange.load(
                "values"
            );



            await context.sync();






            const data =
                usedRange.values;




            if(!data || data.length === 0){


                return;


            }






            const columns =
                data[0];



            console.log(
                "Columns:",
                columns
            );





            fillColumnLists(
                columns
            );




        });



    }


    catch(error){



        console.error(error);



        document.getElementById(
            "status"
        ).innerText =
        "❌ خطأ في قراءة الأعمدة";



    }



}









// ===============================
// تعبئة قوائم الأعمدة
// ===============================


function fillColumnLists(columns){



    const selects = [



        "correctColumn",

        "accountColumn",

        "valueColumn"



    ];





    selects.forEach((id)=>{



        const select =
            document.getElementById(
                id
            );



        if(!select){

            return;

        }



        select.innerHTML = "";






        columns.forEach((column,index)=>{



            const option =
                document.createElement(
                    "option"
                );



            const letter =
                String.fromCharCode(
                    65 + index
                );



            option.text =
                `${letter} - ${column || "عمود"}`;



            option.value =
                letter;



            select.appendChild(
                option
            );



        });



    });



}