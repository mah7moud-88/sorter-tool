export function sortAccounts(reference, source) {


    const result = [];


    const sourceMap = new Map();



    source.forEach(value => {

        sourceMap.set(
            value.trim(),
            value
        );

    });



    reference.forEach(account => {


        const key = account.trim();



        if (sourceMap.has(key)) {


            result.push(
                sourceMap.get(key)
            );


            sourceMap.delete(key);

        }


    });



    // إضافة الحسابات غير الموجودة في الترتيب الصحيح في النهاية

    sourceMap.forEach(value => {

        result.push(value);

    });



    return result;

}