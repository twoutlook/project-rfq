
console.log("---patch-A0602.js---");
function getPatchCellA0602(k) {
    if (k === 1) {
        return getPatchCellsA0602_1();
    }
//    if (k === 2) {
//        return getPatchCellsA0601_2();
//    }
//    if (k === 3) {
//        return getPatchCellsA0601_3();
//    }

}

function getPatchCellsA0602_1() {

//    function styleSubTotal(source2) {
//        var subtotal = {bgc: colorStepEnd, fm: "money|¥|2|none", dsd: "ed", cal: true};
//        return    mergeJSON(subtotal, source2);
//    }
    var cells = [];
    cells.push(
            {sheet: 1, row: 86, col: 3, json: ddl086}, // 
            {sheet: 1, row: 86, col: 4, json: ddl086}, // 
            {sheet: 1, row: 86, col: 5, json: ddl086}, // 
            {sheet: 1, row: 86, col: 6, json: ddl086}, // 
            {sheet: 1, row: 86, col: 7, json: ddl086}, // 
            {sheet: 1, row: 86, col: 8, json: ddl086}, // 
    //
    //
    // LAST ONE ============================================================================================
            {sheet: 1, row: 1, col: 1, json: {data: ""}}
    );
    return cells;
}
