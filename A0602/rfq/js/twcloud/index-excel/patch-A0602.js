
console.log("---patch-A0602.js---");
function getPatchCellsA0602(k) {
    function get1() {
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


    if (k === 1) {
        return get1();
    }
//    if (k === 2) {
//        return getPatchCellsA0601_2();
//    }
//    if (k === 3) {
//        return getPatchCellsA0601_3();
//    }

}

