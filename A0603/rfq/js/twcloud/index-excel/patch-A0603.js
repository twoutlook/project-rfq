
//console.log("---patch-A0602.js---");
function getPatchA0603(k) {
    function get1() {
        var cells = [];
        cells.push(
                {sheet: 1, row: 106, col: 2, json: {data: "人工费/只 "}}, // 
                {sheet: 1, row: 83, col: 3, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
                {sheet: 1, row: 83, col: 4, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
                {sheet: 1, row: 83, col: 5, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
                {sheet: 1, row: 83, col: 6, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
                {sheet: 1, row: 83, col: 7, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
                {sheet: 1, row: 83, col: 8, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
                {sheet: 1, row: 94, col: 3, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
                {sheet: 1, row: 94, col: 4, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
                {sheet: 1, row: 94, col: 5, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
                {sheet: 1, row: 94, col: 6, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
                {sheet: 1, row: 94, col: 7, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
                {sheet: 1, row: 94, col: 8, json: styleInput({fm: "number", dfm: "0.0", data: "1"})},
        //
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

