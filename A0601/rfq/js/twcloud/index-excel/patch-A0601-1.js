
function func_patch_A0601_1_cells() {
    var colorStep = "#A9BCF5";
    var colorStepEnd = "#E6E6E6";
    var colorSect = "#837E7C"; //bgc: colorSect, fm: "money|¥|2|none", dsd: "ed", cal: true
    var colorDdl = "#F9E79F"; //#82E0AA  
    var colorInput = "#F4D03F"; // 
    var colorVersion = "#98AFC7"; // 
    function styleSubTotal(source2) {
        var subtotal = {bgc: colorStepEnd, fm: "money|¥|2|none", dsd: "ed", cal: true};
        return    mergeJSON(subtotal, source2);
    }
    var patch_A0601_1_cells = [];
    patch_A0601_1_cells.push(
            {sheet: 1, row: 113, col: 2, json: {data: '合計（含稅）：'}},
            {sheet: 1, row: 113, col: 3, json: styleSubTotal({data: '=C111*1.17'})},
            {sheet: 1, row: 113, col: 4, json: styleSubTotal({data: '=D111*1.17'})},
            {sheet: 1, row: 113, col: 5, json: styleSubTotal({data: '=E111*1.17'})},
            {sheet: 1, row: 113, col: 6, json: styleSubTotal({data: '=F111*1.17'})},
            {sheet: 1, row: 113, col: 7, json: styleSubTotal({data: '=G111*1.17'})},
            {sheet: 1, row: 113, col: 8, json: styleSubTotal({data: '=H111*1.17'})},
    //
            {sheet: 1, row: 112, col: 3, json: styleSubTotalMoney("$", {data: '=C111/6.35'})},
            {sheet: 1, row: 112, col: 4, json: styleSubTotalMoney("$", {data: '=D111/6.35'})},
            {sheet: 1, row: 112, col: 5, json: styleSubTotalMoney("$", {data: '=E111/6.35'})},
            {sheet: 1, row: 112, col: 6, json: styleSubTotalMoney("$", {data: '=FC111/6.35'})},
            {sheet: 1, row: 112, col: 7, json: styleSubTotalMoney("$", {data: '=G111/6.35'})},
            {sheet: 1, row: 112, col: 8, json: styleSubTotalMoney("$", {data: '=H111/6.35'})},
    //
            {sheet: 1, row: 113, col: 2, json: {data: '合計（含稅）：'}},
            {sheet: 1, row: 113, col: 3, json: styleSubTotal({data: '=C111*1.17'})},
            {sheet: 1, row: 113, col: 4, json: styleSubTotal({data: '=D111*1.17'})},
            {sheet: 1, row: 113, col: 5, json: styleSubTotal({data: '=E111*1.17'})},
            {sheet: 1, row: 113, col: 6, json: styleSubTotal({data: '=F111*1.17'})},
            {sheet: 1, row: 113, col: 7, json: styleSubTotal({data: '=G111*1.17'})},
            {sheet: 1, row: 113, col: 8, json: styleSubTotal({data: '=H111*1.17'})},
            {sheet: 1, row: 114, col: 2, json: {data: '運費：'}},
    // LAST ONE ============================================================================================
            {sheet: 1, row: 1, col: 1, json: {data: ""}}
    );
    return patch_A0601_1_cells;
}