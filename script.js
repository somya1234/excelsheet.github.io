const $ = require("jquery");
const dialog = require("electron").remote.dialog;
const fs = require("fs");

$(document).ready(function () {
    let db = [];
    let lsc;
    let copiedObject = "";

    //click on home
    $(".menu-options").on("click", function () {
        $(".sub-menu-options").removeClass("selected");
        let idName = $(this).attr("id");
        if (idName == "New" || idName == "format" || idName == 'help') {
            idName = "home";
        }
        $(`.${idName}-options`).addClass("selected");
    })

    //get address of the cell clicked in address-container
    $(".cell-container .cell").on("click", function () {
        let rId = Number($(this).attr("row-id")) + 1;
        let cId = Number($(this).attr("col-id"));
        cId = String.fromCharCode(cId + 65);
        $(".address-container").val(rId + cId);
        lsc = this;

        let address = $(".address-container").val();
        let { rowId, colId } = getRCfromAddr(address);


        // highlight its respective row and col
        let leftColArr = $(".left-col .cell");
        $(leftColArr).removeClass("active");
        let curr_row = leftColArr[rowId];
        $(curr_row).addClass("active");
        //for cols in top row
        let topRowArr = $(".top-row").find(".cell");
        $(topRowArr).removeClass("active");
        let curr_col = topRowArr[colId];
        $(curr_col).addClass("active");


        //whenever we click a cell, it should be darker if it has the following properties previously
        //so as to do that, we check properties from database and apply an active class over it.
        if (db[rowId][colId].bold) {
            //if database has bold property, then active the cell
            $("#bold").addClass("active");
        } else {
            //if not, then remove active class.
            $("#bold").removeClass("active");
        }
        if (db[rowId][colId].italic) {
            $("#italic").addClass("active");
        } else {
            $("#italic").removeClass("active");
        }
        if (db[rowId][colId].underline) {
            $("#underline").addClass("active");
        } else {
            $("#underline").removeClass("active");
        }
        if (db[rowId][colId].left) {
            $("#left").addClass("active");
        } else {
            $("#left").removeClass("active");
        }
        if (db[rowId][colId].center) {
            $("#center").addClass("active");
        } else {
            $("#center").removeClass("active");
        }
        if (db[rowId][colId].right) {
            $("#right").addClass("active");
        } else {
            $("#right").removeClass("active");
        }
    })

    //Click on New 
    $("#New").on("click", function () {
        db = [];
        let rows = $(".cell-container").find(".row");
        for (let i = 0; i < rows.length; i++) {
            let row = [];
            let cols = $(rows[i]).find(".cell");
            for (let j = 0; j < cols.length; j++) {
                let cell = {
                    value: "",
                    formula: "",
                    downstream: [],
                    upstream: [],
                    fontFamily: "sans-serif",
                    fontSize: 14,
                    bold: false,
                    italic: false,
                    underline: false,
                    fontColor: "black",
                    bgColor: "white",
                    left: false,
                    center: false,
                    right: false
                }
                $(cols[j]).html(cell.value);
                $(cols[j]).css("font-family", cell.fontFamily);
                $(cols[j]).css("font-size", cell.fontSize + "px");
                $(cols[j]).css("font-weight", (cell.bold) ? "bolder" : "normal");
                $(cols[j]).css("font-style", (cell.italic) ? "italic" : "italic");
                $(cols[j]).css("text-decoration", (cell.underline) ? "underline" : "none");
                $(cols[j]).css("color", cell.fontColor);
                $(cols[j]).css("background-color", cell.bgColor);
                if (cell.center) { $(cols[j]).css("text-align", "center"); }
                else if (cell.right) { $(cols[j]).css("text-align", "right"); }
                else {
                    $(cols[j]).css("text-align", "left");
                }
                row.push(cell);
            }
            db.push(row);
        }
        // cell 1 is always clicked by default.
        let allCells = $(".cell-container").find(".cell");
        $(allCells[0]).click();
        //  console.log(db);
        updateHeight();
    })

    //Open a sheet
    $("#Open").click(async function () {
        let odb = await dialog.showOpenDialog();
        let fPath = odb.filePaths[0];
        let content = await fs.promises.readFile(fPath);
        db = JSON.parse(content);
        let rows = $(".cell-container").find(".row");
        for (let i = 0; i < rows.length; i++) {
            let cells = $(rows[i]).find(".cell");
            for (let j = 0; j < cells.length; j++) {
                let cellObject = db[i][j];
                $(cells[j]).html(cellObject.value);
                $(cells[j]).css("font-family", cellObject.fontFamily);
                $(cells[j]).css("font-size", cellObject.fontSize + "px");
                $(cells[j]).css("font-weight", (cellObject.bold) ? "bolder" : "normal");
                $(cells[j]).css("font-style", (cellObject.italic) ? "italic" : "italic");
                $(cells[j]).css("text-decoration", (cellObject.underline) ? "underline" : "none");
                $(cells[j]).css("color", cellObject.fontColor);
                $(cells[j]).css("background-color", cellObject.bgColor);
                if (cellObject.center) { $(cells[j]).css("text-align", "center"); }
                else if (cellObject.right) { $(cells[j]).css("text-align", "right"); }
                else {
                    $(cells[j]).css("text-align", "left");
                }
            }
            //update height of the left-col cell if the previously opened sheet has something big font size
            let height = $(cells[0]).height();
            let rowNo = i;
            let leftColArr = $(".left-col .cell");
            let cellNo = leftColArr[rowNo];
            $(cellNo).css("height", height);
        }
        console.log("file opened successfully");
    });

    //Save the document
    let fileSaver = document.querySelector("#Save");
    fileSaver.addEventListener("change", function () {
        let fPath = fileSaver.files[0].path;
        let jsonData = JSON.stringify(db);
        fs.writeFileSync(fPath, jsonData);
        console.log("written to disk");
    })

    $("#help").on("click", function () {
        window.open("https://documentation.libreoffice.org/en/english-documentation/");
        console.log("help clicked");
    })

    $(".add-sheet").on("click", function () {
        // window.init();
        console.log("new sheet opened");
    })

    /***************************************Styling *********************************************************************** */
    $("#font-family").on("change", function () {
        let address = $(".address-container").val();
        let { rowId, colId } = getRCfromAddr(address);
        let fontSelected = $(this).val();
        $(lsc).css("font-family", fontSelected);
        db[rowId][colId].fontFamily = fontSelected;
        //update height of the left-col cell
        let height = $(lsc).height();
        let rowNo = $(lsc).attr("row-id");
        let leftColArr = $(".left-col .cell");
        let cellNo = leftColArr[rowNo];
        $(cellNo).css("height", height);
    })

    $("#font-size-list").on("change", function () {
        let { rowId, colId } = getRCfromElem(lsc);
        let cellObject = db[rowId][colId];
        let fontSize = $(this).val();
        cellObject.fontSize = fontSize;
        $(lsc).css("font-size", fontSize + "px");
        //update height of the left-col cell
        let height = $(lsc).height();
        let rowNo = $(lsc).attr("row-id");
        let leftColArr = $(".left-col .cell");
        let cellNo = leftColArr[rowNo];
        $(cellNo).css("height", height);

        console.log("font size updated");
    })

    $("#bold").on("click", function () {
        //this fn doesn't depend on any of the work done before.
        $(this).toggleClass("active");
        let { rowId, colId } = getRCfromElem(lsc);
        let cellObject = db[rowId][colId];
        $(lsc).css("font-weight", cellObject.bold ? "normal" : "bolder");
        cellObject.bold = !cellObject.bold;
        //update height of the left-col cell
        let height = $(lsc).height();
        let rowNo = $(lsc).attr("row-id");
        let leftColArr = $(".left-col .cell");
        let cellNo = leftColArr[rowNo];
        $(cellNo).css("height", height);
    })

    $("#italic").on("click", function () {
        $(this).toggleClass("active");
        let { rowId, colId } = getRCfromElem(lsc);
        let cellObject = db[rowId][colId];
        $(lsc).css("font-style", cellObject.italic ? "normal" : "italic");
        cellObject.italic = !cellObject.italic;
        //update height of the left-col cell
        let height = $(lsc).height();
        let rowNo = $(lsc).attr("row-id");
        let leftColArr = $(".left-col .cell");
        let cellNo = leftColArr[rowNo];
        $(cellNo).css("height", height);
    })

    $("#underline").on("click", function () {
        $(this).toggleClass("active");
        let { rowId, colId } = getRCfromElem(lsc);
        let cellObject = db[rowId][colId];
        $(lsc).css("text-decoration", cellObject.underline ? "none" : "underline");
        cellObject.underline = !cellObject.underline;
    })

    $("#font-color").on("change", function () {
        let fontColorVal = $(this).val();
        let { rowId, colId } = getRCfromElem(lsc);
        let cellObject = db[rowId][colId];
        $(lsc).css("color", fontColorVal);
        cellObject.fontColor = fontColorVal;
    });

    $("#bg-color").on("change", function () {
        let bgColorVal = $(this).val();
        let { rowId, colId } = getRCfromElem(lsc);
        let cellObject = db[rowId][colId];
        $(lsc).css("background-color", bgColorVal);
        cellObject.bgColor = bgColorVal;
    })

    $("#left").on("click", function () {
        $(this).toggleClass("active");
        let { rowId, colId } = getRCfromElem(lsc);
        let cellObject = db[rowId][colId];
        cellObject.left = (cellObject.left) ? false : true;
        if (cellObject.left) {
            $("#center").removeClass("active");
            $("#right").removeClass("active");
            cellObject.center = false;
            cellObject.right = false;
        }
        $(lsc).css("text-align", "left");
    })

    $("#center").on("click", function () {
        $(this).toggleClass("active");
        let { rowId, colId } = getRCfromElem(lsc);
        let cellObject = db[rowId][colId];
        cellObject.center = (cellObject.center) ? false : true;
        if (cellObject.center) {
            $(lsc).css("text-align", "center");
            $("#left").removeClass("active");
            $("#right").removeClass("active");
            cellObject.left = false;
            cellObject.right = false;
        } else {
            $(lsc).css("text-align", "left");
        }
    })

    $("#right").on("click", function () {
        $(this).toggleClass("active");
        let { rowId, colId } = getRCfromElem(lsc);
        let cellObject = db[rowId][colId];
        cellObject.right = (cellObject.right) ? false : true;
        if (cellObject.right) {
            $(lsc).css("text-align", "right");
            $("#left").removeClass("active");
            $("#center").removeClass("active");
            cellObject.left = false;
            cellObject.center = false;
        } else {
            $(lsc).css("text-align", "left");
        }
    })

    $(".cut").on("click",function(){
        let {rowId,colId} = getRCfromElem(lsc);
        let cell = {
            value: "",
            formula: "",
            downstream: [],
            upstream: [],
            fontFamily: "sans-serif",
            fontSize: 14,
            bold: false,
            italic: false,
            underline: false,
            fontColor: "black",
            bgColor: "white",
            left: false,
            center: false,
            right: false
        }
        copiedObject = db[rowId][colId];
        //update formula also.
        if(db[rowId][colId].formula){
            removeFormula(fb[rowId][colId].formula);
        }
        db[rowId][colId] = cell;
        $(lsc).html(cell.value);
        $(lsc).css("font-family", cell.fontFamily);
        $(lsc).css("font-size", cell.fontSize + "px");
        $(lsc).css("font-weight", (cell.bold) ? "bolder" : "normal");
        $(lsc).css("font-style", (cell.italic) ? "italic" : "italic");
        $(lsc).css("text-decoration", (cell.underline) ? "underline" : "none");
        $(lsc).css("color", cell.fontColor);
        $(lsc).css("background-color", cell.bgColor);
        $(lsc).css("text-align", "left");
        //maintain height of rows also.
        let height = $(lsc).height();
        let rowNo = $(lsc).attr("row-id");
        let leftColArr = $(".left-col").find(".cell");
        let cellNo = leftColArr[rowNo];
        $(cellNo).css("height", height);
    })

    $(".copy").on("click",function(){
        let {rowId,colId} = getRCfromElem(lsc);
        let cellObject = db[rowId][colId];
        copiedObject = cellObject;
    })
    
    $(".paste").on("click",function(){
        // when we use paste fn, then (.cell-container .cell, is not clicked), so we need to update db.
        //not copying their formula 
        let formula = copiedObject.formula;
        
        $(lsc).html(copiedObject.value);
        $(lsc).css("font-family", copiedObject.fontFamily);
        $(lsc).css("font-size", copiedObject.fontSize + "px");
        $(lsc).css("font-weight", (copiedObject.bold) ? "bolder" : "normal");
        $(lsc).css("font-style", (copiedObject.italic) ? "italic" : "italic");
        $(lsc).css("text-decoration", (copiedObject.underline) ? "underline" : "none");
        $(lsc).css("color", copiedObject.fontColor);
        $(lsc).css("background-color", copiedObject.bgColor);
        if (copiedObject.center) { $(lsc).css("text-align", "center"); }
        else if (copiedObject.right) { $(lsc).css("text-align", "right"); }
        else {
            $(lsc).css("text-align", "left");
        }
        let {rowId,colId} = getRCfromElem(lsc);
        db[rowId][colId] = copiedObject;
        //maintain height of rows
        let height = $(lsc).height();
        let rowNo = $(lsc).attr("row-id");
        let leftColArr = $(".left-col").find(".cell");
        let cellNo = leftColArr[rowNo];
        $(cellNo).css("height", height);
    })
    /******************************************* Styling Ends ************************************************************ */

    /*********************************************Scrolling and Height *************************************************** */

    $(".cell-container .cell").on("input", function () {
        let height = $(this).height();
        let rowId = $(this).attr("row-id");
        let leftColArr = $(".left-col").find(".cell");
        let cell = leftColArr[rowId];
        $(cell).css("height", height);
    })

    $(".grid").on("scroll", function () {
        let vS = $(this).scrollTop();
        let hS = $(this).scrollLeft();
        // in style.css, increase z-index.
        $(".top-left-box,.top-row").css("top", vS + "px");
        $(".top-left-box,.left-col").css("left", hS);
    })

    /*********************************************Scrolling and Height Ends*************************************************** */

    /********************************************* Formula ******************************************************************** */

    $(".formula-container").on("blur", function () {
        let formula = $(this).val();
        let { rowId, colId } = getRCfromElem(lsc);
        let cellObject = db[rowId][colId];
        if (cellObject.formula == formula) {
            return;
        }

        if (cellObject.formula) {
            removeFormula(rowId, colId);
            $(lsc).html("");
        }

        cellObject.formula = formula;
        let ans = setUpFormula(formula, rowId, colId);
        if(!ans){
            cellObject.formula = "";
            return;
        }
        if (!isValid(rowId, colId)) {
            //if there is a cycle
            const options = {
                type: 'error',
                buttons: ['Ok', 'Help'],
                defaultId: 0,
                title: 'My Excel',
                message: 'There are one or more circular references where a formula refers to its own cell either directly or indirectly. This might cause them to calculate incorrectly.',
                detail: 'Try removing or changing the references, or moving the formulas to different cells.',
            };
            dialog.showMessageBox(null, options, (response) => {
                console.log(response);
            });
            removeFormula(rowId, colId);
            return;
        }
        ans = evaluate(formula, rowId, colId);
        $(lsc).html(ans);
        updateDownstream(rowId, colId);
    })

    function isValid(rowId, colId) {
        let ans = isCyclic(rowId, colId, rowId, colId);
        return ans;
    }

    function isCyclic(rowId, colId, row, col) {
        //check upstream of all starting from the cell where formula is put
        let upstreamArr = db[rowId][colId].upstream;
        for (let i = 0; i < upstreamArr.length; i++) {
            let obj = upstreamArr[i];
            let pRowId = obj.rowId;
            let pColId = obj.colId;
            if (pRowId == row && pColId == col) {
                return false;
            }
            let ans = isCyclic(pRowId, pColId, row, col);
            if (!ans) { return ans; }
        }
        return true;
    }

    //update the data entered in UI to db
    $(".cell-container .cell").on("blur", function () {
        let val = $(this).html();
        let rowId = $(this).attr("row-id");
        let colId = $(this).attr("col-id");
        if (val == db[rowId][colId].value) {
            return;
        }
        if (db[rowId][colId].formula) {
            removeFormula(rowId, colId);
        }
        db[rowId][colId].value = val;
        //when I enter value using tab, then the address-container do not change, so don't use getRCfromAddress method here.
        updateDownstream(rowId, colId);
    })

    function setUpFormula(formula, cRowId, cColId) {
        formula = formula.split(" ");
        for (let i = 0; i < formula.length; i++) {
            let parentAddr = formula[i];
            //check if they have space or not
            if (parentAddr.includes('+') || parentAddr.includes('-') || parentAddr.includes('*') || parentAddr.includes('/')) {
                if (parentAddr.length != 1) {
                    const options = {
                        type: 'error',
                        buttons: ['Ok', 'Help'],
                        defaultId: 0,
                        title: 'My Excel',
                        message: 'Pleaase check the formula.',
                        detail: 'Try adding space after every operand/operator.',
                    };
                    dialog.showMessageBox(null, options, (response) => {
                        console.log(response);
                    });
                    return false;
                }
            }
            if ((parentAddr.charAt(0) >= 'A' && parentAddr.charAt(0) <= 'Z') || (parentAddr.charAt(0) >= 'a' && parentAddr.charAt(0) <= 'z')) {
                addDownstream(parentAddr, cRowId, cColId);
            }
        }
        return true; //for the checking of formula.
    }

    function addDownstream(parentCell, cRowId, cColId) {
        //lowercase cells are also acceptable
        if (parentCell.charAt(0) >= 'a' && parentCell.charAt(0) <= 'z') {
            let ch = parentCell.charCodeAt(0) - 97 + 65;
            ch = String.fromCharCode(ch);
            let ch1 = parentCell.charAt(0);
            parentCell = parentCell.replace(ch1, ch);
        }
        let { rowId, colId } = getRCfromFormula(parentCell);
        // push children in parent's rowId and colId.
        db[rowId][colId].downstream.push({ cRowId, cColId });
        //push parent in children
        db[cRowId][cColId].upstream.push({ rowId, colId });
    }

    function removeFormula(rowId, colId) {
        let upstreamArr = db[rowId][colId].upstream;
        for (let i = 0; i < upstreamArr.length; i++) {
            let pObj = upstreamArr[i];
            //upstream object defined as {rowId, colId}
            let pRowId = pObj.rowId;
            let pColId = pObj.colId;
            removeDownstream(pRowId, pColId, rowId, colId);
        }
    }

    function removeDownstream(pRowId, pColId, cRowId, cColId) {
        let filteredDownstream = db[pRowId][pColId].downstream.filter(function (obj) {
            return (obj.cRowId != pRowId && obj.cColId != pColId);
        });
        db[pRowId][pColId].downstream = filteredDownstream;
        db[cRowId][cColId].upstream = [];
        db[cRowId][cColId].formula = "";
        db[cRowId][cColId].value = ""; // in case we remove formula completely
    }

    function evaluate(formula, cRowId, cColId) {
        let fComponents = formula.split(" ");
        for (let i = 0; i < fComponents.length; i++) {
            let fComp = fComponents[i];
            let ascii = fComp.charAt(0);
            if ((ascii >= 'A' && ascii <= 'Z') || (ascii >= 'a' && ascii <= 'z')) {
                if (ascii >= 'a' && ascii <= 'z') {
                    let ch = fComp.charCodeAt(0) - 97 + 65;
                    ch = String.fromCharCode(ch);
                    let ch1 = fComp.charAt(0);
                    let oldfComp = fComp;
                    fComp = fComp.replace(ch1, ch);
                    formula = formula.replace(oldfComp, fComp);
                }
                let { rowId, colId } = getRCfromFormula(fComp);
                let value = db[rowId][colId].value;
                // if cell has no value, set it to 0.
                if (value === "") {
                    value = 0;
                }
                formula = formula.replace(fComp, value);
            }
        }
        let ans = prefixEvaluation(formula);
        db[cRowId][cColId].value = ans;
        return ans;
    }

    function updateDownstream(pRowId, pColId) {
        let downstreamArr = db[pRowId][pColId].downstream;
        for (let i = 0; i < downstreamArr.length; i++) {
            let chObj = downstreamArr[i];
            let cRowId = chObj.cRowId;
            let cColId = chObj.cColId;
            let formula = db[cRowId][cColId].formula;
            ans = evaluate(formula, cRowId, cColId);
            $(`.cell-container .cell[row-id=${cRowId}][col-id=${cColId}]`).html(ans);
            db[cRowId][cColId].value = ans;
            updateDownstream(cRowId, cColId);
        }
    }

    function prefixEvaluation(formula) {
        const operands = new Array();
        const operators = new Array();
        let fComponents = formula.split(" ");
        for (let i = 0; i < fComponents.length; i++) {
            let fComp = fComponents[i];
            let ascii = fComp.charAt(0);
            if (ascii >= '0' && ascii <= '9') {
                operands.push(fComp);
            } else if (fComp == '(') {
                operators.push(fComp);
            } else if (fComp == ')') {
                while (operators[operators.length - 1] != '(') {
                    let op = operators.pop();
                    let op2 = operands.pop();
                    let op1 = operands.pop();
                    let ans = prefixEvaluate(op1, op2, op);
                    operands.push(ans);
                }
                operators.pop();
            } else if (fComp == '+' || fComp == '-' || fComp == '*' || fComp == '/') {
                while (operators.length > 0 && priority(operators[operators.length - 1]) >= priority(fComp)) {
                    let op = operators.pop();
                    let op2 = operands.pop();
                    let op1 = operands.pop();
                    let ans = prefixEvaluate(op1, op2, op);
                    operands.push(ans);
                }
                operators.push(fComp);
            }
        }
        while (operators.length > 0) {
            let op = operators.pop();
            let op2 = operands.pop();
            let op1 = operands.pop();
            let ans = prefixEvaluate(op1, op2, op);
            operands.push(ans);
        }
        return operands[operands.length - 1];
    }

    function priority(op) {
        if (op == '+') {
            return 1;
        } else if (op == '-') {
            return 1;
        } else if (op == '*' || op == '/') {
            return 2;
        } else {
            return 0;
        }
    }

    function prefixEvaluate(op1, op2, op) {
        op1 = Number(op1);
        op2 = Number(op2);
        if (op == '+') {
            return op1 + op2;
        } else if (op == '-') {
            return op1 - op2;
        } else if (op == '*') {
            return op1 * op2;
        } else if (op == '/') {
            return op1 / op2;
        } else {
            return 0;
        }
    }



    /********************************************** Formula Ends *************************************************************** */

    // return rowId and colId from formula address
    function getRCfromFormula(formulaComp) {
        let colId = formulaComp.charAt(0);
        let rowId = Number(formulaComp.substring(1)) - 1;
        colId = Number(colId.charCodeAt(colId)) - 65;
        return { rowId, colId };
    }

    //return rowId and colId from elem(this)
    function getRCfromElem(elem) {
        let rowId = $(elem).attr("row-id");
        let colId = $(elem).attr("col-id");
        return { rowId, colId };
    }

    //fn to return rowId and colId from address-container
    function getRCfromAddr(address) {

        let rowId = Number(address.charAt(0));
        if (address.length == 3) rowId = rowId * 10 + Number(address.charAt(1))
        rowId = rowId - 1;
        let colId = address.charCodeAt(address.length - 1);
        colId = colId - 65;
        return { rowId, colId };
    }

    function updateHeight(){
        //update height of the left-col cell
        let height = $(lsc).height();
        let leftColArr = $(".left-col .cell");
        for(let i=0;i<leftColArr.length;i++){
            let cellNo = leftColArr[i];
            $(cellNo).css("height", height);
        }
    }

    function init() {
        $("#home").trigger("click");
        $("#New").trigger("click");
    }
    init();
})