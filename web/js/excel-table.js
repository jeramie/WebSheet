/* Author: Jeramie Hallyburton
 * Date: 18/02/2011
 * Purpose: To replicate functionality of excel in the webbrowser
 */
/********************************** Helper Functions **************************************/
/*****************************************************************************************/
        //strips numeric chars from a string, returning only alphabetic chars
        function getAlpha(str) {
            return str.match(/[^\d]/g).join('');
        }

        //strips alphabetic chars from a string, returning only numeric chars
        function getNums(str) {
            return str.match(/\d/g).join('');
        }
		
		//converts a decimal number into its alphabetic base 26 equivalent
        function convertToBase26(i) {
            var BASE = 26;
            var result = "";
            var remainder;

            while (i > 0) {
                remainder = i % BASE === 0 ? 26 : i % BASE;
                i = Math.floor(i / BASE);
                result = chars[remainder] + result;
            }
            return result;
        }
		
		//converts a alphabetic number in base 26 into decimal
        function convertFromBase26(val) {
            var BASE = 26.0;
            var result = 0;

            for (var x = 0; x < val.length; x++) {
                var char = (val.charAt(val.length - 1 - x).charCodeAt() - 'A'.charCodeAt()) + 1;
                var power = Math.pow(BASE, x);
                result += Math.floor(power * char);
            }
            return result;
        }
		
		//returns the value of a function
        function evaluateFunction(cell) {
            var id = "#" + cell.attr("id");
            var formula = cell.attr("formula");
            var result = 0.0;
            var func = formula.substring(1, formula.indexOf('(', 1));
            if (func === "SUM") {
                var lhs = formula.substring(formula.indexOf('(', 1) + 1, formula.indexOf(':', 1));
                var rhs = formula.substring(formula.indexOf(':', 1) + 1, formula.indexOf(')', 1));

                if( !lhs.match("[A-Za-z]+[0-9]+") || !rhs.match("[A-Za-z]+[0-9]+") ) {
                    alert( "Invalid function: " + formula );
                    return "#INVALID";
                }

                var startingColumn = getAlpha(lhs);
                var endingColumn = getAlpha(rhs);

                var startingRow = getNums(lhs);
                var endingRow = getNums(rhs);

    

                //same row sum
                if (startingRow === endingRow) {
                    for (var i = convertFromBase26(startingColumn); i <= convertFromBase26(endingColumn); i++) {
                        var cellID = '#' + convertToBase26(i) + startingRow;
                        if (cellID === id) {
                            alert("Circular reference detected");
                            return "#CIRCULAR";
                        }
                        var col = $( cellID );
                        var formula = col.attr("formula");
                        //if cell contains formula, recursively call evaluateFunction, else assign col.text()
                        var value = formula !== undefined && formula.indexOf('=', 0) === 0 ? evaluateFunction(col) : col.text();
                        result += 1 * value;
                        if (isNaN(result)) {
                            alert("Invalid function or literal '" + col.text() + " at cell: " + cellID + ". Must be a valid number or formula!");
                            return "#INVALID";
                        }
                    }
                }

                //same column sum
                if ( startingColumn.toUpperCase() === endingColumn.toUpperCase() ) {
                    for (var i = 0; i <= endingRow - startingRow; i++) {
                        var cellID = '#' + startingColumn + (i * 1 + startingRow * 1);
                        if (cellID === id) {
                            alert("Circular reference detected");
                            return "#CIRCULAR";
                        }
                        var col = $(cellID);
                        var formula = col.attr("formula");
                        //if cell contains formula, recursively call evaluateFunction, else assign col.text()
                        var value = formula !== undefined && formula.indexOf('=', 0) === 0 ? evaluateFunction(col) : col.text();
                        result += 1 * value;
                        if (isNaN(result)) {
                            alert("Invalid function or literal '" + col.text() + " at cell: " + cellID + ". Must be a valid number or formula!");
                            return "#INVALID";
                        } 
                    }
                }

                //sum of set
                if (startingColumn.toUpperCase() !== endingColumn.toUpperCase() && startingRow !== endingRow) {
                    var colStart = convertFromBase26(startingColumn);
                    var colEnd = convertFromBase26(endingColumn);

                    if (colEnd * 1 < colStart * 1) {
                        var temp = colStart;
                        colStart = colEnd;
                        colEnd = temp;
                    }

                    if (endingRow * 1 < startingRow * 1) {
                        temp = startingRow;
                        startingRow = endingRow;
                        endingRow = temp;
                    }

                    for (var i = 0; i <= endingRow - startingRow; i++) {
                        for (var j = colStart; j <= colEnd; j++) {
                            var cellID = '#' + convertToBase26(j) + (i * 1 + startingRow * 1);
                            if (cellID === id) {
                                alert("Circular reference detected");
                                return "#CIRCULAR";
                            }
                            var col = $(cellID);
                            var formula = col.attr("formula");
                            //if cell contains formula, recursively call evaluateFunction, else assign col.text()
                            var value = formula !== undefined && formula.indexOf('=', 0) === 0 ? evaluateFunction(col) : col.text();
                            result += 1 * value;
                            if (isNaN(result)) {
                                alert("Invalid function or literal '" + col.text() + " at cell: " + cellID + ". Must be a valid number or formula!");
                                return "#INVALID";
                            } 
                        }
                    }
                }
            }
            return result;
        }

        //checks if string is only alpahbetic
        function isAlpha(str) {
            return /^[a-zA-Z]+$/.test(str);
        }
		//updates the value returned by all functions in the document
        function updateFunctions() {
            $('td.cell[formula *= "="]').each(function (e) {
                var form = evaluateFunction($(this));
                if (form === "#INVALID" || form === "#CIRCULAR")
                    $(this).attr("formula", "");
                $(this).text( form );
            });
        }

        /*****************************************************************************************/
        /******************************** End: Helper Functions *********************************/


        /******************************** Object Literals *****************************************/
        /*****************************************************************************************/

        //object holding all properties of the spreadsheet
        var table = {
            ref             : null, //reference to the table element
            thead           : null, //reference to the thead element
            tbody           : null, //reference to the tbody element
            numRows         : null, //current number of rows in the table
            numColumns      : null, //current number of cols in the table

            insertColumns   : function (numCols) { //function to insert new columns to the table
                table.ref.find('tr').each(function (rowCount) {
                    var row = $(this);
                    for (var i = 1; i <= numCols; i++) {
                        var colAlpha = convertToBase26(table.numColumns + (i + 1));
                        if (rowCount === 0) //create header cells
                            $.tmpl(templates.th, { 'classID': 'header', 'headerID': colAlpha, 'header': colAlpha }).appendTo(row);
                        else //create regular cells
                            $.tmpl(templates.td, { 'classID': 'cell', 'cellID': colAlpha + '' + rowCount, 'cell': '' }).appendTo(row);
                    }
                });
                table.numColumns += numCols;
                //adjust the width of the table to account for new columns
                table.ref.width(table.ref.width() + (numCols * table.cells.cellWidth));
                //wire event handler for new cells
                $('td.cell').click(tdClick);
            },

            insertRows      : function (numRows) { //function to insert new rows to the table
                //append rows to table
                for (var i = table.numRows; i < table.numRows + numRows; i++) {
                    var row = $.tmpl(templates.tr, { "trID": "row" + i });
                    table.ref.width(table.cells.cellWidth * table.numColumns + 30);
                    for (var j = 0; j <= table.numColumns; j++) {
                        var colAlpha = j > 26 ? convertToBase26(j) : chars[j];
                        if (i === 0) //create header cells
                            $.tmpl(templates.th, { 'classID': j === 0 ? 'sideColumn selectButton' : 'header', 'headerID': colAlpha, 'header': colAlpha }).appendTo(row);
                        else //create regular cells
                            $.tmpl(templates.td, { 'classID': j === 0 ? 'sideColumn' : 'cell', 'cellID': colAlpha + '' + i, 'cell': j === 0 ? i : '' }).appendTo(row);
                    }
                    if (i === 0) 
                        table.thead.append(row);
                    else 
                        table.tbody.append(row);
                }
                table.numRows += numRows;
                //wire event handler for new cells
                $('td.cell').click(tdClick);
            },

            cells : { //object to hold properties of table cells
                cellWidth       : 55,       //default width of a cell
                cellHeight      : 10,       //default height of a cell
                initialClick    : false,    //indicates whether a cell was just clicked
                previousCell    : null,     //reference to previously selected table cell
                currentCell     : null,     //reference to currently selected table cell
                previousCellID  : null,     //identifier of previously selected table cell, eg. C4
                currentCellID   : null,     //identifier of currently selected table cell, eg. D5
                previousColumnID: null,     //column identifier of previously selected cell, eg, C
                currentColumnID : null,     //column identifier of currently selected cell, eg, D
                previousRowID   : null,     //row identifier of previoulsy selected cell, eg, 4
                currentRowID    : null     //row identifier of currently selected cell, eg, 5  
            }
        };
        
        //create object literal of cached/compiled templates
        var templates = {
            th: $("#thTemplate").template(),
            td: $("#tdTemplate").template(),
            tr: $("#trTemplate").template()
        };
        
        //properties of main textbox
        var functionBox = {
            ref: null,
            hasFocus: false
        };
        
        //alphabetic character array mapping each letter to its corresponding int value
        var chars = { 1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I',
                     10: 'J', 11: 'K', 12: 'L', 13: 'M', 14: 'N', 15: 'O', 16: 'P', 17: 'Q', 18: 'R',
                     19: 'S', 20: 'T', 21: 'U', 22: 'V', 23: 'W', 24: 'X', 25: 'Y', 26: 'Z'
                 };

        /*****************************************************************************************/
        /**************************** End: Object Literals **************************************/


        /*********************************** Event Handlers **************************************/
       /*****************************************************************************************/

        var tdClick = function () {
            table.cells.previousCell = table.cells.currentCell;
            table.cells.currentCell = $(this);
            table.cells.initialClick = true;
            updateFunctions();

            //set text box to current cell value or formula
            if (table.cells.currentCell.attr("formula") === undefined)
                functionBox.ref.val(table.cells.currentCell.text());
            else
                functionBox.ref.val(table.cells.currentCell.attr("formula"));

            //unhighlight previous cell
            table.cells.previousCell.removeClass('cellSelected');

            //highight current cell
            table.cells.currentCell.addClass('cellSelected');

            table.cells.currentCellID = table.cells.currentCell.attr("id");
            table.cells.previousCellID = table.cells.previousCell.attr("id");

            table.cells.currentColumnID = getAlpha(table.cells.currentCellID);
            table.cells.previousColumnID = getAlpha(table.cells.previousCellID);

            table.cells.currentRowID = getNums(table.cells.currentCellID);
            table.cells.previousRowID = getNums(table.cells.previousCellID);

            //unhighlight old header and side bar
            $("#" + table.cells.previousColumnID).removeClass('headerSelected');
            $("#row" + table.cells.previousRowID + " > td.sideColumn").removeClass('sideColumnSelected');

            //highlight new header and side bar
            $("#" + table.cells.currentColumnID).addClass('headerSelected');
            $("#row" + table.cells.currentRowID + " > td.sideColumn").addClass('sideColumnSelected');

            //add currentCellID to top-left select box
            $("#selected-cell").text(table.cells.currentCellID);
        };

        $(document).ready(function () {
            //now that the document is rendered, set references to elements
            functionBox.ref = $(".functionBox");
            table.ref = $(".excelTable");
            table.thead = $("<thead></thead>");
            table.tbody = $("<tbody></tbody>");
            table.ref.append(table.thead);
            table.ref.append(table.tbody);

            var width = $(window).width();
            var height = $(window).height();

            //set the default number of rows and cells
            table.numRows = 0;
            table.numColumns = 32; //set table width at 4 cells past viewport

            table.insertRows(75);

            //capture global keystrokes. enter text in cell/functionbox only 
            //if a cell is selected and functionbox does not have focus
            $(document).keypress(function (e) {
                if (table.cells.currentCell != null) {
                    var char = String.fromCharCode(e.keyCode);

                    //if cell has a value, and it has been re-clicked on, clear the cell/textbox
                    if (table.cells.currentCell.text() !== "" && table.cells.initialClick && !functionBox.hasFocus) {
                        functionBox.ref.val("");
                        table.cells.currentCell.text("");
                    }
                    if (!functionBox.hasFocus)
                        functionBox.ref.val(functionBox.ref.val() + char);
                    table.cells.initialClick = false;
                    table.cells.currentCell.text(table.cells.currentCell.text() + char);

                    if (e.keyCode === 13) { //enter key press
                        if (functionBox.ref.val().charAt(0) === '=') {
                            //add formula to the cells formula attribute
                            table.cells.currentCell.attr("formula", functionBox.ref.val());
                        }
                        //update all functions
                        updateFunctions();
                    }
                }
            });

            //update cell for keystroke commands: enter, delete, backspace
            functionBox.ref.keyup(function (e) {
                if (e.keyCode !== 13)
                    table.cells.currentCell.text(functionBox.ref.val());
            });

            //scroll event for dynamically adding new columns/rows
            $(window).scroll(function (e) {
                if ($(document).width() - $(window).width() === $(window).scrollLeft()) { //if h-scrollbar has reached end
                    table.insertColumns(3);
                    //re-position the horizontal scroll bar
                    $(window).scrollLeft($(document).width() - $(window).width() - 5);
                } else if ($(document).height() - $(window).height() === $(window).scrollTop()) { //if v-scrollbar has reached bottom
                    table.insertRows(3);
                    //re-position the vertical scroll bar
                    $(window).scrollTop($(document).height() - $(window).height() - 5);
                }
            });

            //clear button event
            $(".clear").click(function () {
                $("td.cell").each(function () {
                    $(this).text("");
                    $(this).attr("formula", "");
                });
            });
        });

        /*****************************************************************************************/
        /***************************** End: Event Handlers **************************************/