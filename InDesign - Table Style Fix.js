(function () {

    var doc = app.activeDocument;
    var STYLE_NAME = "Table";
    var HEADER_STYLE_NAME = "Header";
    var BODY_STYLE_NAME = "Body";
    var styleChanges = 0;
    var headerChanges = 0;
    var rowChanges = 0;
    var colWidthChanges = 0;

    function roundTo(n, dec) {
        var f = Math.pow(10, dec);
        return Math.round(n * f) / f;
    }

    // The width of whatever directly contains the table: a Cell's width
    // if the table is nested inside another table, or a TextFrame's raw
    // geometric width if the table sits directly in the main text flow.
    // Adapted from getTablePage() in "InDesign - Selection - Match and
    // merge - 19.js", which walks the same parent chain but always
    // continues past a Cell to reach a TextFrame; here we stop at
    // whichever of the two is found first, since that's the table's
    // immediate container.
    function getContainerWidth(tbl) {
        try {
            var cnt = tbl.parent;
            var cntname = cnt.constructor.name
            while (cnt && cnt.constructor &&
                   cntname !== "Cell" &&
                   cntname !== "TextFrame") {
                cnt = cnt.parent;
            }
            if (!cnt) return null;
            if (cntname === "Cell") return cnt.width;
            if (cntname === "TextFrame") {
                var b = cnt.geometricBounds;
                return b[3] - b[1]; // b[0] = top, b[1] = left, b[2] = bottom, b[3] = right. Therefore b[3] - b[1] = Right bound - Left bound
            }
        } catch (e) {}
        return null;
    }

    // Resolve the target style once, up front — this is a document-level
    // question, not a per-table one, so it only needs asking once.
    var targetStyle = null; // Construct a semantic object "target style" by declaring it as a variable and assigning null – the absence of a value. 
    var allTS = doc.allTableStyles; // Construct a semantic object that references all table styles in the active document. 

    // # SEARCH FUNCTION: DOC - STYLES - TABLES + CELLS 
    // Construct a for loop that iterates over all elements within the array allTS {TS_1, TS_2... TS_n}
    // If one such element in the array TS_[i] corresponds to the value assigned to STYLE_NAME, such that TS_[i].name === "Table" =>
    // assign tagetStyle to the relevant tableStyle object && Break the loop, such that targetStyle := TS_[i].name.
    for (var i = 0; i < allTS.length; i++) {
        if (allTS[i].name === STYLE_NAME) { targetStyle = allTS[i]; break; }
    }

    // if targetStyle = null => !targetStyle = True, such that the user is provide an alert.
    // if targetStyle ≠ null, yet the object assigned to it from doc.allTableStyles is invalid => provide alert (this is a fail safe in case targetStyle is assigned something but that thing is not valid syntax...).
    if (!targetStyle || !targetStyle.isValid) {
        alert("Table style \"" + STYLE_NAME + "\" does not exist in this document.");
        return;
    }
    
    // # SEARCH FUNCTION: DOC.STYLES – CELLS
    var headerStyle = null, bodyStyle = null;
    var allCS = doc.allCellStyles;
    for (var j = 0; j < allCS.length; j++) {
        if (allCS[j].name === HEADER_STYLE_NAME) headerStyle = allCS[j]; // Find the cell style X element in the doc (X) whose name corresponds to "Header". If this is found => headerStyle := X
        if (allCS[j].name === BODY_STYLE_NAME) bodyStyle = allCS[j]; // Find the cell style element in the doc (Y) whose name correponds to "Body". If this is found => bodyStyle := Y
    }

    if (!headerStyle || !headerStyle.isValid) {
        alert("Cell style \"" + HEADER_STYLE_NAME + "\" does not exist in this document.");
        return;
    }
    if (!bodyStyle || !bodyStyle.isValid) {
        alert("Cell style \"" + BODY_STYLE_NAME + "\" does not exist in this document.");
        return;
    }

    // Wrap all writes in a single undo step so Cmd+Z reverses everything at once.
    app.doScript(function () {

        // Collect every table in the document, including tables nested
        // inside a cell of another table, at any depth. Works for both a
        // Story and a Cell, since both expose a .tables collection.
        function collectTables(container, list) { // Container = an object with elements, List = an empty array to assign relevant container objects to (in this case, Tables ∈ Stories)
            var tables = container.tables;
            for (var t = 0; t < tables.length; t++) {
                list.push(tables[t]); //push any instance of a table at index t in the given container under iteration.
                var cells = tables[t].cells; // assign a variable to the cells property of all table objects
                for (var c = 0; c < cells.length; c++) {
                    try { collectTables(cells[c], list); } catch (e) {} // apply the collectTables function for all cells in a table and push said tables into the array (Table ∈ Cells ∈ Table) – this is recursive and will check the cells of each nested table
                }
            }
        }

        var allTables = [];
        var stories = doc.stories;
        for (var si = 0; si < stories.length; si++) {
            collectTables(stories[si], allTables);
            // assign the following inputs to the collectTables function – Container := app.activeDocument.stories, list := [].
            // collectTables nested in a for loop produces an array that extracts every instance of a table in the doc
            // this would produce something like [Table_1, Table_2... Table_n]
        }

        // Only write when something actually needs to change; direct
        // property assignment preserves local cell/row overrides.
        for (var k = 0; k < allTables.length; k++) {
            var tbl = allTables[k]; // transform the present table into a variable to be used
            try {
                if (tbl.appliedTableStyle !== targetStyle) { // check if the table style applied to the object corresponds to the targetStyle
                    tbl.appliedTableStyle = targetStyle;  // if the table does not have the appropriate style, change it to the target one.
                    styleChanges++; // indicate that a change has been made ***highly important for understanding transformations applied to a doc***
                }
            } catch (e) {}

            // First row becomes the header row; every other row is body.
            // rowType converts the existing row in place — unlike
            // headerRowCount, which inserts a brand new row instead of
            // reclassifying row 0. A row already marked HEADER_ROW is
            // treated as conformant and left alone.
            try {
                var rows = tbl.rows;
                if (rows.length > 0 && rows[0].rowType !== RowTypes.HEADER_ROW) {
                    try { rows[0].rowType = RowTypes.HEADER_ROW; headerChanges++; } catch (e) {}
                }

                for (var r = 0; r < rows.length; r++) {
                    var isHeaderRow  = (r === 0);
                    var targetCS     = isHeaderRow ? headerStyle : bodyStyle; // determines what cell style row_r should have (if row_0 => targetCS = headerStyle, ELSE targetCS = bodyStyle) 
                    
                    var rowCells = rows[r].cells;
                    for (var rc = 0; rc < rowCells.length; rc++) {
                        try {
                            if (rowCells[rc].appliedCellStyle !== targetCS) {  // does the present cell style === targetCS (headerStyle OR bodyStyle)?
                                rowCells[rc].appliedCellStyle = targetCS;
                                rowChanges++;
                            }
                        } catch (e) {}
                    }
                }
            } catch (e) {}

            // All columns become equal width, sized to fill the table's
            // immediate container. Rounded to 3 decimal places before
            // comparing, so floating-point noise doesn't register as a
            // change on every re-run.
            try {
                var containerWidth = getContainerWidth(tbl);
                if (containerWidth !== null) {
                    var cols = tbl.columns;
                    var targetColWidth = roundTo(containerWidth / cols.length, 3);
                    for (var ci = 0; ci < cols.length; ci++) {
                        if (roundTo(cols[ci].width, 3) !== targetColWidth) {
                            cols[ci].width = targetColWidth;
                            colWidthChanges++;
                        }
                    }
                }
            } catch (e) {}
        }

    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Fix Table Styles");
    // ScriptLanguage.JAVASCRIPT = Interpret the supplied script as JS
    // UndoModes.ENTIRE_SCRIPT = treat everything performed inside this function as one atomic operation
    // "Fix table styles" = what is displayed in InDesign's undo history

    alert(
        "Done.\n\n" +
        styleChanges + " table style" + (styleChanges === 1 ? "" : "s") + "\n" +
        headerChanges + " row" + (headerChanges === 1 ? "" : "s") + " converted to header\n" +
        rowChanges + " cell style" + (rowChanges === 1 ? "" : "s") + "\n" +
        colWidthChanges + " column width" + (colWidthChanges === 1 ? "" : "s") + " corrected."
    );
})();
