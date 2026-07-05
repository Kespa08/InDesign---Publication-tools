(function () {

    var doc = app.activeDocument;
    var STYLE_NAME = "Table";
    var changes = 0;

    // Resolve the target style once, up front — this is a document-level
    // question, not a per-table one, so it only needs asking once.
    var targetStyle = null; // Construct a semantic object "target style" by declaring it as a variable and assigning null – the absence of a value. 
    var allTS = doc.allTableStyles; // Construct a semantic object that references all table styles in the active document. 
    
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
            // Container := app.activeDocument.stories, list := []. 
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
                    changes++; // indicate that a change has been made ***highly important for understanding transformations applied to a doc***
                }
            } catch (e) {}
        }

    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Fix Table Styles"); 
    // ScriptLanguage.JAVASCRIPT = Interpret the supplied script as JS
    // UndoModes.ENTIRE_SCRIPT = treat everything performed inside this function as one atomic operation
    // "Fix table styles" = what is displayed in InDesign's undo history

    alert("Done. " + changes + " table" + (changes === 1 ? "" : "s") + " corrected.");

})();
