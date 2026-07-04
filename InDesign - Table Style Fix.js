(function () {

    var doc = app.activeDocument;
    var STYLE_NAME = "Table";
    var changes = 0;

    // Resolve the target style once, up front — this is a document-level
    // question, not a per-table one, so it only needs asking once.
    var targetStyle = null;
    var allTS = doc.allTableStyles;
    for (var i = 0; i < allTS.length; i++) {
        if (allTS[i].name === STYLE_NAME) { targetStyle = allTS[i]; break; }
    }

    if (!targetStyle || !targetStyle.isValid) {
        alert("Table style \"" + STYLE_NAME + "\" does not exist in this document.");
        return;
    }

    // Wrap all writes in a single undo step so Cmd+Z reverses everything at once.
    app.doScript(function () {

        // Collect every table in the document, including tables nested
        // inside a cell of another table, at any depth. Works for both a
        // Story and a Cell, since both expose a .tables collection.
        function collectTables(container, list) {
            var tables = container.tables;
            for (var t = 0; t < tables.length; t++) {
                list.push(tables[t]);
                var cells = tables[t].cells;
                for (var c = 0; c < cells.length; c++) {
                    try { collectTables(cells[c], list); } catch (e) {}
                }
            }
        }

        var allTables = [];
        var stories = doc.stories;
        for (var si = 0; si < stories.length; si++) {
            collectTables(stories[si], allTables);
        }

        // Only write when something actually needs to change; direct
        // property assignment preserves local cell/row overrides.
        for (var k = 0; k < allTables.length; k++) {
            var tbl = allTables[k];
            try {
                if (tbl.appliedTableStyle.name !== STYLE_NAME) {
                    tbl.appliedTableStyle = targetStyle;
                    changes++;
                }
            } catch (e) {}
        }

    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Fix Table Styles");

    alert("Done. " + changes + " table" + (changes === 1 ? "" : "s") + " corrected.");

})();
