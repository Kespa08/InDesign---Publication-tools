(function () {

    var doc = app.activeDocument;
    var STYLE_NAME = "Table";
    var changes = 0;

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

    // Resolve the paragraph style used for the line a table sits on, up
    // front, same reasoning as targetStyle above. doc.paragraphStyles only
    // sees styles outside any Paragraph Style Group — allParagraphStyles
    // recurses into groups, same as allTableStyles/allCellStyles above.
    var PARA_STYLE_NAME = "Table spacing";
    var tableSpacingStyle = null;
    var allPS = doc.allParagraphStyles;
    for (var p = 0; p < allPS.length; p++) {
        if (allPS[p].name === PARA_STYLE_NAME) { tableSpacingStyle = allPS[p]; break; }
    }

    if (!tableSpacingStyle || !tableSpacingStyle.isValid) {
        alert("Paragraph style \"" + PARA_STYLE_NAME + "\" does not exist in this document.");
        return;
    }

    var spacingChanges = 0;

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

        // # SPACING FIX: isolate every anchored table onto its own paragraph,
        // then make sure that paragraph carries "Table spacing".
        //
        // A table anchored inline in text is a single character embedded in
        // whatever paragraph it was typed into. If other text shares that
        // paragraph (before and/or after the table), split it out into its
        // own paragraph(s) first — the table's own paragraph is then styled
        // "Table spacing", while the split-off text keeps whatever style it
        // already had (e.g. "Body copy"), untouched.
        //
        // Story-level paragraphs are walked top scope only (doc.stories),
        // not into cells, matching the top-level pass above — this does not
        // reach tables nested inside another table's cell.
        for (var ssi = 0; ssi < stories.length; ssi++) {
            var story = stories[ssi];

            // Reverse order: splitting a paragraph only ever adds new
            // paragraphs AFTER it, so walking from the last paragraph down
            // to the first means every not-yet-visited (lower) index is
            // still valid when we get to it.
            for (var pi = story.paragraphs.length - 1; pi >= 0; pi--) {
                var para = story.paragraphs[pi];

                // Locate every table anchor in this paragraph, left to
                // right, before making any edits. A Character is a Text
                // range one character long, so character.tables.length > 0
                // means that character IS a table's anchor — the same
                // .tables property collectTables() already relies on above.
                var scanChars = para.characters;
                var anchorIdx = [];
                for (var ci = 0; ci < scanChars.length; ci++) {
                    try {
                        if (scanChars[ci].tables.length > 0) anchorIdx.push(ci);
                    } catch (e) {}
                }
                if (anchorIdx.length === 0) continue;

                // Process anchors right to left: each anchor's own index is
                // only ever affected by edits made to its right, and those
                // happen first in this order.
                for (var ai = anchorIdx.length - 1; ai >= 0; ai--) {
                    try {
                        var idx = anchorIdx[ai];
                        var chars = story.paragraphs[pi].characters;

                        var lastIdx = chars.length - 1;
                        var hasReturn = (chars[lastIdx].contents === "\r");
                        var lastContentIdx = hasReturn ? lastIdx - 1 : lastIdx;

                        var trailingExists = idx < lastContentIdx;
                        var leadingExists = idx > 0;

                        if (trailingExists) {
                            // Break immediately after the anchor.
                            story.paragraphs[pi].characters[idx].insertionPoints[1].contents = "\r";
                        }
                        if (leadingExists) {
                            // Re-fetch: the trailing break above (if any) only
                            // affects text after idx, so idx itself is still
                            // valid, but re-fetching keeps the specifier
                            // unambiguous rather than relying on a stale one.
                            story.paragraphs[pi].characters[idx].insertionPoints[0].contents = "\r";
                        }

                        // The table now sits alone on its own paragraph —
                        // at pi if there was no leading text to split off,
                        // otherwise at pi + 1.
                        var tablePara = story.paragraphs[pi + (leadingExists ? 1 : 0)];

                        // Split and restyle are independent checks: a table
                        // already alone on its own paragraph still gets
                        // corrected here if that paragraph carries the wrong
                        // style.
                        if (tablePara.appliedParagraphStyle !== tableSpacingStyle) {
                            tablePara.applyParagraphStyle(tableSpacingStyle, false);
                            spacingChanges++;
                        }
                    } catch (e) {}
                }
            }
        }

    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Fix Table Styles & Spacing");
    // ScriptLanguage.JAVASCRIPT = Interpret the supplied script as JS
    // UndoModes.ENTIRE_SCRIPT = treat everything performed inside this function as one atomic operation
    // "Fix Table Styles & Spacing" = what is displayed in InDesign's undo history

    alert("Done. " + changes + " table" + (changes === 1 ? "" : "s") + " corrected. "
        + spacingChanges + " table paragraph" + (spacingChanges === 1 ? "" : "s") + " reformatted.");

})();
