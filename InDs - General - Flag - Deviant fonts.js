/**
 * Apply character style "Font flag" to every run of text whose applied
 * font is not in the "Neue Frutiger World" family.
 *
 * Checked across body text AND paragraphs inside table cells.
 * Count of flagged runs is reported at the end.
 */

(function () {
    if (!app.documents.length) {
        alert("No document open.");
        return;
    }

    var doc = app.activeDocument;

    // Ensure the "Font flag" character style exists; create it if not.
    // Predicate form: ∀x [ ActiveDoc(x) → ∀y [ FlagCharacterStyle(y) → ( ¬Contain(x,y) → CreateIn(y,x) ) ] ]
    var flagStyle;
    try {
        flagStyle = doc.characterStyles.itemByName("Font Flag");
        flagStyle.name; // force resolution — throws if style is merely a default specifier
    } catch (e) {
        // Style not found: create it with the required colour (C100 M0 Y0 K0)
        flagStyle = doc.characterStyles.add({ name: "Font flag" });
        var flagColor = doc.colors.add({
            name: "Flag Cyan",
            model: ColorModel.PROCESS,
            colorValue: [100, 0, 0, 0],   // [C, M, Y, K]
            space: ColorSpace.CMYK
        });
        flagStyle.fillColor = flagColor;
    }

    var stories = doc.stories;

    var TARGET_FONT_FAMILY = "Neue Frutiger World";
    var fontCount = 0;

    // Check a single paragraph's style runs for non-conforming fonts
    function checkParagraph(para) {
        if (!para.isValid) return;

        var ranges = para.textStyleRanges;
        for (var i = 0; i < ranges.length; i++) {
            var run = ranges[i];
            try {
                if (run.appliedFont.fontFamily !== TARGET_FONT_FAMILY) {
                    run.appliedCharacterStyle = flagStyle;
                    fontCount++;
                }
            } catch (e) {}
        }
    }

    // Walk all paragraphs inside a story's tables (cells are not exposed
    // through story.paragraphs and must be traversed explicitly)
    function checkStoryTables(story) {
        var tables;
        try {
            tables = story.tables;
        } catch (e) {
            return;
        }

        for (var t = 0; t < tables.length; t++) {
            var table = tables[t];
            var rows;
            try { rows = table.rows; } catch (e) { continue; }

            for (var r = 0; r < rows.length; r++) {
                var cells;
                try { cells = rows[r].cells; } catch (e) { continue; }

                for (var c = 0; c < cells.length; c++) {
                    var cell = cells[c];
                    var cellParas;
                    try { cellParas = cell.paragraphs; } catch (e) { continue; }

                    for (var p = 0; p < cellParas.length; p++) {
                        checkParagraph(cellParas[p]);
                    }
                }
            }
        }
    }

    // Main loop over stories
    for (var s = 0; s < stories.length; s++) {
        var story = stories[s];

        // Body paragraphs (outside tables)
        var paras = story.paragraphs;
        for (var p = 0; p < paras.length; p++) {
            checkParagraph(paras[p]);
        }

        // Paragraphs inside table cells
        checkStoryTables(story);
    }

    alert("Font mismatches flagged: " + fontCount);
})();
