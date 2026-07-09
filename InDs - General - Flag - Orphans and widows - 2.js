/**
 * Apply character style "Flag" to:
 *
 * WIDOWS:
 *   Last line of a paragraph appears in a different text container
 *   (frame OR column) than the first line.
 *
 * ORPHANS:
 *   Last line of a paragraph contains exactly one word,
 *   BUT paragraphs that are only one word long are ignored.
 *
 * Both checks are applied to body text AND to paragraphs inside table cells.
 * Entire offending line is highlighted.
 * Counts of widows and orphans are reported at the end.
 */

(function () {
    if (!app.documents.length) {
        alert("No document open.");
        return;
    }

    var doc = app.activeDocument;

    // Ensure the "Flag" character style exists; create it if not.
    // Predicate form: ∀x [ ActiveDoc(x) → ∀y [ FlagCharacterStyle(y) → ( ¬Contain(x,y) → CreateIn(y,x) ) ] ]
    var flagStyle;
    try {
        flagStyle = doc.characterStyles.itemByName("Flag");
        flagStyle.name; // force resolution — throws if style is merely a default specifier
    } catch (e) {
        // Style not found: create it with the required colour (C0 M100 Y0 K0)
        flagStyle = doc.characterStyles.add({ name: "Flag" });
        var flagColor = doc.colors.add({
            name: "Flag Magenta",
            model: ColorModel.PROCESS,
            colorValue: [0, 100, 0, 0],   // [C, M, Y, K]
            space: ColorSpace.CMYK
        });
        flagStyle.fillColor = flagColor;
    }

    var stories = doc.stories;

    var widowCount = 0;
    var orphanCount = 0;

    // Safely get the parent text frame of a line
    function getFrame(line) {
        try {
            if (!line || !line.isValid) return null;
            var frames = line.parentTextFrames;
            if (!frames || frames.length === 0) return null;
            return frames[0];
        } catch (e) {
            return null;
        }
    }

    // Get the top Y coordinate of a line
    function getLineTop(line) {
        try {
            if (!line || !line.isValid) return null;
            return line.baseline - line.pointSize;
        } catch (e) {
            return null;
        }
    }

    // Count words, treating hyphens as separators
    function countWords(str) {
        var cleaned = str.replace(/^\s+|\s+$/g, "");
        if (cleaned === "") return 0;
        return cleaned.split(/[\s\-–—]+/).length;
    }

    // Check a single paragraph for widows and orphans and apply flag style
    function checkParagraph(para, inCell) {
        if (!para.isValid) return;

        var lines = para.lines;
        if (lines.length === 0) return;

        var firstLine = lines[0];
        var lastLine  = lines[-1];

        if (!firstLine.isValid || !lastLine.isValid) return;

        // --- WIDOW DETECTION ---
        // Inside a cell: compare frame identity, which is reliable across
        // page-split tables where Y-coordinates reset between pages.
        // Outside a cell: use Y-coordinate inversion (original behaviour).
        var isWidow = false;

        if (inCell) {
            var firstFrame = getFrame(firstLine);
            var lastFrame  = getFrame(lastLine);
            if (firstFrame && lastFrame) {
                isWidow = (firstFrame !== lastFrame);
            }
        } else {
            var firstFrame = getFrame(firstLine);
            var lastFrame  = getFrame(lastLine);
            if (!firstFrame || !lastFrame) return;

            var firstTop = getLineTop(firstLine);
            var lastTop  = getLineTop(lastLine);
            if (firstTop !== null && lastTop !== null) {
                isWidow = (lastTop < firstTop);
            }
        }

        if (isWidow) {
            try {
                lastLine.appliedCharacterStyle = flagStyle;
                widowCount++;
            } catch (e) {}
        }

        // --- ORPHAN DETECTION (single-word last line) ---
        var lastText      = lastLine.contents;
        var lastWordCount = countWords(lastText);
        var paraWordCount = countWords(para.contents);
        var isOrphan      = (lastWordCount === 1 && paraWordCount > 1);

        if (isOrphan) {
            try {
                lastLine.appliedCharacterStyle = flagStyle;
                orphanCount++;
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
                        checkParagraph(cellParas[p], true);
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
            checkParagraph(paras[p], false);
        }

        // Paragraphs inside table cells
        checkStoryTables(story);
    }

    alert(
        "Widows flagged: " + widowCount +
        "\nOrphans flagged: " + orphanCount
    );
})();
