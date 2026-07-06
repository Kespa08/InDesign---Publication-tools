(function () {

    // Reapplies every paragraph's OWN currently-assigned style back onto
    // itself. This is the scripting equivalent of the "Apply <Style Name>"
    // menu command (NOT "Apply <Style Name>, Clear Overrides") run across
    // the whole document in one pass, for every style at once.
    //
    // Because each paragraph already knows which style it carries, there is
    // no need to loop doc.allParagraphStyles and then search for matching
    // paragraphs — that would be a style-loop nested inside a paragraph-loop
    // for no benefit. Reading each paragraph's own appliedParagraphStyle and
    // handing it straight back is the same end result with one pass instead
    // of (styles x paragraphs).

    var doc = app.activeDocument;
    var changes = 0;

    // Wrap all writes in a single undo step so Cmd+Z reverses everything at once.
    app.doScript(function () {

        // doc.stories covers all threaded text flows in the document.
        var stories = doc.stories;

        for (var si = 0; si < stories.length; si++) {
            var paras = stories[si].paragraphs;

            for (var pi = 0; pi < paras.length; pi++) {
                var para = paras[pi];

                try {
                    // false = preserve overrides (character-level and local
                    // paragraph formatting), matching "Apply", not
                    // "Apply, Clear Overrides".
                    para.applyParagraphStyle(para.appliedParagraphStyle, false);
                    changes++;
                } catch (e) {}
            }
        }

    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Apply Text Style");

    alert("Done. " + changes + " paragraph" + (changes === 1 ? "" : "s") + " reapplied.");

})();
