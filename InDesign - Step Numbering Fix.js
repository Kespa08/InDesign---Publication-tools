// Every declaration lives INSIDE the app.doScript callback.
// In some InDesign versions, app.doScript serialises the function to a string
// and re-runs it in a fresh scope — any variable declared outside the callback
// becomes undefined inside it, causing a silent failure. Keeping everything
// self-contained avoids that entirely.
(function () {
    app.doScript(function () {

        var doc = app.activeDocument;

        // Returns true if the paragraph's applied style belongs to the
        // "Step style" paragraph style group folder (direct parent only).
        // Reads style.parent — no itemByName lookup needed, so the
        // doc.allParagraphStyles / grouped-style pitfall does not apply here.
        function isInStepFolder(para) {
            try {
                var style = para.appliedParagraphStyle;
                if (!style || !style.isValid) return false;
                return style.parent && style.parent.name === "Step style";
            } catch (e) { return false; }
        }

        var changes = 0;
        var stories = doc.stories;

        for (var si = 0; si < stories.length; si++) {
            var paras = stories[si].paragraphs;

            for (var pi = 0; pi < paras.length; pi++) {
                var para = paras[pi];

                if (!isInStepFolder(para)) continue;

                // Is the immediately preceding paragraph also a step?
                var prevIsStep = (pi > 0) && isInStepFolder(paras[pi - 1]);

                try {
                    if (!prevIsStep) {
                        // First step after a non-step paragraph — restart at 1.
                        if (para.numberingContinue !== false || para.numberingStartAt !== 1) {
                            para.numberingContinue = false;
                            para.numberingStartAt  = 1;
                            changes++;
                        }
                    } else {
                        // Mid-sequence step — must continue the running count.
                        if (para.numberingContinue !== true) {
                            para.numberingContinue = true;
                            changes++;
                        }
                    }
                } catch (e) {}
            }
        }

        alert("Done. " + changes + " paragraph" + (changes === 1 ? "" : "s") + " corrected.");

    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Fix Step Numbering");
})();
