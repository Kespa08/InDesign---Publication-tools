(function () {

    var doc = app.activeDocument;

    // Is this style name any bullet style (base OR "last" variant)?
    // ^Bullet list\s+\d+ matches the start of names like:
    //   "Bullet list 1", "Bullet list 2", "Bullet list 1 last", "Bullet list 2 last"
    function isBulletStyle(name) {
        return /^Bullet list\s+\d+/i.test(name);
    }

    // Does this style name end with " last" (in any capitalisation)?
    function isLastVariant(name) {
        return /\s+last$/i.test(name);
    }

    var changes = 0;

    // Wrap all writes in a single undo step so Cmd+Z reverses everything at once.
    app.doScript(function () {

        // doc.stories covers all threaded text flows in the document.
        // Iterating stories → paragraphs gives paragraphs in their correct
        // reading order across all linked text frames.
        var stories = doc.stories;

        for (var si = 0; si < stories.length; si++) {
            var paras = stories[si].paragraphs;

            for (var pi = 0; pi < paras.length; pi++) {
                var para      = paras[pi];
                var styleName = para.appliedParagraphStyle.name;

                // Skip anything that isn't a bullet style — continue jumps
                // straight to the next iteration without executing the rest.
                if (!isBulletStyle(styleName)) continue;

                // Check whether the next paragraph is also a bullet style.
                // The pi + 1 < paras.length guard prevents reading past the
                // end of the array; a missing next paragraph is treated the
                // same as a non-bullet next paragraph.
                var nextIsBullet = false;
                if (pi + 1 < paras.length) {
                    nextIsBullet = isBulletStyle(
                        paras[pi + 1].appliedParagraphStyle.name
                    );
                }

                // Derive what style this paragraph SHOULD carry.
                var targetName;
                if (!nextIsBullet) {
                    // Last bullet before a non-bullet (or end of story):
                    // the "last" suffix must be present.
                    targetName = isLastVariant(styleName)
                        ? styleName
                        : styleName + " last";
                } else {
                    // Bullet followed by another bullet:
                    // the "last" suffix must NOT be present.
                    targetName = isLastVariant(styleName)
                        ? styleName.replace(/\s+last$/i, "")
                        : styleName;
                }

                // Only write when something actually needs to change.
                if (targetName !== styleName) {
                    try {
                        // itemByName never returns null in InDesign — it returns
                        // a "dead" object instead. isValid is the correct check.
                        var targetStyle = doc.paragraphStyles.itemByName(targetName);
                        if (targetStyle.isValid) {
                            // false = preserve character-level overrides
                            // (bold, italic, etc.) applied within the paragraph.
                            para.applyParagraphStyle(targetStyle, false);
                            changes++;
                        }
                    } catch (e) {}
                }
            }
        }

    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Fix Bullet Last Styles");

    alert("Done. " + changes + " paragraph style" + (changes === 1 ? "" : "s") + " corrected.");

})();
