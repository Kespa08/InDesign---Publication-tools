// Every declaration lives INSIDE the app.doScript callback.
// In some InDesign versions, app.doScript serialises the function to a string
// and re-runs it in a fresh scope — any variable declared outside the callback
// becomes undefined inside it, causing a silent failure. Keeping everything
// self-contained avoids that entirely.
(function () {
    app.doScript(function () {

        var doc = app.activeDocument;

        // Is this style name any bullet style (base OR "last" variant)?
        function isBulletStyle(name) {
            return /^Bullet list\s+\d+/i.test(name);
        }

        // Does this style name end with " last"?
        function isLastVariant(name) {
            return /\s+last$/i.test(name);
        }

        var changes = 0;
        var stories = doc.stories;

        for (var si = 0; si < stories.length; si++) {
            var paras = stories[si].paragraphs;

            for (var pi = 0; pi < paras.length; pi++) {
                var para      = paras[pi];
                var styleName = para.appliedParagraphStyle.name;

                if (!isBulletStyle(styleName)) continue;

                // Is the next paragraph also a bullet style?
                var nextIsBullet = false;
                if (pi + 1 < paras.length) {
                    nextIsBullet = isBulletStyle(
                        paras[pi + 1].appliedParagraphStyle.name
                    );
                }

                // What style SHOULD this paragraph carry?
                var targetName;
                if (!nextIsBullet) {
                    // Last bullet before a non-bullet — "last" suffix required.
                    targetName = isLastVariant(styleName)
                        ? styleName
                        : styleName + " last";
                } else {
                    // Bullet followed by another bullet — "last" suffix must not appear.
                    targetName = isLastVariant(styleName)
                        ? styleName.replace(/\s+last$/i, "")
                        : styleName;
                }

                if (targetName !== styleName) {
                    try {
                        var targetStyle = doc.paragraphStyles.itemByName(targetName);
                        if (targetStyle.isValid) {
                            // false = preserve character-level overrides within the paragraph
                            para.applyParagraphStyle(targetStyle, false);
                            changes++;
                        }
                    } catch (e) {}
                }
            }
        }

        alert("Done. " + changes + " paragraph style" + (changes === 1 ? "" : "s") + " corrected.");

    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Fix Bullet Last Styles");
})();
