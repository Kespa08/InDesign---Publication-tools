(function () {
    var doc = app.activeDocument;

    var REPLACEMENTS_PARAGRAPH = [
        { from: "Normal",              to: "Body copy" },
        { from: "Heading 1",           to: "Heading 1" },        // root -> nested, same name
        { from: "IP - Bullet level 1", to: "Bullet list 1" }
        // add more { from: "...", to: "..." } pairs here over time
    ];

    var REPLACEMENTS_CHARACTER = [
        { from: "Italics", to: "Book italic" }
        // add more { from: "...", to: "..." } pairs here over time
    ];

    var paragraphRemoved = 0;
    var characterRemoved = 0;
    var missingTargetsParagraph = [];
    var ambiguousTargetsParagraph = [];
    var missingTargetsCharacter = [];
    var ambiguousTargetsCharacter = [];

    app.doScript(function () {

        // Deletes each pair's root-level "from" style, reassigning its usage
        // to the nested "to" style. Works identically for paragraph and
        // character styles -- only which collections get passed in differs.
        // Root/all collections are re-fetched every iteration: an earlier
        // pair's removal can invalidate a previously-cached style reference.
        function replaceRootStyles(pairs, getRootStyles, getAllStyles, missingTargets, ambiguousTargets) {
            var removedCount = 0;
            for (var p = 0; p < pairs.length; p++) {
                var pair = pairs[p];

                var rootStyles = getRootStyles();
                var rootSet = [];
                for (var rs = 0; rs < rootStyles.length; rs++) rootSet.push(rootStyles[rs]);

                var allStyles = getAllStyles();

                // Resolve the target. Root candidates are only excluded when
                // the source and target share a name -- that's the only case
                // where a root-level match could actually be the source
                // itself (e.g. root "Heading 1" vs. nested "Heading 1").
                // When the names differ, a root-level match is unambiguous
                // and valid, so it isn't excluded.
                var sameName = (pair.from === pair.to);
                var matches = [];
                for (var a = 0; a < allStyles.length; a++) {
                    if (allStyles[a].name !== pair.to) continue;
                    if (sameName) {
                        var isRoot = false;
                        for (var i = 0; i < rootSet.length; i++) {
                            if (rootSet[i] === allStyles[a]) { isRoot = true; break; }
                        }
                        if (isRoot) continue;
                    }
                    matches.push(allStyles[a]);
                }

                if (matches.length === 0) {
                    missingTargets.push(pair.from + " -> " + pair.to);
                    continue;
                }
                if (matches.length > 1) {
                    ambiguousTargets.push(pair.from + " -> " + pair.to + " (" + matches.length + "x)");
                }
                var target = matches[0];

                // Resolve the source, root-only, from the freshly-fetched set.
                var source = null;
                for (var rs2 = 0; rs2 < rootSet.length; rs2++) {
                    if (rootSet[rs2].name === pair.from) { source = rootSet[rs2]; break; }
                }
                if (!source) continue; // nothing at root with this name -- not an error

                try {
                    source.remove(target);
                    removedCount++;
                } catch (e) {}
            }
            return removedCount;
        }

        paragraphRemoved = replaceRootStyles(
            REPLACEMENTS_PARAGRAPH,
            function () { return doc.paragraphStyles; },
            function () { return doc.allParagraphStyles; },
            missingTargetsParagraph,
            ambiguousTargetsParagraph
        );

        characterRemoved = replaceRootStyles(
            REPLACEMENTS_CHARACTER,
            function () { return doc.characterStyles; },
            function () { return doc.allCharacterStyles; },
            missingTargetsCharacter,
            ambiguousTargetsCharacter
        );

    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Replace Root Styles");

    var msg = "Done.\n\n";
    msg += "Paragraph styles: " + paragraphRemoved + " removed and replaced.";
    if (missingTargetsParagraph.length) {
        msg += "\nSkipped (target not found):\n" + missingTargetsParagraph.join("\n");
    }
    if (ambiguousTargetsParagraph.length) {
        msg += "\nAmbiguous targets (used first match):\n" + ambiguousTargetsParagraph.join("\n");
    }

    msg += "\n\nCharacter styles: " + characterRemoved + " removed and replaced.";
    if (missingTargetsCharacter.length) {
        msg += "\nSkipped (target not found):\n" + missingTargetsCharacter.join("\n");
    }
    if (ambiguousTargetsCharacter.length) {
        msg += "\nAmbiguous targets (used first match):\n" + ambiguousTargetsCharacter.join("\n");
    }

    alert(msg);
})();
