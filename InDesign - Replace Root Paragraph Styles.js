(function () {
    var doc = app.activeDocument;

    var REPLACEMENTS = [
        { from: "Normal",              to: "Body copy" },
        { from: "Heading 1",           to: "Heading 1" },        // root -> nested, same name
        { from: "IP - Bullet level 1", to: "Bullet list 1" }
        // add more { from: "...", to: "..." } pairs here over time
    ];

    var removed = 0;
    var missingTargets = [];   // pair "from" names whose "to" style wasn't found
    var ambiguousTargets = []; // "from -> to (Nx)" strings for >1 nested match

    // Root-level styles, by reference, so target resolution can exclude
    // them -- a target name may collide with a root-level source name
    // (e.g. "Heading 1" existing both at root and nested).
    var rootStyles = doc.paragraphStyles;
    var rootSet = [];
    for (var rs = 0; rs < rootStyles.length; rs++) rootSet.push(rootStyles[rs]);

    function isRootStyle(candidate) {
        for (var i = 0; i < rootSet.length; i++) {
            if (rootSet[i] === candidate) return true;
        }
        return false;
    }

    app.doScript(function () {
        var allPS = doc.allParagraphStyles;

        for (var p = 0; p < REPLACEMENTS.length; p++) {
            var pair = REPLACEMENTS[p];

            // Resolve the target, restricted to nested (non-root) styles.
            var nestedMatches = [];
            for (var a = 0; a < allPS.length; a++) {
                if (allPS[a].name === pair.to && !isRootStyle(allPS[a])) {
                    nestedMatches.push(allPS[a]);
                }
            }

            if (nestedMatches.length === 0) {
                missingTargets.push(pair.from + " -> " + pair.to);
                continue;
            }
            if (nestedMatches.length > 1) {
                ambiguousTargets.push(pair.from + " -> " + pair.to + " (" + nestedMatches.length + "x)");
            }
            var target = nestedMatches[0];

            // Resolve the source, root-only.
            var source = null;
            for (var rs2 = 0; rs2 < rootSet.length; rs2++) {
                if (rootSet[rs2].name === pair.from) { source = rootSet[rs2]; break; }
            }
            if (!source) continue; // nothing at root with this name -- not an error

            try {
                source.remove(target);
                removed++;
            } catch (e) {}
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Replace Root Paragraph Styles");

    var msg = "Done. " + removed + " root-level style" + (removed === 1 ? "" : "s") + " removed and replaced.";
    if (missingTargets.length) {
        msg += "\n\nSkipped (target not found):\n" + missingTargets.join("\n");
    }
    if (ambiguousTargets.length) {
        msg += "\n\nAmbiguous targets (used first match):\n" + ambiguousTargets.join("\n");
    }
    alert(msg);
})();
