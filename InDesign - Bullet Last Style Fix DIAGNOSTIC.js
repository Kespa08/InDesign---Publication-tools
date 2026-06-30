// ── DIAGNOSTIC — run this to see what the fix script actually finds ──────────
// It makes NO changes to the document. It reports:
//   • Which document is active
//   • How many stories and paragraphs exist
//   • The first 30 paragraph style names it reads
//   • How many match the bullet pattern
//   • What corrections WOULD be applied
// ─────────────────────────────────────────────────────────────────────────────
(function () {

    var doc = app.activeDocument;

    function isBulletStyle(name) {
        return /^Bullet list\s+\d+/i.test(name);
    }
    function isLastVariant(name) {
        return /\s+last$/i.test(name);
    }

    var storyCount    = 0;
    var paraCount     = 0;
    var bulletCount   = 0;
    var sampleNames   = [];   // first 30 style names seen
    var wouldChange   = [];   // corrections the fix script would make

    try {
        var stories = doc.stories;
        storyCount = stories.length;

        for (var si = 0; si < stories.length; si++) {
            var paras;
            try { paras = stories[si].paragraphs; } catch (e) { continue; }
            paraCount += paras.length;

            for (var pi = 0; pi < paras.length; pi++) {
                var styleName;
                try {
                    styleName = paras[pi].appliedParagraphStyle.name;
                } catch (e) {
                    styleName = "<<ERROR reading style: " + e.message + ">>";
                }

                if (sampleNames.length < 30) sampleNames.push(styleName);

                if (!isBulletStyle(styleName)) continue;
                bulletCount++;

                var nextIsBullet = false;
                if (pi + 1 < paras.length) {
                    try {
                        nextIsBullet = isBulletStyle(
                            paras[pi + 1].appliedParagraphStyle.name
                        );
                    } catch (e) {}
                }

                var targetName;
                if (!nextIsBullet) {
                    targetName = isLastVariant(styleName) ? styleName : styleName + " last";
                } else {
                    targetName = isLastVariant(styleName)
                        ? styleName.replace(/\s+last$/i, "")
                        : styleName;
                }

                if (targetName !== styleName && wouldChange.length < 20) {
                    wouldChange.push(
                        "Story " + si + ", Para " + pi + ":\n" +
                        "  \"" + styleName + "\"  →  \"" + targetName + "\""
                    );
                }
            }
        }
    } catch (e) {
        alert("Outer error: " + e.message);
        return;
    }

    // ── Build report ─────────────────────────────────────────────────────────
    var msg = "DOCUMENT: " + doc.name + "\n";
    msg += "Stories: " + storyCount + "  |  Total paragraphs: " + paraCount + "\n";
    msg += "Bullet paragraphs found: " + bulletCount + "\n";
    msg += "Corrections that WOULD be made: " + wouldChange.length + "\n\n";

    if (wouldChange.length > 0) {
        msg += "── Pending corrections ──────────────────\n";
        msg += wouldChange.join("\n") + "\n\n";
    }

    msg += "── First 30 style names seen ────────────\n";
    for (var i = 0; i < sampleNames.length; i++) {
        msg += (i + 1) + ". \"" + sampleNames[i] + "\"\n";
    }

    alert(msg);

})();
