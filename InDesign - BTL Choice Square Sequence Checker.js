// BTL Choice Square Sequence Checker (layout-agnostic)
// Checks that every named choice square exists, sits inside a table cell, and is in
// the right order relative to its neighbours: candidates read top to bottom within a
// group, and groups read left to right, row by row. No grid layout is assumed.

var doc = app.activeDocument;
var scriptTitle = "BP - BTLChoiceSquareGrp{n}Can{x} Sequence Checker";

// Format a number as a two-digit name fragment (01-09 or 10+)
function pad(n) {
    return (n <= 9) ? "0" + n : String(n);
}

// Climb up the object's structural hierarchy to locate its container Table Cell
function findCell(obj) {
    var currentParent = obj.parent;

    while (currentParent !== null) {
        var parentType = currentParent.constructor.name;

        if (parentType === "Cell") {
            return currentParent;
        }
        // Handle objects anchored inside a text frame/story line
        if (parentType === "Character" || parentType === "InsertionPoint" || parentType === "Story") {
            if (currentParent.parent.constructor.name === "Cell") {
                return currentParent.parent;
            }
        }
        // Keep climbing up if nested inside a group or specialized container frame
        if (currentParent.hasOwnProperty("parent")) {
            currentParent = currentParent.parent;
        } else {
            break;
        }
    }
    return null;
}

// Centre point and width of an object on the page, or null if it has no readable position
function getPosition(obj) {
    try {
        var b = obj.geometricBounds; // [top, left, bottom, right]
        return { x: (b[1] + b[3]) / 2, y: (b[0] + b[2]) / 2, width: b[3] - b[1] };
    } catch (e) {
        // e.g. anchored inside an overset table cell, so it isn't drawn on the page
        return null;
    }
}

function main() {
    // Prompt the user for the maximum number of groups (Starts completely empty)
    var maxInput = prompt("Enter the maximum number of groups to check:", "", scriptTitle);
    if (maxInput === null || maxInput === "") {
        return;
    }
    var maxGroups = parseInt(maxInput, 10);
    if (isNaN(maxGroups) || maxGroups <= 0) {
        alert("Please enter a valid positive number.", scriptTitle);
        return;
    }

    // Prompt the user for each group's candidate count, in group order
    var countsInput = prompt("Enter the number of candidates in each group, in order, separated by commas (e.g. 5,3,4):", "", scriptTitle);
    if (countsInput === null || countsInput === "") {
        return;
    }
    var parts = countsInput.replace(/\s/g, "").split(",");
    var candidateCounts = [];
    for (var p = 0; p < parts.length; p++) {
        if (!/^\d+$/.test(parts[p]) || parseInt(parts[p], 10) <= 0) {
            alert("Candidate counts must be positive whole numbers separated by commas.", scriptTitle);
            return;
        }
        candidateCounts.push(parseInt(parts[p], 10));
    }
    if (candidateCounts.length !== maxGroups) {
        alert("You entered " + maxGroups + " groups but " + candidateCounts.length + " candidate counts. These must match.", scriptTitle);
        return;
    }

    // Build a quick lookup map of all page items by name to avoid scanning the doc over and over
    var allItems = doc.allPageItems;
    var itemsMap = {};
    for (var k = 0; k < allItems.length; k++) {
        if (allItems[k].name && allItems[k].name !== "") {
            itemsMap[allItems[k].name] = allItems[k];
        }
    }

    var missing = [];
    var tableErrors = [];
    var orderErrors = [];
    var flagged = {}; // names of found squares with at least one error
    var totalFoundObjects = 0;
    var totalExpectedObjects = 0;
    var prevGroupTop = null; // top square of the previous group that had any squares

    for (var i = 1; i <= maxGroups; i++) {
        var groupTop = null;  // first found square in this group
        var prevSquare = null; // last found square in this group

        for (var j = 1; j <= candidateCounts[i - 1]; j++) {
            totalExpectedObjects++;
            var targetName = "BTLChoiceSquareGrp" + pad(i) + "Can" + pad(j);

            if (!itemsMap.hasOwnProperty(targetName)) {
                missing.push(targetName);
                continue;
            }
            totalFoundObjects++;
            var obj = itemsMap[targetName];

            // --- TABLE CHECK ---
            if (findCell(obj) === null) {
                tableErrors.push(targetName + " is floating loose (not inside a table cell).");
                flagged[targetName] = true;
            }

            var pos = getPosition(obj);
            if (pos === null) {
                orderErrors.push(targetName + "'s position could not be read (it may be in overset text).");
                flagged[targetName] = true;
                continue;
            }
            var square = { name: targetName, group: i, pos: pos };

            // --- CHECK 1: WITHIN GROUP (top to bottom, lined up) ---
            // Tolerance is half a square's width, so "close enough" scales with the layout
            if (prevSquare !== null) {
                var tol = prevSquare.pos.width / 2;
                if (pos.y - prevSquare.pos.y <= tol) {
                    orderErrors.push(targetName + " is not below " + prevSquare.name + ".");
                    flagged[targetName] = true;
                }
                if (Math.abs(pos.x - prevSquare.pos.x) > tol) {
                    orderErrors.push(targetName + " is not lined up under " + prevSquare.name + ".");
                    flagged[targetName] = true;
                }
            }
            if (groupTop === null) {
                groupTop = square;
            }
            prevSquare = square;
        }

        // --- CHECK 2: BETWEEN GROUPS (left to right, row by row) ---
        if (groupTop !== null) {
            if (prevGroupTop !== null) {
                var gTol = prevGroupTop.pos.width / 2;
                var dx = groupTop.pos.x - prevGroupTop.pos.x;
                var dy = groupTop.pos.y - prevGroupTop.pos.y;
                var nextInSameRow = Math.abs(dy) <= gTol && dx > gTol;
                var startOfNewRow = dy > gTol && dx < -gTol;
                if (!nextInSameRow && !startOfNewRow) {
                    orderErrors.push("Group " + pad(i) + " (" + groupTop.name + ") is out of reading order after Group " + pad(prevGroupTop.group) + " (" + prevGroupTop.name + ").");
                    flagged[groupTop.name] = true;
                }
            }
            prevGroupTop = groupTop;
        }
    }

    var flaggedCount = 0;
    for (var f in flagged) {
        if (flagged.hasOwnProperty(f)) {
            flaggedCount++;
        }
    }
    var perfectMatches = totalFoundObjects - flaggedCount;

    // Build the secondary detailed errors log text
    var reportLog = "";

    if (orderErrors.length > 0) {
        reportLog += "READING ORDER ERRORS:\n- " + orderErrors.join("\n- ") + "\n\n";
    }

    if (tableErrors.length > 0) {
        reportLog += "TABLE ERRORS:\n- " + tableErrors.join("\n- ") + "\n\n";
    }

    if (missing.length > 0) {
        reportLog += "MISSING NAMED SQUARES (Not Found):\n- " + missing.join("\n- ") + "\n\n";
    }

    // --- CUSTOM UI WINDOW ---
    if (missing.length === 0 && tableErrors.length === 0 && orderErrors.length === 0) {
        alert("Success! Perfect Reading Order.\nAll " + totalExpectedObjects + " candidate choice squares are present, inside table cells, and in order.", scriptTitle);
    } else {
        var resultWin = new Window("dialog", scriptTitle);
        resultWin.orientation = "column";
        resultWin.alignChildren = ["fill", "top"];
        resultWin.margins = 15;
        resultWin.spacing = 10;

        // Fixed Header Panel Module
        var summaryPanel = resultWin.add("panel", undefined, "Audit Status");
        summaryPanel.orientation = "column";
        summaryPanel.alignChildren = ["left", "top"];
        summaryPanel.margins = 12;

        summaryPanel.add("statictext", undefined, "• Objects Located: " + totalFoundObjects + " / " + totalExpectedObjects + " expected");
        summaryPanel.add("statictext", undefined, "• In Order & In Tables: " + perfectMatches + " / " + totalExpectedObjects);

        // Large Scrollable Box for detailed issue logs below it
        resultWin.add("statictext", undefined, "Detailed Issue Breakdown:");
        resultWin.add("edittext", undefined, reportLog, {multiline: true, readonly: true});

        // OK Close button at the very bottom
        var btnGroup = resultWin.add("group");
        btnGroup.alignment = "center";
        var okBtn = btnGroup.add("button", undefined, "OK");
        okBtn.onClick = function() { resultWin.close(); };

        resultWin.show();
    }
}

main();
