// InDesign Script: Word/RTF Import — Final with:
// - Tagged Table Conversion
// - Case Study Insertion (correct placement)
// - Robust Hyperlink Detection
// - Headings, Bullets, Bold
// - Table Cross-Page Detection and Skipping
//
// Supports:
//   [TABLE] ... [/TABLE]
//   [CASESTUDY] / [CASE STUDY] ... [/CASESTUDY] / [/CASE STUDY]

#target "InDesign"

if (app.documents.length === 0) {
    alert("Please open an InDesign template first!");
    exit();
}

var doc = app.activeDocument;
var originalInteraction = app.scriptPreferences.userInteractionLevel;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

// -------------------------------------------------------------
// 1. CASE STUDY INSERTION — MUST RUN EARLY (BEFORE STORY CHANGES)
// -------------------------------------------------------------
function insertCaseStudies(storyRef) {

    var re = /\[CASE ?STUDY\]([\s\S]*?)\[\/CASE ?STUDY\]/g;
    var matches = [], m;
    var fullText = storyRef.contents;

    while ((m = re.exec(fullText)) !== null) {
        matches.push({
            start: m.index,
            fullLen: m[0].length,
            inner: m[1]
        });
    }

    if (!matches.length) return;

    // Process from end → start (stabilizes character indices)
    for (var i = matches.length - 1; i >= 0; i--) {
        var block = matches[i];

        var path = block.inner.replace(/^\s+|\s+$/g, "");
        if (!path) continue;

        var inddFile = File(path);
        if (!inddFile.exists) {
            alert("Case Study INDD file not found:\n" + path);
            continue;
        }

        var start = block.start;
        var end = start + block.fullLen - 1;

        // Capture the insertion page BEFORE removing anything
        var parentPage = null;
        try {
            var ip = storyRef.characters[start];
            var tf = ip.parentTextFrames[0];
            parentPage = tf.parentPage;
        } catch (_) {
            parentPage = null;
        }

        // Remove the case study tag block
        var range = storyRef.characters.itemByRange(start, end);
        range.remove();

        // Open referenced INDD silently
        var sourceDoc = app.open(inddFile, false);

        try {
            if (parentPage) {
                var insertAfter = parentPage;

                // Duplicate pages one-by-one
                for (var p = 0; p < sourceDoc.pages.length; p++) {
                    var newPage =
                        sourceDoc.pages[p].duplicate(LocationOptions.AFTER, insertAfter);

                    insertAfter = newPage;
                }

            } else {
                // Fallback: append at end
                for (var p2 = 0; p2 < sourceDoc.pages.length; p2++) {
                    sourceDoc.pages[p2].duplicate(LocationOptions.AT_END, doc.pages[-1]);
                }
            }

        } catch (e) {
            $.writeln("Case study insertion error: " + e);
        } finally {
            sourceDoc.close(SaveOptions.NO);
        }
    }
}

// -------------------------------------------------------------
// MAIN SCRIPT
// -------------------------------------------------------------
try {

    // ---- 1) Select file ----
    var file = File.openDialog("Select Word (.docx) or RTF file", "*.docx;*.rtf", false);
    if (!file || !file.exists) throw new Error("No file selected.");

    // ---- 2) Place file into first frame ----
    function makeFrame(pg) {
        var m = pg.marginPreferences, b = pg.bounds;
        return pg.textFrames.add({
            geometricBounds: [b[0]+m.top, b[1]+m.left, b[2]-m.bottom, b[3]-m.right]
        });
    }

    var page = doc.pages[0];
    var frame = page.textFrames.length ? page.textFrames[0] : makeFrame(page);
    frame.place(file);

    var story = frame.parentStory;
    if (!story) throw new Error("No story found after placing file.");

    // ---- 2b) CRITICAL — INSERT CASE STUDIES NOW (BEFORE ANY REFLOW) ----
    insertCaseStudies(story);

    // ---- 2c) Check for tables that cross pages and skip them ----
    // Helper function to safely get parent text frame of a table
    function getTableParentFrame(table) {
        try {
            // Method 1: Try to get parent directly
            try {
                var parent = table.parent;
                if (parent) {
                    // Check if parent is a TextFrame (has geometricBounds and overflows)
                    try {
                        if (parent.geometricBounds && typeof parent.overflows !== 'undefined') {
                            return parent;
                        }
                    } catch (e) {}
                    
                    // If parent is a Cell, traverse up
                    try {
                        if (parent.constructor && parent.constructor.name === "Cell") {
                            var cellParent = parent.parent;
                            if (cellParent) {
                                // Cell's parent might be a Table or TextFrame
                                try {
                                    if (cellParent.geometricBounds && typeof cellParent.overflows !== 'undefined') {
                                        return cellParent;
                                    }
                                } catch (e) {}
                                // If cell's parent is another table, go up one more level
                                if (cellParent.constructor && cellParent.constructor.name === "Table") {
                                    var tableParent = cellParent.parent;
                                    if (tableParent && tableParent.geometricBounds) {
                                        return tableParent;
                                    }
                                }
                            }
                        }
                    } catch (e) {}
                }
            } catch (e) {
                $.writeln("Error getting table.parent: " + e);
            }
            
            // Method 2: Get table's first cell and find its text frame
            try {
                var firstCell = table.cells[0];
                if (firstCell) {
                    var cellTexts = firstCell.texts;
                    if (cellTexts.length > 0) {
                        var cellText = cellTexts[0];
                        if (cellText && cellText.parentTextFrames && cellText.parentTextFrames.length > 0) {
                            return cellText.parentTextFrames[0];
                        }
                    }
                }
            } catch (e) {
                $.writeln("Error getting cell text frame: " + e);
            }
            
            // Method 3: Use story's text containers (find the one that contains this table)
            try {
                var tableStory = table.parentStory;
                if (tableStory && tableStory.textContainers.length > 0) {
                    // Return the first text container (usually the main frame)
                    // This is a fallback - not perfect but better than null
                    for (var i = 0; i < tableStory.textContainers.length; i++) {
                        try {
                            var container = tableStory.textContainers[i];
                            if (container.geometricBounds) {
                                return container;
                            }
                        } catch (e) {}
                    }
                }
            } catch (e) {
                $.writeln("Error searching text containers: " + e);
            }
            
            // Method 4: Last resort - get the first text frame from the document
            try {
                if (doc.pages.length > 0) {
                    var firstPage = doc.pages[0];
                    if (firstPage.textFrames.length > 0) {
                        return firstPage.textFrames[0];
                    }
                }
            } catch (e) {
                $.writeln("Error getting first page text frame: " + e);
            }
        } catch (e) {
            $.writeln("Error in getTableParentFrame: " + e);
        }
        return null;
    }
    
    // Helper function to check if a table crosses pages
    function tableCrossesPages(table) {
        try {
            var parentFrame = getTableParentFrame(table);
            if (!parentFrame) {
                $.writeln("  Cannot determine if table crosses pages - no parent frame");
                return false;
            }
            
            var frameH = parentFrame.geometricBounds[2] - parentFrame.geometricBounds[0];
            var tableH = table.height;
            
            // Check 1: Table height > frame height
            if (tableH > frameH) {
                $.writeln("  Table crosses pages: height " + tableH.toFixed(2) + "pt > frame " + frameH.toFixed(2) + "pt");
                return true;
            }
            
            // Check 2: Table spans multiple text containers
            try {
                var tableStory = table.parentStory;
                if (tableStory) {
                    // Get all text containers that contain this table
                    var containersWithTable = [];
                    for (var c = 0; c < tableStory.textContainers.length; c++) {
                        try {
                            var container = tableStory.textContainers[c];
                            // Check if this container has the table by checking if table's first cell is in it
                            var firstCell = table.cells[0];
                            if (firstCell && firstCell.texts.length > 0) {
                                var cellText = firstCell.texts[0];
                                if (cellText.parentTextFrames && cellText.parentTextFrames.length > 0) {
                                    for (var tf = 0; tf < cellText.parentTextFrames.length; tf++) {
                                        if (cellText.parentTextFrames[tf] === container) {
                                            containersWithTable.push(container);
                                            break;
                                        }
                                    }
                                }
                            }
                        } catch (e) {}
                    }
                    
                    // If table is in multiple containers, it crosses pages
                    if (containersWithTable.length > 1) {
                        $.writeln("  Table crosses pages: spans " + containersWithTable.length + " text containers");
                        return true;
                    }
                }
            } catch (e) {
                $.writeln("  Could not check text containers: " + e);
            }
            
            return false;
        } catch (e) {
            $.writeln("  Error checking if table crosses pages: " + e);
            return false;
        }
    }
    
    // Helper function to remove table and insert marker
    function removeTableAndInsertMarker(table, tableIndex) {
        try {
            $.writeln("  Removing table " + tableIndex + " and inserting marker...");
            
            // Find insertion point BEFORE removing table - get the first character of the table's first cell
            var insertIndex = -1;
            var markerText = "\r<<TABLE HAS BEEN SKIPPED. PLEASE INSERT MANUALLY>>\r";
            
            try {
                var firstCell = table.cells[0];
                if (firstCell && firstCell.texts.length > 0) {
                    var firstText = firstCell.texts[0];
                    if (firstText.characters.length > 0) {
                        insertIndex = firstText.characters[0].index;
                    }
                }
            } catch (e) {
                $.writeln("    Could not get table insertion point: " + e);
            }
            
            // Insert marker text BEFORE removing table (so index is still valid)
            if (insertIndex >= 0) {
                try {
                    story.insertionPoints[insertIndex].contents = markerText;
                    $.writeln("    Marker text inserted at index " + insertIndex);
                } catch (e) {
                    $.writeln("    Error inserting marker text at index: " + e);
                    insertIndex = -1; // Will use fallback
                }
            }
            
            // Delete the table
            try {
                table.remove();
                $.writeln("    Table removed successfully");
            } catch (e) {
                $.writeln("    Error removing table: " + e);
                // If marker was inserted but table removal failed, that's okay - marker is there
                if (insertIndex >= 0) return true;
                return false;
            }
            
            // If marker wasn't inserted before removal, try fallback methods
            if (insertIndex < 0) {
                // Try to find where table was by looking for text before/after
                try {
                    // Find the last non-overflowing frame and insert there
                    for (var f = story.textContainers.length - 1; f >= 0; f--) {
                        try {
                            var container = story.textContainers[f];
                            if (!container.overflows && container.parentStory) {
                                var lastChar = container.parentStory.characters[-1];
                                story.insertionPoints[lastChar.index + 1].contents = markerText;
                                $.writeln("    Marker text inserted after last placed content (fallback)");
                                return true;
                            }
                        } catch (e) {}
                    }
                } catch (e) {
                    $.writeln("    Fallback insertion failed: " + e);
                }
                
                // Last resort: insert at end of story
                try {
                    story.insertionPoints[-1].contents = markerText;
                    $.writeln("    Marker text inserted at end of story (last resort)");
                    return true;
                } catch (e) {
                    $.writeln("    Could not insert marker text: " + e);
                    return false;
                }
            }
            
            return true;
        } catch (e) {
            $.writeln("  Error in removeTableAndInsertMarker: " + e);
            return false;
        }
    }
    
    // Check all tables and skip those that cross pages
    // Process in reverse order to maintain indices
    var tablesToSkip = [];
    var tableCount = story.tables.length;
    $.writeln("Found " + tableCount + " table(s) in document");
    
    if (tableCount > 0) {
        // Check each table
        for (var ts = 0; ts < story.tables.length; ts++) {
            try {
                var checkTbl = story.tables[ts];
                $.writeln("Checking table " + ts + " for page crossing...");
                
                if (tableCrossesPages(checkTbl)) {
                    $.writeln("  ✓ Table " + ts + " crosses pages - will be skipped");
                    tablesToSkip.push({table: checkTbl, index: ts});
                } else {
                    $.writeln("  ✓ Table " + ts + " fits on one page - will be imported");
                }
            } catch (e) {
                $.writeln("  Error checking table " + ts + ": " + e);
            }
        }
        
        // Remove tables that cross pages (process in reverse to maintain indices)
        if (tablesToSkip.length > 0) {
            $.writeln("Found " + tablesToSkip.length + " table(s) that cross pages - skipping them");
            
            for (var skipIdx = tablesToSkip.length - 1; skipIdx >= 0; skipIdx--) {
                var skipTable = tablesToSkip[skipIdx];
                if (removeTableAndInsertMarker(skipTable.table, skipTable.index)) {
                    $.writeln("✓ Skipped table " + skipTable.index + " - marker inserted");
                } else {
                    $.writeln("✗ Failed to skip table " + skipTable.index);
                }
            }
            
            // Recompose after removing tables
            story.recompose();
            $.writeln("Skipped " + tablesToSkip.length + " table(s) that cross pages");
            
            // Inform user about skipped tables
            var skippedMsg = "⚠️ Table Import Notice\n\n";
            skippedMsg += tablesToSkip.length === 1 
                ? "1 table that crosses page boundaries has been skipped.\n\n"
                : tablesToSkip.length + " tables that cross page boundaries have been skipped.\n\n";
            skippedMsg += "These tables have been replaced with the marker:\n";
            skippedMsg += "<<TABLE HAS BEEN SKIPPED. PLEASE INSERT MANUALLY>>\n\n";
            skippedMsg += "Please import these tables manually.";
            alert(skippedMsg);
        }
    }

    // ---- Continue threading pages ----
    while (story.overflows) {
        var last = doc.pages[-1];
        var np = doc.pages.add(LocationOptions.AFTER, last);
        var nf = makeFrame(np);
        story.textContainers[story.textContainers.length - 1].nextTextFrame = nf;
    }

    // ---- 3) Load styles ----
    function s(n, t) {
        try {
            return (t == "p") ? doc.paragraphStyles.itemByName(n)
                              : doc.characterStyles.itemByName(n);
        } catch (_) { return null; }
    }

    var ps = {
        h: [
            s("Heading 1", "p"),
            s("Heading 2", "p"),
            s("Heading 3", "p"),
            s("Heading 4", "p"),
            s("Heading 5", "p"),
            s("Heading 6", "p")
        ],
        body:   s("Body", "p"),
        bullets:s("Bullets", "p"),
        sub:    s("Sub-Bullets", "p"),
        num:    s("Numbered", "p"),
        table:  s("Table Style", "p")
    };

    var cs = {
        bold: s("Body - Bold", "c"),
        link: s("Hyperlink Highlight", "c")
    };

    // ---- 4) Collapse blank lines ----
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.findWhat = "(?<=\\r)\\r+";
    app.changeGrepPreferences.changeTo = "";
    story.changeGrep();

    // ---- 5) Apply paragraph styles ----
    for (var i=0; i<story.paragraphs.length; i++) {
        var p = story.paragraphs[i], n = "";
        try { n = p.appliedParagraphStyle.name; } catch(_) {}

        if (/Heading\s*1/i.test(n) && ps.h[0]) p.appliedParagraphStyle = ps.h[0];
        else if (/Heading\s*2/i.test(n) && ps.h[1]) p.appliedParagraphStyle = ps.h[1];
        else if (/Heading\s*3/i.test(n) && ps.h[2]) p.appliedParagraphStyle = ps.h[2];
        else if (/Heading\s*4/i.test(n) && ps.h[3]) p.appliedParagraphStyle = ps.h[3];
        else if (/Heading\s*5/i.test(n) && ps.h[4]) p.appliedParagraphStyle = ps.h[4];
        else if (/Heading\s*6/i.test(n) && ps.h[5]) p.appliedParagraphStyle = ps.h[5];
        else if (/Normal/i.test(n) && ps.body)     p.appliedParagraphStyle = ps.body;
        else if (ps.body)                           p.appliedParagraphStyle = ps.body;
    }

    // ---- 6) Bullets/Sub-Bullets ----
    var bulletGlyphs=/^[\u2022\u25CF\u25E6\uF0B7\u2219\u00B7○◦\-–]+\s*/;

    for (var j=0; j<story.paragraphs.length; j++) {
        var para = story.paragraphs[j], txt = para.contents;
        if (bulletGlyphs.test(txt)) {
            var indent = (txt.match(/^\t+/)||[""])[0].length;
            para.contents = txt.replace(bulletGlyphs,"").replace(/^\t+/,'');
            if (indent>0 && ps.sub) para.appliedParagraphStyle = ps.sub;
            else if (ps.bullets)    para.appliedParagraphStyle = ps.bullets;
        }
    }

    // ---- 7) Bold → Character Style ----
    if (cs.bold) {
        for (var c=0; c<story.characters.length; c++) {
            var ch = story.characters[c];
            try {
                if (ch.fontStyle && /bold/i.test(ch.fontStyle))
                    ch.appliedCharacterStyle = cs.bold;
            } catch(_){}
        }
    }

    // ---- 8) Apply table style to native Word tables ----
    if (ps.table && story.tables.length) {
        for (var t=0; t<story.tables.length; t++) {
            try { story.tables[t].appliedTableStyle = ps.table; } catch(_){}
        }
    }

    // ---- 9) Convert [TABLE]...[/TABLE] blocks ----
    function convertTaggedTables(storyRef, tableStyleName) {
        var re = /\[TABLE\]([\s\S]*?)\[\/TABLE\]/g;
        var matches = [], m;
        var full = storyRef.contents;

        while ((m = re.exec(full)) !== null) {
            matches.push({ start:m.index, full:m[0], inner:m[1] });
        }
        if (!matches.length) return;

        var style=null;
        try { style = doc.tableStyles.itemByName(tableStyleName); } catch(_){}

        for (var k=matches.length-1; k>=0; k--) {
            var block = matches[k];
            var start = block.start;
            var end   = start + block.full.length - 1;

            var tableText = block.inner
                .replace(/ +/g, "\t")
                .replace(/\r{2,}/g, "\r");

            var range = storyRef.characters.itemByRange(start, end);
            range.contents = tableText;

            var tbl = range.convertToTable("\t","\r");
            if (style && style.isValid) tbl.appliedTableStyle = style;
            try { tbl.rows[0].rowType = RowTypes.HEADER_ROW; } catch(_){}
        }
    }

    convertTaggedTables(story, "Table Style");

    // ---- 10) Hyperlinks -------
    if (cs.link) {
        // True hyperlinks
        for (var h=0; h<doc.hyperlinks.length; h++) {
            try {
                var l = doc.hyperlinks[h], d=l.destination, src=l.source;
                if (d && d.destinationURL &&
                    /^(https?|mailto):/i.test(d.destinationURL) &&
                    src instanceof HyperlinkTextSource)
                {
                    src.sourceText.appliedCharacterStyle = cs.link;
                }
            } catch(_){}
        }

        // Visible URLs
        var urlPattern = /\b((https?|ftp|file|mailto):\/\/[^\s<>"']+|www\.[^\s<>"']+)\b/g;
        var txt = story.contents, hits=[], mm;
        while ((mm=urlPattern.exec(txt)) !== null) {
            hits.push({start:mm.index, end:mm.index+mm[0].length});
        }
        for (var h2=0; h2<hits.length; h2++) {
            try {
                var r = story.characters.itemByRange(hits[h2].start, hits[h2].end-1);
                r.appliedCharacterStyle = cs.link;
            } catch(_){}
        }
    }

    story.recompose();
    alert("✅ Import complete.\nHeadings, Bullets, Bold, Tables, Case Studies, and Hyperlinks applied.");

} catch(err) {
    alert("❌ " + err.message);
    $.writeln("❌ " + err.message);
} finally {
    app.scriptPreferences.userInteractionLevel = originalInteraction;
}
