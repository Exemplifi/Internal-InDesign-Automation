// InDesign Script: Word/RTF Import — Final with:
// - Tagged Table Conversion
// - Case Study Insertion (correct placement)
// - Robust Hyperlink Detection
// - Headings, Bullets, Bold
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

    if (!matches.length) {
        $.writeln("  No case studies found");
        return;
    }
    
    $.writeln("  Found " + matches.length + " case study tag(s)");

    // Process from end → start (stabilizes character indices)
    for (var i = matches.length - 1; i >= 0; i--) {
        var block = matches[i];

        var path = block.inner.replace(/^\s+|\s+$/g, "");
        if (!path) continue;

        var inddFile = File(path);
        if (!inddFile.exists) {
            $.writeln("  ERROR: Case Study INDD file not found: " + path);
            alert("Case Study INDD file not found:\n" + path);
            continue;
        }
        $.writeln("  Processing case study: " + path);

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
    $.writeln("=== Starting InDesign Import Script ===");
    var file = File.openDialog("Select Word (.docx) or RTF file", "*.docx;*.rtf", false);
    if (!file || !file.exists) throw new Error("No file selected.");
    $.writeln("File selected: " + file.name);

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
    $.writeln("File placed. Story length: " + story.characters.length + " characters, " + story.paragraphs.length + " paragraphs");

    // ---- 2b) CRITICAL — INSERT CASE STUDIES NOW (BEFORE ANY REFLOW) ----
    $.writeln("Checking for case studies...");
    insertCaseStudies(story);

    // ---- 2c) Fix table widths BEFORE threading to prevent overflow issues ----
    $.writeln("Found " + story.tables.length + " table(s) in document");
    if (story.tables.length) {
        for (var t=0; t<story.tables.length; t++) {
            try {
                var tbl = story.tables[t];
                var parentFrame = tbl.parentTextFrames[0];
                if (parentFrame) {
                    var frameWidth = parentFrame.geometricBounds[3] - parentFrame.geometricBounds[1];
                    var tableWidth = tbl.width;
                    // If table is wider than frame (or close to it), resize it to fit
                    // Use 90% threshold to catch tables that are just slightly too wide
                    if (tableWidth > frameWidth * 0.9) {
                        var newWidth = frameWidth - 20; // Leave small margin
                        tbl.width = newWidth;
                        $.writeln("Pre-threading: Resized table " + t + " from " + tableWidth.toFixed(2) + " to " + newWidth.toFixed(2) + " (frame: " + frameWidth.toFixed(2) + ")");
                    } else {
                        $.writeln("Pre-threading: Table " + t + " width OK (" + tableWidth.toFixed(2) + " <= " + (frameWidth * 0.9).toFixed(2) + ")");
                    }
                } else {
                    $.writeln("Pre-threading: Table " + t + " has no parent frame");
                }
            } catch (e) {
                $.writeln("Could not check/resize table " + t + " before threading: " + e);
            }
        }
        story.recompose(); // Update overflow state after table resizing
        $.writeln("After table resizing: Overflow = " + (story.overflows ? "YES" : "NO"));
    }

    // ---- Continue threading pages with safety limits ----
    var maxPages = 100; // Safety limit to prevent infinite loops
    var pageCount = 0;
    var initialPageCount = doc.pages.length;
    var lastOverflowState = story.overflows;
    var stableIterations = 0;
    
    $.writeln("Initial page count: " + initialPageCount);
    $.writeln("Initial overflow state: " + (story.overflows ? "OVERFLOW" : "NO OVERFLOW"));
    
    if (story.overflows) {
        $.writeln("Starting page threading loop (max " + maxPages + " pages)...");
    }
    
    while (story.overflows && pageCount < maxPages) {
        var last = doc.pages[-1];
        var np = doc.pages.add(LocationOptions.AFTER, last);
        var nf = makeFrame(np);
        
        // Get the last text container and thread it
        // Use try-catch to handle cases where last container isn't a TextFrame
        try {
            var containers = story.textContainers;
            if (containers.length > 0) {
                var lastContainer = containers[containers.length - 1];
                // Try to thread the last container
                try {
                    lastContainer.nextTextFrame = nf;
                    // Successfully threaded - log occasionally
                    if (pageCount <= 3 || pageCount % 10 === 0) {
                        $.writeln("  Iteration " + pageCount + ": Successfully threaded last container");
                    }
                } catch (e1) {
                    // If last container isn't threadable (e.g., it's a table), 
                    // search backwards for the last text frame
                    var found = false;
                    for (var tc = containers.length - 2; tc >= 0; tc--) {
                        try {
                            containers[tc].nextTextFrame = nf;
                            found = true;
                            break;
                        } catch (e2) {
                            continue;
                        }
                    }
                    if (!found) {
                        $.writeln("Warning: No threadable text frame found in iteration " + pageCount + " (containers: " + containers.length + ")");
                    } else {
                        $.writeln("  Iteration " + pageCount + ": Threaded from container " + (containers.length - 2 - tc) + " (last container was not threadable)");
                    }
                }
            }
        } catch (e) {
            $.writeln("Threading error: " + e);
            // If threading fails completely, break to prevent infinite loop
            break;
        }
        
        // Force recompose to update overflow state
        story.recompose();
        
        pageCount++;
        
        // Log progress every 10 pages or on first few iterations
        if (pageCount <= 3 || pageCount % 10 === 0) {
            $.writeln("  Iteration " + pageCount + ": Pages=" + doc.pages.length + ", Overflow=" + (story.overflows ? "YES" : "NO") + ", Containers=" + story.textContainers.length);
        }
        
        // Check if overflow state changed - if it's been stable, investigate
        if (story.overflows === lastOverflowState) {
            stableIterations++;
            
            // If overflow hasn't changed for 2 iterations, check if it's a table blocking flow
            if (stableIterations >= 2) {
                $.writeln("Warning: Overflow state not changing. Checking for blocking tables...");
                
                // Check if there are tables that might be blocking flow
                var blockingTable = false;
                try {
                    for (var ti = 0; ti < story.tables.length; ti++) {
                        var t = story.tables[ti];
                        var tf = t.parentTextFrames[0];
                        if (tf && tf.overflows) {
                            // Table is in an overflowing frame - try to resize it more aggressively
                            try {
                                var frameWidth = tf.geometricBounds[3] - tf.geometricBounds[1];
                                if (t.width > frameWidth * 0.9) {
                                    t.width = frameWidth * 0.85; // More aggressive resize
                                    $.writeln("  Aggressively resized blocking table " + ti + " to " + t.width);
                                    story.recompose();
                                    blockingTable = true;
                                }
                            } catch (e) {
                                $.writeln("  Could not resize blocking table " + ti + ": " + e);
                            }
                        }
                    }
                } catch (e) {
                    $.writeln("  Error checking for blocking tables: " + e);
                }
                
                // If we fixed a table, reset stable iterations and continue
                if (blockingTable) {
                    stableIterations = 0;
                    lastOverflowState = story.overflows;
                    continue; // Try again with resized table
                }
                
                // If still no change after table fix attempt, continue for a few more iterations
                // to allow content after the table to flow
                if (stableIterations >= 5) {
                    $.writeln("Warning: Overflow persists after " + stableIterations + " stable iterations. Continuing for content after table...");
                    // Don't break yet - allow a few more iterations for text after table
                }
                
                // Only break if we've tried many times and overflow still persists
                if (stableIterations >= 10) {
                    $.writeln("Breaking after " + stableIterations + " stable iterations. Content may still be in overflow.");
                    break;
                }
            }
        } else {
            stableIterations = 0;
            lastOverflowState = story.overflows;
        }
    }
    
    // Final recompose
    story.recompose();
    
    $.writeln("Page threading complete. Iterations: " + pageCount + ", Final pages: " + doc.pages.length + ", Final overflow: " + (story.overflows ? "YES" : "NO"));
    
    // If there's still overflow, try additional strategies to place remaining content
    if (story.overflows && pageCount < maxPages) {
        $.writeln("Attempting to place remaining overflow content...");
        
        // Try creating a few more pages with different strategy
        var additionalAttempts = 0;
        var maxAdditional = 20; // Limit additional attempts
        
        while (story.overflows && additionalAttempts < maxAdditional && (pageCount + additionalAttempts) < maxPages) {
            try {
                var last = doc.pages[-1];
                var np = doc.pages.add(LocationOptions.AFTER, last);
                var nf = makeFrame(np);
                
                // Try to thread from the last text container
                var containers = story.textContainers;
                if (containers.length > 0) {
                    // Try all containers from last to first
                    var threaded = false;
                    for (var ac = containers.length - 1; ac >= 0; ac--) {
                        try {
                            containers[ac].nextTextFrame = nf;
                            threaded = true;
                            $.writeln("  Additional attempt " + (additionalAttempts + 1) + ": Threaded container " + ac);
                            break;
                        } catch (e) {
                            continue;
                        }
                    }
                    
                    if (!threaded) {
                        $.writeln("  Additional attempt " + (additionalAttempts + 1) + ": Could not thread any container");
                        break; // No point continuing if we can't thread
                    }
                }
                
                story.recompose();
                additionalAttempts++;
                
                // Check if we made progress
                if (!story.overflows) {
                    $.writeln("  Successfully placed remaining content after " + additionalAttempts + " additional attempts");
                    break;
                }
            } catch (e) {
                $.writeln("  Error in additional attempt " + (additionalAttempts + 1) + ": " + e);
                break;
            }
        }
        
        pageCount += additionalAttempts;
        $.writeln("Additional attempts: " + additionalAttempts + ", Total iterations: " + pageCount);
    }
    
    // Final recompose after additional attempts
    story.recompose();
    
    // Check for tables that couldn't be placed and add marker text
    if (story.overflows && story.tables.length > 0) {
        $.writeln("Checking for unplaced tables and adding markers...");
        var unplacedTables = [];
        
        for (var ut = 0; ut < story.tables.length; ut++) {
            try {
                var tbl = story.tables[ut];
                var parentFrame = tbl.parentTextFrames[0];
                
                // Check if table is in an overflowing frame or if we can't access it
                var isUnplaced = false;
                try {
                    if (parentFrame && parentFrame.overflows) {
                        isUnplaced = true;
                    }
                } catch (e) {
                    // If we can't access the parent frame, table might be in overflow
                    isUnplaced = true;
                }
                
                if (isUnplaced) {
                    unplacedTables.push({
                        table: tbl,
                        index: ut,
                        parentFrame: parentFrame
                    });
                    $.writeln("  Table " + ut + " appears to be unplaced");
                }
            } catch (e) {
                $.writeln("  Error checking table " + ut + ": " + e);
            }
        }
        
        // Insert marker text for each unplaced table
        if (unplacedTables.length > 0) {
            // Process from end to start to maintain character indices
            for (var ui = unplacedTables.length - 1; ui >= 0; ui--) {
                try {
                    var unplaced = unplacedTables[ui];
                    var tbl = unplaced.table;
                    var markerInserted = false;
                    
                    // Try multiple methods to find insertion point
                    try {
                        // Method 1: Get the table's first cell and find its position
                        var tableCells = tbl.cells;
                        if (tableCells.length > 0) {
                            var firstCell = tableCells[0];
                            if (firstCell.texts.length > 0) {
                                var firstText = firstCell.texts[0];
                                if (firstText.characters.length > 0) {
                                    var firstChar = firstText.characters[0];
                                    var insertIndex = firstChar.index;
                                    
                                    // Insert marker text before the table
                                    var markerText = "\r<<TABLE NOT EMBEDDED. PLEASE PLACE MANUALLY>>\r";
                                    story.insertionPoints[insertIndex].contents = markerText;
                                    markerInserted = true;
                                    $.writeln("  Inserted marker text before table " + unplaced.index + " (method 1)");
                                }
                            }
                        }
                    } catch (e1) {
                        // Method 1 failed, try method 2
                        try {
                            // Method 2: Try to get table's parent text frame and find position
                            var parentFrame = tbl.parentTextFrames[0];
                            if (parentFrame) {
                                var frameStart = parentFrame.parentStory.characters[0];
                                // Find where table starts by looking for it in the frame
                                // This is a fallback - insert at a safe position
                                var markerText = "\r<<TABLE " + (unplaced.index + 1) + " NOT EMBEDDED. PLEASE PLACE MANUALLY>>\r";
                                // Insert at the end of the last placed content
                                var lastPlacedFrame = null;
                                for (var f = 0; f < story.textContainers.length; f++) {
                                    try {
                                        var tf = story.textContainers[f];
                                        // Check if it's a text frame by trying to access overflows property
                                        try {
                                            var overflows = tf.overflows;
                                            if (!overflows) {
                                                lastPlacedFrame = tf;
                                            }
                                        } catch (e) {
                                            // Not a text frame or can't access overflows
                                        }
                                    } catch (e) {}
                                }
                                
                                if (lastPlacedFrame) {
                                    try {
                                        var lastChar = lastPlacedFrame.parentStory.characters[-1];
                                        story.insertionPoints[lastChar.index + 1].contents = markerText;
                                        markerInserted = true;
                                        $.writeln("  Inserted marker text after last placed content for table " + unplaced.index + " (method 2)");
                                    } catch (e) {
                                        // Continue to method 3
                                    }
                                }
                            }
                        } catch (e2) {
                            // Method 2 failed, use method 3
                        }
                    }
                    
                    // Method 3: Fallback - insert at end of story
                    if (!markerInserted) {
                        try {
                            var markerText = "\r<<TABLE " + (unplaced.index + 1) + " NOT EMBEDDED. PLEASE PLACE MANUALLY>>\r";
                            story.insertionPoints[-1].contents = markerText;
                            $.writeln("  Inserted marker text at end of story for table " + unplaced.index + " (method 3)");
                        } catch (e3) {
                            $.writeln("  Could not insert marker text for table " + unplaced.index + " (all methods failed): " + e3);
                        }
                    }
                } catch (e) {
                    $.writeln("  Error inserting marker for table " + unplacedTables[ui].index + ": " + e);
                }
            }
            
            // Recompose after inserting markers
            story.recompose();
            $.writeln("Processed " + unplacedTables.length + " unplaced table(s) - markers inserted where possible");
        }
    }
    
    // Warn if we hit the safety limit
    if (pageCount >= maxPages) {
        alert("⚠️ Warning: Reached safety limit of " + maxPages + " pages.\n" +
              "Document may contain problematic tables or formatting.\n" +
              "Pages created: " + (doc.pages.length - initialPageCount));
        $.writeln("⚠️ Safety limit reached. Initial pages: " + initialPageCount + ", Final pages: " + doc.pages.length);
    } else if (story.overflows) {
        alert("⚠️ Warning: Text still overflows after creating " + pageCount + " pages.\n" +
              "This may indicate a table or object that cannot flow properly.\n" +
              "Some content may not be visible.");
        $.writeln("⚠️ Overflow persists after " + pageCount + " iterations");
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

    // ---- 8) Apply table style to native Word tables and fix width issues ----
    $.writeln("Processing " + story.tables.length + " table(s) for styling and width fixes...");
    if (story.tables.length) {
        for (var t=0; t<story.tables.length; t++) {
            try {
                var tbl = story.tables[t];
                
                // Apply table style if available
                if (ps.table) {
                    tbl.appliedTableStyle = ps.table;
                }
                
                // Fix tables that are too wide for the frame
                // This prevents tables from blocking text flow
                try {
                    var parentFrame = tbl.parentTextFrames[0];
                    if (parentFrame) {
                        var frameWidth = parentFrame.geometricBounds[3] - parentFrame.geometricBounds[1];
                        var tableWidth = tbl.width;
                        
                        // If table is wider than frame, resize it to fit
                        if (tableWidth > frameWidth) {
                            tbl.width = frameWidth - 20; // Leave small margin
                            $.writeln("Resized table " + t + " from " + tableWidth + " to " + tbl.width);
                        }
                    }
                } catch (e) {
                    $.writeln("Could not resize table " + t + ": " + e);
                }
            } catch(e) {
                $.writeln("Error processing table " + t + ": " + e);
            }
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

    $.writeln("Converting tagged table blocks...");
    convertTaggedTables(story, "Table Style");

    // ---- 10) Hyperlinks -------
    $.writeln("Processing hyperlinks...");
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
    $.writeln("=== Import Complete ===");
    $.writeln("Final document: " + doc.pages.length + " pages, " + story.paragraphs.length + " paragraphs");
    alert("✅ Import complete.\nHeadings, Bullets, Bold, Tables, Case Studies, and Hyperlinks applied.");

} catch(err) {
    alert("❌ " + err.message);
    $.writeln("❌ " + err.message);
} finally {
    app.scriptPreferences.userInteractionLevel = originalInteraction;
}
