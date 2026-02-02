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
                    // If table is wider than frame, resize it to fit
                    if (tableWidth > frameWidth) {
                        tbl.width = frameWidth - 20; // Leave small margin
                        $.writeln("Pre-threading: Resized table " + t + " from " + tableWidth + " to " + tbl.width);
                    }
                }
            } catch (e) {
                $.writeln("Could not check/resize table " + t + " before threading: " + e);
            }
        }
        story.recompose(); // Update overflow state after table resizing
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
        
        // Check if overflow state changed - if it's been stable for 2 iterations, break
        if (story.overflows === lastOverflowState) {
            stableIterations++;
            if (stableIterations >= 2) {
                $.writeln("Warning: Overflow state not changing after threading. Breaking to prevent infinite loop.");
                break;
            }
        } else {
            stableIterations = 0;
            lastOverflowState = story.overflows;
        }
    }
    
    // Final recompose
    story.recompose();
    
    $.writeln("Page threading complete. Iterations: " + pageCount + ", Final pages: " + doc.pages.length + ", Final overflow: " + (story.overflows ? "YES" : "NO"));
    
    // Warn if we hit the safety limit
    if (pageCount >= maxPages) {
        alert("⚠️ Warning: Reached safety limit of " + maxPages + " pages.\n" +
              "Document may contain problematic tables or formatting.\n" +
              "Pages created: " + (doc.pages.length - initialPageCount));
        $.writeln("⚠️ Safety limit reached. Initial pages: " + initialPageCount + ", Final pages: " + doc.pages.length);
    } else if (story.overflows) {
        alert("⚠️ Warning: Text still overflows after creating " + pageCount + " pages.\n" +
              "This may indicate a table or object that cannot flow properly.");
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
