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

    // ---- 1b) Set import preferences to ensure tables are imported correctly ----
    // This is critical for fixed-width tables from Word
    var importPrefsSet = false;
    try {
        // InDesign uses wordImportPreferences for Word files
        var importPrefs = app.wordImportPreferences;
        if (importPrefs) {
            try {
                importPrefs.convertTablesTo = TableFormat.INDD_TABLE;
                $.writeln("Set import preference: convertTablesTo = INDD_TABLE");
                importPrefsSet = true;
            } catch (e) {
                var errorMsg = "Could not set table import preference: " + e;
                $.writeln(errorMsg);
                // Don't alert for this - it's not critical if it fails
            }
            try {
                if (importPrefs.hasOwnProperty("preserveLocalOverrides")) {
                    importPrefs.preserveLocalOverrides = false;
                }
            } catch (e) {
                // Ignore if not available
            }
        } else {
            $.writeln("⚠️ wordImportPreferences not available - may not be a Word file");
        }
    } catch (e) {
        // Silently fail - import preferences are optional
        $.writeln("Could not set import preferences: " + e + " (this is OK, continuing...)");
    }
    
    if (!importPrefsSet) {
        $.writeln("⚠️ WARNING: Table import preference may not be set correctly");
    }

    // ---- 2) Place file into first frame ----
    function makeFrame(pg) {
        var m = pg.marginPreferences, b = pg.bounds;
        return pg.textFrames.add({
            geometricBounds: [b[0]+m.top, b[1]+m.left, b[2]-m.bottom, b[3]-m.right]
        });
    }
    
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
                var story = table.parentStory;
                if (story && story.textContainers.length > 0) {
                    // Return the first text container (usually the main frame)
                    // This is a fallback - not perfect but better than null
                    for (var i = 0; i < story.textContainers.length; i++) {
                        try {
                            var container = story.textContainers[i];
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

    var page = doc.pages[0];
    var frame = page.textFrames.length ? page.textFrames[0] : makeFrame(page);
    frame.place(file);

    var story = frame.parentStory;
    if (!story) throw new Error("No story found after placing file.");
    $.writeln("File placed. Story length: " + story.characters.length + " characters, " + story.paragraphs.length + " paragraphs");
    
    // DEEP DIAGNOSTIC: Understand the actual import state
    $.writeln("=== DEEP DIAGNOSTIC: Post-Import State ===");
    $.writeln("Initial overflow: " + (story.overflows ? "YES" : "NO"));
    $.writeln("Text containers: " + story.textContainers.length);
    $.writeln("Tables found: " + story.tables.length);
    
    // Check what's actually in the story
    try {
        var firstChars = story.characters.length > 0 ? story.characters.item(0).contents.substring(0, 50) : "empty";
        var lastChars = story.characters.length > 0 ? story.characters.item(-1).contents.substring(0, 50) : "empty";
        $.writeln("First 50 chars: " + firstChars);
        $.writeln("Last 50 chars: " + lastChars);
    } catch (e) {
        $.writeln("Could not get character samples: " + e);
    }
    
    // Check frame state
    try {
        $.writeln("Initial frame overflow: " + (frame.overflows ? "YES" : "NO"));
        var frameBounds = frame.geometricBounds;
        $.writeln("Frame size: " + (frameBounds[3] - frameBounds[1]).toFixed(2) + "pt × " + (frameBounds[2] - frameBounds[0]).toFixed(2) + "pt");
    } catch (e) {
        $.writeln("Could not get frame info: " + e);
    }
    $.writeln("=== End Deep Diagnostic ===");

    // ---- 2b) CRITICAL — INSERT CASE STUDIES NOW (BEFORE ANY REFLOW) ----
    $.writeln("Checking for case studies...");
    insertCaseStudies(story);

    // ---- 2c) Check for tables that cross pages and skip them ----
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
        }
    }
    
    // ---- 2d) Process remaining tables (those that don't cross pages) ----
    // Fixed-width tables from Word can cause overflow because they're set to specific widths
    // that may not fit the InDesign frame. We need to detect and resize them aggressively.
    // CRITICAL: Also enable page breaks so tables can split across pages.
    if (story.tables.length) {
        for (var t=0; t<story.tables.length; t++) {
            try {
                var tbl = story.tables[t];
                
                // DEEP DIAGNOSTIC: Understand this specific table BEFORE processing
                $.writeln("=== DEEP DIAGNOSTIC: Table " + t + " (BEFORE PROCESSING) ===");
                var tableDiagnosticMsg = "";
                try {
                    var tableRows = tbl.rows.length;
                    var tableCols = tbl.columns.length;
                    var tableW = tbl.width.toFixed(2);
                    var tableH = tbl.height.toFixed(2);
                    
                    $.writeln("Table rows: " + tableRows + ", cols: " + tableCols);
                    $.writeln("Table width: " + tableW + "pt, height: " + tableH + "pt");
                    
                    tableDiagnosticMsg = "📊 Table " + t + " Analysis:\n\n" +
                                        "Dimensions: " + tableRows + " rows × " + tableCols + " cols\n" +
                                        "Size: " + tableW + "pt × " + tableH + "pt\n";
                    
                    // Check table's position in story - find where it starts
                    try {
                        var firstCell = tbl.cells[0];
                        if (firstCell) {
                            var cellText = firstCell.texts[0];
                            if (cellText && cellText.parentStory) {
                                var charIndex = cellText.characters[0].index;
                                $.writeln("Table starts at character index: " + charIndex + " (out of " + story.characters.length + " total)");
                                
                                // Check what's before and after the table
                                if (charIndex > 0) {
                                    var beforeText = story.characters.itemByRange(0, Math.min(charIndex - 1, 100)).contents;
                                    $.writeln("Text before table (first 100 chars): " + beforeText.substring(0, 100));
                                }
                                if (charIndex < story.characters.length - 1) {
                                    var afterStart = charIndex + 100;
                                    var afterEnd = Math.min(afterStart + 100, story.characters.length - 1);
                                    if (afterStart < story.characters.length) {
                                        var afterText = story.characters.itemByRange(afterStart, afterEnd).contents;
                                        $.writeln("Text after table (100 chars): " + afterText.substring(0, 100));
                                    }
                                }
                            }
                        }
                    } catch (e) {
                        $.writeln("Could not get table position in story: " + e);
                    }
                    
                    // Check parent frame
                    var parentFrame = getTableParentFrame(tbl);
                    if (parentFrame) {
                        $.writeln("Table parent frame: " + (parentFrame.overflows ? "OVERFLOWS" : "OK"));
                        var frameW = parentFrame.geometricBounds[3] - parentFrame.geometricBounds[1];
                        var frameH = parentFrame.geometricBounds[2] - parentFrame.geometricBounds[0];
                        var frameWStr = frameW.toFixed(2);
                        var frameHStr = frameH.toFixed(2);
                        var widthPct = ((tbl.width / frameW) * 100).toFixed(1);
                        var heightPct = ((tbl.height / frameH) * 100).toFixed(1);
                        
                        $.writeln("Parent frame size: " + frameWStr + "pt × " + frameHStr + "pt");
                        $.writeln("Table vs frame: " + widthPct + "% width, " + heightPct + "% height");
                        
                        tableDiagnosticMsg += "Frame size: " + frameWStr + "pt × " + frameHStr + "pt\n" +
                                             "Table vs frame: " + widthPct + "% width, " + heightPct + "% height\n";
                        
                        // Check if table is taller than frame
                        if (tbl.height > frameH) {
                            var criticalMsg = "🚨 CRITICAL: Table is TALLER than frame!\n\n" +
                                            "Table height: " + tableH + "pt\n" +
                                            "Frame height: " + frameHStr + "pt\n" +
                                            "Difference: " + (tbl.height - frameH).toFixed(2) + "pt\n" +
                                            "Table is " + heightPct + "% of frame height\n\n" +
                                            "This table CANNOT fit in one frame!\n" +
                                            "Page breaks MUST be enabled for it to flow!";
                            $.writeln("⚠️ CRITICAL: Table is " + heightPct + "% taller than frame!");
                            tableDiagnosticMsg += "\n⚠️ CRITICAL: Table is " + heightPct + "% taller than frame!";
                        }
                        
                        // Check page break status - try multiple methods
                        var pageBreakStatus = "unknown";
                        var pageBreaksEnabled = false;
                        try {
                        // Method 1: Try allowPageBreak property
                        if (tbl.hasOwnProperty("allowPageBreak")) {
                            pageBreaksEnabled = tbl.allowPageBreak;
                            pageBreakStatus = pageBreaksEnabled ? "enabled" : "disabled";
                        } else {
                            // Method 2: Try via table style
                            try {
                                var tableStyle = tbl.appliedTableStyle;
                                if (tableStyle && tableStyle.hasOwnProperty("allowPageBreak")) {
                                    pageBreaksEnabled = tableStyle.allowPageBreak;
                                    pageBreakStatus = pageBreaksEnabled ? "enabled (via style)" : "disabled (via style)";
                                } else {
                                    pageBreakStatus = "property not available";
                                }
                            } catch (e) {
                                pageBreakStatus = "cannot check (property doesn't exist)";
                            }
                        }
                        $.writeln("Page breaks: " + pageBreakStatus);
                        tableDiagnosticMsg += "\nPage breaks: " + pageBreakStatus;
                        
                        // Log if we can't enable page breaks and table is tall
                        if (!pageBreaksEnabled && parentFrame) {
                            var frameH = parentFrame.geometricBounds[2] - parentFrame.geometricBounds[0];
                            if (tbl.height > frameH * 0.8) {
                                $.writeln("⚠️ WARNING: Page breaks not enabled on tall table");
                                tableDiagnosticMsg += "\n⚠️ WARNING: Page breaks not enabled!";
                            }
                        }
                        } catch (e) {
                            $.writeln("Could not check page breaks: " + e);
                            pageBreakStatus = "error: " + e;
                        }
                    } else {
                        $.writeln("⚠️ WARNING: Could not find table's parent frame!");
                    }
                } catch (e) {
                    $.writeln("Error in table diagnostic: " + e);
                }
                $.writeln("=== End Table " + t + " Diagnostic ===");
                
                // CRITICAL FIX: Enable page breaks for tables so they can split across pages
                // Without this, tables that are taller than one page will cause import to fail
                var pageBreaksSet = false;
                try {
                    // Method 1: Try allowPageBreak property directly
                    if (tbl.hasOwnProperty("allowPageBreak")) {
                        try {
                            var oldValue = tbl.allowPageBreak;
                            tbl.allowPageBreak = true;
                            var newValue = tbl.allowPageBreak;
                            
                            if (newValue === true) {
                                $.writeln("Table " + t + ": ✓ Page breaks enabled (was: " + oldValue + ", now: " + newValue + ")");
                                pageBreaksSet = true;
                            } else {
                                $.writeln("⚠️ Table " + t + ": WARNING - allowPageBreak set to true but value is " + newValue);
                            }
                        } catch (e) {
                            $.writeln("Error setting allowPageBreak: " + e);
                        }
                    } else {
                        $.writeln("⚠️ Table " + t + ": allowPageBreak property does not exist on table object");
                        
                        // Method 2: Try setting on rows instead
                        $.writeln("Attempting to enable page breaks via rows...");
                        try {
                            var rowsWithPageBreaks = 0;
                            for (var r = 0; r < tbl.rows.length; r++) {
                                try {
                                    if (tbl.rows[r].hasOwnProperty("allowPageBreak")) {
                                        tbl.rows[r].allowPageBreak = true;
                                        if (tbl.rows[r].allowPageBreak === true) {
                                            rowsWithPageBreaks++;
                                        }
                                    }
                                } catch (e) {}
                            }
                            if (rowsWithPageBreaks > 0) {
                                $.writeln("  ✓ Enabled page breaks on " + rowsWithPageBreaks + " of " + tbl.rows.length + " rows");
                                pageBreaksSet = true;
                            } else {
                                $.writeln("  ⚠️ Could not enable page breaks on rows either");
                            }
                        } catch (e) {
                            $.writeln("  Error enabling page breaks on rows: " + e);
                        }
                        
                        // Method 3: Try via table style
                        try {
                            var tableStyle = tbl.appliedTableStyle;
                            if (tableStyle && tableStyle.hasOwnProperty("allowPageBreak")) {
                                tableStyle.allowPageBreak = true;
                                $.writeln("  ✓ Set page breaks via table style");
                                pageBreaksSet = true;
                            }
                        } catch (e) {
                            $.writeln("  Could not set via table style: " + e);
                        }
                    }
                    
                    if (!pageBreaksSet && parentFrame) {
                        var frameH = parentFrame.geometricBounds[2] - parentFrame.geometricBounds[0];
                        if (tbl.height > frameH) {
                            // Table is taller than frame and we can't enable page breaks
                            // Try alternative: Force recompose and check if InDesign handles it automatically
                            $.writeln("⚠️ Cannot enable page breaks on table - attempting recompose");
                            
                            // Force recompose - sometimes InDesign handles tall tables automatically
                            try {
                                story.recompose();
                                $.writeln("Forced recompose after page break attempt");
                            } catch (e) {
                                $.writeln("Error forcing recompose: " + e);
                            }
                        }
                    }
                } catch (e) {
                    $.writeln("Error in page break setup: " + e);
                }
                
                var parentFrame = getTableParentFrame(tbl);
                if (parentFrame) {
                    var frameWidth = parentFrame.geometricBounds[3] - parentFrame.geometricBounds[1];
                    var tableWidth = tbl.width;
                    var widthDifference = Math.abs(tableWidth - frameWidth);
                    
                    // Detect fixed-width tables: width is very close to frame width (within 5 points)
                    // or table is >= 95% of frame width - these are likely fixed-width from Word
                    var isFixedWidthTable = (widthDifference < 5) || (tableWidth >= frameWidth * 0.95);
                    
                    if (isFixedWidthTable) {
                        // Fixed-width table detected - resize aggressively to 90% of frame
                        // This gives it room to flow and prevents blocking
                        var newWidth = frameWidth * 0.90;
                        tbl.width = newWidth;
                        $.writeln("Pre-threading: FIXED-WIDTH table " + t + " detected (width: " + tableWidth.toFixed(2) + ", frame: " + frameWidth.toFixed(2) + ", diff: " + widthDifference.toFixed(2) + ")");
                        $.writeln("  Resized to " + newWidth.toFixed(2) + " (90% of frame)");
                        
                        // Try to redistribute column widths to make it more flexible
                        // This helps convert fixed-width layout to auto-fit
                        try {
                            var numCols = tbl.columns.length;
                            if (numCols > 0) {
                                var colWidth = newWidth / numCols;
                                for (var col = 0; col < numCols; col++) {
                                    try {
                                        tbl.columns[col].width = colWidth;
                                    } catch (e) {
                                        // Ignore individual column errors
                                    }
                                }
                                $.writeln("  Redistributed " + numCols + " columns to equal width (" + colWidth.toFixed(2) + " each)");
                            }
                        } catch (e) {
                            $.writeln("  Could not redistribute columns: " + e);
                        }
                    } else if (tableWidth > frameWidth * 0.85) {
                        // Regular oversized table - resize to fit with margin
                        var newWidth = frameWidth - 30; // Leave larger margin for safety
                        tbl.width = newWidth;
                        $.writeln("Pre-threading: Resized oversized table " + t + " from " + tableWidth.toFixed(2) + " to " + newWidth.toFixed(2) + " (frame: " + frameWidth.toFixed(2) + ")");
                    } else {
                        $.writeln("Pre-threading: Table " + t + " width OK (" + tableWidth.toFixed(2) + " <= " + (frameWidth * 0.85).toFixed(2) + ")");
                    }
                } else {
                    $.writeln("Pre-threading: Table " + t + " has no parent frame");
                }
            } catch (e) {
                $.writeln("Could not check/resize table " + t + " before threading: " + e);
            }
        }
        story.recompose(); // Update overflow state after table resizing
        var overflowAfterTables = story.overflows;
        $.writeln("After table resizing: Overflow = " + (overflowAfterTables ? "YES" : "NO"));
        
        // DEEP DIAGNOSTIC: What's in the overflow?
        if (overflowAfterTables) {
            $.writeln("=== DEEP DIAGNOSTIC: Overflow Analysis ===");
            try {
                // Find the overflowing frame
                var overflowingFrame = null;
                for (var of = 0; of < story.textContainers.length; of++) {
                    try {
                        var container = story.textContainers[of];
                        if (container.overflows) {
                            overflowingFrame = container;
                            $.writeln("Found overflowing frame at container index: " + of);
                            break;
                        }
                    } catch (e) {}
                }
                
                if (overflowingFrame) {
                    try {
                        // Get the last visible character in this frame
                        var frameStory = overflowingFrame.parentStory;
                        var frameChars = frameStory.characters;
                        
                        // Try to find where the overflow starts
                        // This is tricky - we need to find the last character that fits
                        $.writeln("Total characters in story: " + frameChars.length);
                        $.writeln("Overflowing frame bounds: " + overflowingFrame.geometricBounds);
                        
                        // Check if table is in the overflowing frame
                        for (var ot = 0; ot < story.tables.length; ot++) {
                            try {
                                var otbl = story.tables[ot];
                                var otblFrame = getTableParentFrame(otbl);
                                if (otblFrame === overflowingFrame) {
                                    $.writeln("⚠️ Table " + ot + " is in the overflowing frame!");
                                    
                                    // Check table dimensions vs frame
                                    var oframeW = otblFrame.geometricBounds[3] - otblFrame.geometricBounds[1];
                                    var oframeH = otblFrame.geometricBounds[2] - otblFrame.geometricBounds[0];
                                    $.writeln("  Table: " + otbl.width.toFixed(2) + "pt × " + otbl.height.toFixed(2) + "pt");
                                    $.writeln("  Frame: " + oframeW.toFixed(2) + "pt × " + oframeH.toFixed(2) + "pt");
                                    
                                    if (otbl.height > oframeH) {
                                        var heightPct = ((otbl.height / oframeH) * 100).toFixed(1);
                                        $.writeln("  ⚠️ CRITICAL: Table is " + heightPct + "% of frame height");
                                        $.writeln("  ⚠️ Table MUST split across pages - checking page breaks...");
                                        
                                        var pbStatus = "unknown";
                                        var pbEnabled = false;
                                        var pbError = null;
                                        try {
                                            if (otbl.hasOwnProperty("allowPageBreak")) {
                                                pbEnabled = otbl.allowPageBreak;
                                                pbStatus = pbEnabled ? "ENABLED" : "DISABLED";
                                            } else {
                                                pbStatus = "PROPERTY NOT AVAILABLE";
                                                pbError = "allowPageBreak property doesn't exist on table object";
                                            }
                                            $.writeln("  Page breaks: " + pbStatus);
                                            
                                            if (pbError) {
                                                // Property doesn't exist - this is the root cause
                                                $.writeln("  ❌ ROOT CAUSE: Page breaks property NOT AVAILABLE");
                                            } else if (!pbEnabled) {
                                                $.writeln("  ❌ THIS IS THE PROBLEM: Page breaks are disabled!");
                                                
                                                // Try to enable it NOW
                                                try {
                                                    otbl.allowPageBreak = true;
                                                    for (var pr = 0; pr < otbl.rows.length; pr++) {
                                                        try {
                                                            otbl.rows[pr].allowPageBreak = true;
                                                        } catch (e) {}
                                                    }
                                                    story.recompose();
                                                    $.writeln("  ✓ Page breaks enabled in overflow fix");
                                                } catch (e) {
                                                    $.writeln("  ❌ Failed to enable page breaks: " + e);
                                                }
                                            } else {
                                                $.writeln("  ⚠️ Page breaks enabled but table still not flowing");
                                            }
                                        } catch (e) {
                                            $.writeln("  Could not check page breaks: " + e);
                                        }
                                    }
                                }
                            } catch (e) {
                                $.writeln("Error checking table " + ot + " in overflow: " + e);
                            }
                        }
                    } catch (e) {
                        $.writeln("Error analyzing overflow: " + e);
                    }
                } else {
                    $.writeln("⚠️ Could not find overflowing frame!");
                }
            } catch (e) {
                $.writeln("Error in overflow diagnostic: " + e);
            }
            $.writeln("=== End Overflow Analysis ===");
        }
        
        // Alert if overflow persists after table processing
        if (overflowAfterTables) {
            // Get detailed table info for diagnostics
            var tableDetails = "";
            try {
                for (var diag = 0; diag < story.tables.length; diag++) {
                    try {
                        var diagTbl = story.tables[diag];
                        var diagFrame = getTableParentFrame(diagTbl);
                        if (diagFrame) {
                            var frameW = diagFrame.geometricBounds[3] - diagFrame.geometricBounds[1];
                            var frameH = diagFrame.geometricBounds[2] - diagFrame.geometricBounds[0];
                            var tblW = diagTbl.width;
                            var tblH = diagTbl.height;
                            var pageBreakStatus = "unknown";
                            try {
                                pageBreakStatus = diagTbl.allowPageBreak ? "enabled" : "disabled";
                            } catch (e) {}
                            
                            tableDetails += "\n\nTable " + diag + ":";
                            tableDetails += "\n  Size: " + tblW.toFixed(1) + "pt × " + tblH.toFixed(1) + "pt";
                            tableDetails += "\n  Frame: " + frameW.toFixed(1) + "pt × " + frameH.toFixed(1) + "pt";
                            tableDetails += "\n  Page breaks: " + pageBreakStatus;
                            tableDetails += "\n  Overflow: " + (diagFrame.overflows ? "YES" : "NO");
                            
                            // If table is too wide or too tall, note it
                            if (tblW > frameW * 0.95) {
                                tableDetails += "\n  ⚠️ Table is " + ((tblW / frameW) * 100).toFixed(1) + "% of frame width";
                            }
                            if (tblH > frameH * 0.9) {
                                tableDetails += "\n  ⚠️ Table is " + ((tblH / frameH) * 100).toFixed(1) + "% of frame height";
                            }
                        }
                    } catch (e) {
                        tableDetails += "\n\nTable " + diag + ": Error getting details - " + e;
                    }
                }
            } catch (e) {}
            
            $.writeln("⚠️ Overflow detected after processing " + tableCount + " table(s)");
        }
        
        // Second pass: if overflow persists, check for tables that still might be blocking
        if (story.overflows) {
            $.writeln("Overflow persists after first table resize pass. Checking for remaining blocking tables...");
            for (var t2=0; t2<story.tables.length; t2++) {
                try {
                    var tbl2 = story.tables[t2];
                    var parentFrame2 = getTableParentFrame(tbl2);
                    if (parentFrame2 && parentFrame2.overflows) {
                        var frameWidth2 = parentFrame2.geometricBounds[3] - parentFrame2.geometricBounds[1];
                        var tableWidth2 = tbl2.width;
                        // If table is still > 85% of frame, resize more aggressively
                        if (tableWidth2 > frameWidth2 * 0.85) {
                            var newWidth2 = frameWidth2 * 0.85; // More aggressive resize
                            tbl2.width = newWidth2;
                            $.writeln("Second pass: Aggressively resized table " + t2 + " to " + newWidth2.toFixed(2) + " (85% of frame)");
                        }
                    }
                } catch (e) {
                    $.writeln("Could not check/resize table " + t2 + " in second pass: " + e);
                }
            }
            story.recompose();
            $.writeln("After second table resize pass: Overflow = " + (story.overflows ? "YES" : "NO"));
        }
    }

    // ---- Continue threading pages with safety limits ----
    var maxPages = 50; // Safety limit to prevent infinite loops (reduced from 100)
    var pageCount = 0;
    var initialPageCount = doc.pages.length;
    var lastOverflowState = story.overflows;
    var stableIterations = 0;
    var lastVisibleCharCount = 0; // Track visible character count to detect if content is flowing
    
    // Get initial visible character count
    try {
        var lastFrame = story.textContainers[story.textContainers.length - 1];
        if (lastFrame && !lastFrame.overflows) {
            lastVisibleCharCount = lastFrame.parentStory.characters.length;
        }
    } catch (e) {
        // Ignore if we can't get initial count
    }
    
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
        
        // Check if content is actually flowing by tracking the last visible character index
        var currentLastVisibleIndex = -1;
        var contentFlowing = false;
        try {
            // Find the last non-overflowing frame and get its last character index
            for (var vc = story.textContainers.length - 1; vc >= 0; vc--) {
                try {
                    var vf = story.textContainers[vc];
                    if (!vf.overflows && vf.parentStory) {
                        // Get the last character in this frame
                        var frameChars = vf.parentStory.characters;
                        if (frameChars.length > 0) {
                            try {
                                var lastChar = frameChars[-1];
                                currentLastVisibleIndex = lastChar.index;
                                // Check if we've made progress
                                contentFlowing = (currentLastVisibleIndex > lastVisibleCharCount);
                                if (contentFlowing) {
                                    lastVisibleCharCount = currentLastVisibleIndex;
                                }
                                break;
                            } catch (e) {
                                // Try alternative method
                                currentLastVisibleIndex = frameChars.length - 1;
                                contentFlowing = (currentLastVisibleIndex > lastVisibleCharCount);
                                if (contentFlowing) {
                                    lastVisibleCharCount = currentLastVisibleIndex;
                                }
                                break;
                            }
                        }
                    }
                } catch (e) {}
            }
        } catch (e) {
            // If we can't check, assume no progress
            contentFlowing = false;
        }
        
        // Log progress every 5 pages or on first few iterations
        if (pageCount <= 3 || pageCount % 5 === 0) {
            $.writeln("  Iteration " + pageCount + ": Pages=" + doc.pages.length + ", Overflow=" + (story.overflows ? "YES" : "NO") + ", Containers=" + story.textContainers.length + ", ContentFlowing=" + (contentFlowing ? "YES" : "NO"));
        }
        
        // Early break if we've created many pages but overflow persists
        // This catches cases where tables are blocking and we're just creating empty pages
        if (pageCount >= 15 && story.overflows) {
            // If we've created 15+ pages and still have overflow, something is seriously wrong
            var earlyBreakMsg = "⚠️ EARLY BREAK: Created " + pageCount + " pages but overflow persists.\n" +
                               "Likely a blocking table or object that cannot flow.\n" +
                               "Breaking to prevent excessive page creation.";
            $.writeln("⚠️ EARLY BREAK: " + earlyBreakMsg);
            break;
        }
        
        // Check if overflow state changed - if it's been stable, investigate
        if (story.overflows === lastOverflowState) {
            stableIterations++;
            
            // If overflow hasn't changed AND content isn't flowing, we're stuck
            if (stableIterations >= 2 && !contentFlowing) {
                $.writeln("⚠️ STUCK DETECTED: Overflow unchanged for " + stableIterations + " iterations AND no content flowing. Breaking to prevent infinite loop.");
                break;
            }
            
            // If overflow hasn't changed for 2 iterations, check if it's a table blocking flow
            if (stableIterations >= 2) {
                $.writeln("Warning: Overflow state not changing. Checking for blocking tables...");
                
                // Check if there are tables that might be blocking flow
                var blockingTable = false;
                try {
                    for (var ti = 0; ti < story.tables.length; ti++) {
                        var t = story.tables[ti];
                        var tf = getTableParentFrame(t);
                        if (tf && tf.overflows) {
                            // Table is in an overflowing frame - try to resize it more aggressively
                            try {
                                var frameWidth = tf.geometricBounds[3] - tf.geometricBounds[1];
                                var tableWidth = t.width;
                                var widthDifference = Math.abs(tableWidth - frameWidth);
                                
                                // Check if this is a fixed-width table (very close to frame width)
                                var isFixedWidth = (widthDifference < 5) || (tableWidth >= frameWidth * 0.95);
                                
                                if (isFixedWidth) {
                                    // Fixed-width table blocking flow - resize very aggressively
                                    var newWidth = frameWidth * 0.80; // 80% of frame
                                    t.width = newWidth;
                                    $.writeln("  FIXED-WIDTH blocking table " + ti + " detected. Resized from " + tableWidth.toFixed(2) + " to " + newWidth.toFixed(2));
                                    
                                    // Try to redistribute columns
                                    try {
                                        var numCols = t.columns.length;
                                        if (numCols > 0) {
                                            var colWidth = newWidth / numCols;
                                            for (var col = 0; col < numCols; col++) {
                                                try {
                                                    t.columns[col].width = colWidth;
                                                } catch (e) {}
                                            }
                                            $.writeln("    Redistributed " + numCols + " columns");
                                        }
                                    } catch (e) {}
                                    
                                    story.recompose();
                                    blockingTable = true;
                                } else if (tableWidth > frameWidth * 0.9) {
                                    // Regular oversized table - resize aggressively
                                    t.width = frameWidth * 0.85; // More aggressive resize
                                    $.writeln("  Aggressively resized blocking table " + ti + " from " + tableWidth.toFixed(2) + " to " + (frameWidth * 0.85).toFixed(2));
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
                
                // If still no change after table fix attempt, check if content is flowing
                // If content is flowing, allow a few more iterations for text after table
                if (stableIterations >= 3) {
                    if (!contentFlowing) {
                        $.writeln("⚠️ Breaking: Overflow persists for " + stableIterations + " iterations AND no content flowing. Table likely blocking.");
                        break;
                    } else {
                        $.writeln("Warning: Overflow persists after " + stableIterations + " stable iterations, but content is flowing. Continuing...");
                    }
                }
                
                // Hard limit: break after 5 stable iterations regardless
                if (stableIterations >= 5) {
                    $.writeln("⚠️ Breaking after " + stableIterations + " stable iterations (hard limit). Content may still be in overflow.");
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
    
    // FINAL AGGRESSIVE TABLE FIX: If overflow still persists, try one last desperate attempt
    if (story.overflows && story.tables.length > 0) {
        $.writeln("⚠️ FINAL ATTEMPT: Overflow persists. Trying aggressive table fixes...");
        var finalFixAttempted = false;
        
        for (var finalT = 0; finalT < story.tables.length; finalT++) {
            try {
                var finalTbl = story.tables[finalT];
                var finalFrame = getTableParentFrame(finalTbl);
                
                if (finalFrame && finalFrame.overflows) {
                    var finalFrameW = finalFrame.geometricBounds[3] - finalFrame.geometricBounds[1];
                    var finalTblW = finalTbl.width;
                    
                    // Ultra-aggressive resize: force to 75% of frame width
                    if (finalTblW > finalFrameW * 0.75) {
                        var ultraWidth = finalFrameW * 0.75;
                        finalTbl.width = ultraWidth;
                        $.writeln("  FINAL FIX: Table " + finalT + " resized to " + ultraWidth.toFixed(2) + "pt (75% of frame)");
                        finalFixAttempted = true;
                        
                        // Try to redistribute columns
                        try {
                            var numCols = finalTbl.columns.length;
                            if (numCols > 0) {
                                var colWidth = ultraWidth / numCols;
                                for (var col = 0; col < numCols; col++) {
                                    try {
                                        finalTbl.columns[col].width = colWidth;
                                    } catch (e) {}
                                }
                            }
                        } catch (e) {}
                        
                        // Force page breaks one more time
                        try {
                            finalTbl.allowPageBreak = true;
                            for (var r = 0; r < finalTbl.rows.length; r++) {
                                try {
                                    finalTbl.rows[r].allowPageBreak = true;
                                } catch (e) {}
                            }
                        } catch (e) {}
                    }
                }
            } catch (e) {
                $.writeln("  Error in final table fix for table " + finalT + ": " + e);
            }
        }
        
        if (finalFixAttempted) {
            story.recompose();
            $.writeln("After final aggressive table fix: Overflow = " + (story.overflows ? "YES" : "NO"));
        }
    }
    
    // Check for tables that couldn't be placed and add marker text
    if (story.overflows && story.tables.length > 0) {
        $.writeln("Checking for unplaced tables and adding markers...");
        var unplacedTables = [];
        
        for (var ut = 0; ut < story.tables.length; ut++) {
            try {
                var tbl = story.tables[ut];
                var parentFrame = getTableParentFrame(tbl);
                
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
                            var parentFrame = getTableParentFrame(tbl);
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
    
    // Diagnostic: Check table page break status after threading
    $.writeln("=== Table Page Break Diagnostic ===");
    if (story.tables.length > 0) {
        for (var diag = 0; diag < story.tables.length; diag++) {
            try {
                var diagTbl = story.tables[diag];
                var hasPageBreak = false;
                var pageBreakValue = "unknown";
                
                try {
                    if (diagTbl.hasOwnProperty("allowPageBreak")) {
                        pageBreakValue = diagTbl.allowPageBreak;
                        hasPageBreak = (diagTbl.allowPageBreak === true);
                    }
                } catch (e) {
                    pageBreakValue = "error: " + e;
                }
                
                // Check if table spans multiple pages
                var spansPages = false;
                var tableFrames = [];
                try {
                    var tempFrame = getTableParentFrame(diagTbl);
                    tableFrames = tempFrame ? [tempFrame] : [];
                    if (tableFrames.length > 1) {
                        spansPages = true;
                    } else if (tableFrames.length === 1) {
                        // Check if table is taller than its frame
                        try {
                            var frameHeight = tableFrames[0].geometricBounds[2] - tableFrames[0].geometricBounds[0];
                            var tableHeight = diagTbl.height;
                            if (tableHeight > frameHeight * 0.9) {
                                spansPages = true; // Likely spans pages
                            }
                        } catch (e) {}
                    }
                } catch (e) {}
                
                $.writeln("Table " + diag + ": allowPageBreak=" + pageBreakValue + ", spansPages=" + spansPages + ", frames=" + tableFrames.length);
                
                if (spansPages && !hasPageBreak) {
                    $.writeln("  ⚠️ WARNING: Table " + diag + " spans pages but allowPageBreak is NOT enabled!");
                } else if (spansPages && hasPageBreak) {
                    $.writeln("  ✓ Table " + diag + " spans pages and has page breaks enabled");
                }
            } catch (e) {
                $.writeln("  ❌ Error checking table " + diag + ": " + e);
            }
        }
    } else {
        $.writeln("No tables found in document");
    }
    $.writeln("=== End Table Diagnostic ===");
    
    // Warn if we hit the safety limit
    if (pageCount >= maxPages) {
        $.writeln("⚠️ Safety limit reached. Initial pages: " + initialPageCount + ", Final pages: " + doc.pages.length);
    } else if (story.overflows) {
        // Enhanced diagnostic message
        var tableInfo = "";
        try {
            var remainingTables = story.tables.length;
            if (remainingTables > 0) {
                tableInfo = "\n\nTables in document: " + remainingTables;
                // Check if any tables are in overflowing frames
                var blockingTables = 0;
                for (var diag = 0; diag < remainingTables; diag++) {
                    try {
                        var diagTbl = story.tables[diag];
                        var diagFrame = getTableParentFrame(diagTbl);
                        if (diagFrame && diagFrame.overflows) {
                            blockingTables++;
                            tableInfo += "\n  - Table " + diag + " is in an overflowing frame";
                        }
                    } catch (e) {}
                }
                if (blockingTables > 0) {
                    tableInfo += "\n  ⚠️ " + blockingTables + " table(s) may be blocking flow";
                }
            }
        } catch (e) {}
        
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
                
                // Ensure page breaks are enabled (defensive - in case tables were modified)
                try {
                    if (tbl.hasOwnProperty("allowPageBreak")) {
                        var currentValue = tbl.allowPageBreak;
                        tbl.allowPageBreak = true;
                        if (tbl.allowPageBreak !== true) {
                            $.writeln("⚠️ Table " + t + " (styling pass): Page break setting failed (value: " + tbl.allowPageBreak + ")");
                        }
                        // Also ensure rows can break
                        var rowBreakCount = 0;
                        for (var r = 0; r < tbl.rows.length; r++) {
                            try {
                                if (tbl.rows[r].hasOwnProperty("allowPageBreak")) {
                                    tbl.rows[r].allowPageBreak = true;
                                    if (tbl.rows[r].allowPageBreak === true) rowBreakCount++;
                                }
                            } catch (e) {}
                        }
                        if (rowBreakCount < tbl.rows.length) {
                            $.writeln("⚠️ Table " + t + " (styling pass): Only " + rowBreakCount + " of " + tbl.rows.length + " rows have page breaks");
                        }
                    } else {
                        $.writeln("⚠️ Table " + t + " (styling pass): allowPageBreak property missing");
                    }
                } catch (e) {
                    $.writeln("❌ ERROR: Could not ensure page breaks for table " + t + " (styling pass): " + e);
                }
                
                // Apply table style if available
                if (ps.table) {
                    tbl.appliedTableStyle = ps.table;
                }
                
                // Fix tables that are too wide for the frame
                // This prevents tables from blocking text flow
                try {
                    var parentFrame = getTableParentFrame(tbl);
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
            
            // Enable page breaks for tagged tables so they can split across pages
            try {
                if (tbl.hasOwnProperty("allowPageBreak")) {
                    tbl.allowPageBreak = true;
                    var tagTableBreakSet = (tbl.allowPageBreak === true);
                    if (!tagTableBreakSet) {
                        $.writeln("⚠️ Tagged table: Page break setting failed (value: " + tbl.allowPageBreak + ")");
                    }
                    
                    var tagRowBreakCount = 0;
                    for (var r = 0; r < tbl.rows.length; r++) {
                        try {
                            if (tbl.rows[r].hasOwnProperty("allowPageBreak")) {
                                tbl.rows[r].allowPageBreak = true;
                                if (tbl.rows[r].allowPageBreak === true) tagRowBreakCount++;
                            }
                        } catch (e) {
                            $.writeln("  Tagged table row " + r + " error: " + e);
                        }
                    }
                    
                    if (tagTableBreakSet && tagRowBreakCount === tbl.rows.length) {
                        $.writeln("✓ Tagged table: Page breaks enabled for table and all " + tbl.rows.length + " rows");
                    } else {
                        $.writeln("⚠️ Tagged table: Page breaks - table: " + tagTableBreakSet + ", rows: " + tagRowBreakCount + "/" + tbl.rows.length);
                    }
                } else {
                    $.writeln("⚠️ Tagged table: allowPageBreak property missing");
                }
            } catch (e) {
                $.writeln("❌ ERROR: Could not enable page breaks for tagged table: " + e);
                $.writeln("  Error details: " + e.toString());
            }
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
