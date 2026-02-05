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


