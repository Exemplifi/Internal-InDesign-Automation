// InDesign Script: Word/RTF Import — Final with Robust Hyperlink Detection

#target "InDesign"

if (app.documents.length === 0) {
    alert("Please open an InDesign template first!");
    exit();
}

var doc = app.activeDocument;
var originalInteraction = app.scriptPreferences.userInteractionLevel;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

try {
    // ---- 1) Select file ----
    var file = File.openDialog("Select Word (.docx) or RTF file", "*.docx;*.rtf", false);
    if (!file || !file.exists) throw new Error("No file selected or not found.");

    // ---- 2) Place file ----
    function makeFrame(pg) {
        var m = pg.marginPreferences, b = pg.bounds;
        return pg.textFrames.add({ geometricBounds: [b[0]+m.top, b[1]+m.left, b[2]-m.bottom, b[3]-m.right] });
    }

    var page = doc.pages[0];
    var frame = page.textFrames.length ? page.textFrames[0] : makeFrame(page);
    frame.place(file);
    var story = frame.parentStory;
    if (!story) throw new Error("No story found after placing file.");

    while (story.overflows) {
        var last = doc.pages[-1];
        var np = doc.pages.add(LocationOptions.AFTER, last);
        var nf = makeFrame(np);
        story.textContainers[story.textContainers.length - 1].nextTextFrame = nf;
    }

    // ---- 3) Get styles ----
    function s(n, t) { try { return t=="p"?doc.paragraphStyles.itemByName(n):doc.characterStyles.itemByName(n); } catch(_) { return null; } }
    var ps = {
        h:[s("Heading 1","p"),s("Heading 2","p"),s("Heading 3","p"),s("Heading 4","p"),s("Heading 5","p"),s("Heading 6","p")],
        body:s("Body","p"), bullets:s("Bullets","p"), sub:s("Sub-Bullets","p"), num:s("Numbered","p"), table:s("Table Style","p")
    };
    var cs = { bold:s("Body - Bold","c"), link:s("Hyperlink Highlight","c") };

    // ---- 4) Collapse ^p^p → ^p ----
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.findWhat = "(?<=\\r)\\r+";
    app.changeGrepPreferences.changeTo = "";
    story.changeGrep();

    // ---- 5) Apply paragraph styles ----
    for (var i=0;i<story.paragraphs.length;i++){
        var p=story.paragraphs[i], n="";
        try{n=p.appliedParagraphStyle.name;}catch(_){}
        if(/Heading\\s*1/i.test(n)&&ps.h[0])p.appliedParagraphStyle=ps.h[0];
        else if(/Heading\\s*2/i.test(n)&&ps.h[1])p.appliedParagraphStyle=ps.h[1];
        else if(/Heading\\s*3/i.test(n)&&ps.h[2])p.appliedParagraphStyle=ps.h[2];
        else if(/Heading\\s*4/i.test(n)&&ps.h[3])p.appliedParagraphStyle=ps.h[3];
        else if(/Heading\\s*5/i.test(n)&&ps.h[4])p.appliedParagraphStyle=ps.h[4];
        else if(/Heading\\s*6/i.test(n)&&ps.h[5])p.appliedParagraphStyle=ps.h[5];
        else if(/Normal/i.test(n)&&ps.body)p.appliedParagraphStyle=ps.body;
        else if(ps.body)p.appliedParagraphStyle=ps.body;
    }

    // ---- 6) Bullets/Sub-Bullets ----
    var bulletGlyphs=/^[\\u2022\\u25CF\\u25E6\\uF0B7\\u2219\\u00B7○◦\\-–]+\\s*/;
    for (var j=0;j<story.paragraphs.length;j++){
        var para=story.paragraphs[j], txt=para.contents;
        if(bulletGlyphs.test(txt)){
            var indent=(txt.match(/^\\t+/)||[""])[0].length;
            para.contents=txt.replace(bulletGlyphs,"").replace(/^\\t+/,'');
            if(indent>0&&ps.sub)para.appliedParagraphStyle=ps.sub;
            else if(ps.bullets)para.appliedParagraphStyle=ps.bullets;
        }
    }

    // ---- 7) Bold → Body - Bold ----
    if(cs.bold){
        for(var c=0;c<story.characters.length;c++){
            var ch=story.characters[c];
            try{if(ch.fontStyle&&/bold/i.test(ch.fontStyle))ch.appliedCharacterStyle=cs.bold;}catch(_){}
        }
    }

    // ---- 8) Table Style ----
    if(ps.table && story.tables.length){
        for(var t=0;t<story.tables.length;t++)try{story.tables[t].appliedTableStyle=ps.table;}catch(_){}
    }

    // ---- 9) Hyperlink Highlight (robust) ----
    if(cs.link){
        // (a) Real hyperlinks
        for(var h=0;h<doc.hyperlinks.length;h++){
            try{
                var l=doc.hyperlinks[h], d=l.destination, s=l.source;
                if(d && d.destinationURL && /^(https?|mailto):/i.test(d.destinationURL) && s instanceof HyperlinkTextSource)
                    s.sourceText.appliedCharacterStyle=cs.link;
            }catch(_){}
        }

        // (b) Manual GREP for visible URLs (handles fragmented text)
        var urlPattern = /\\b((https?|ftp|file|mailto):\\/\\/[^\\s<>\"]+|www\.[^\\s<>\"]+)\\b/g;
        var storyText = story.contents;
        var matches = [];
        var m;
        while ((m = urlPattern.exec(storyText)) !== null) {
            matches.push({start:m.index, end:m.index+m[0].length});
        }
        for (var k=0;k<matches.length;k++){
            try{
                var r = story.characters.itemByRange(matches[k].start, matches[k].end-1);
                r.appliedCharacterStyle = cs.link;
            }catch(_){}
        }
    }

    story.recompose();
    alert("✅ Import complete.\nHyperlinks, Headings, Bullets, Bold & spacing fully applied.");

} catch(err){
    alert("❌ "+err.message);
    $.writeln("❌ "+err.message);
} finally {
    app.scriptPreferences.userInteractionLevel = originalInteraction;
}

