//cd Volumes/HD3/Esto\ no\ es\ una\ escuela/Corriendo\ por\ las\ olas/Transcripciones/Maquetación/Scripts/

main();
function main(){
    var template = "transcriptions";
    var myDoc = createFromTemplate(template);
    set_Word_import_preferences ();
    var text = getWordDocument();
    placeWordDocument(myDoc, text);

    var pages = myDoc.pages;
    var numPages = pages.length;
    var tableStyle = createTableStyle(myDoc, "Test Style");
    //Tablas anidadas en Archivos de vídeo y Archivos de audio
    var specialTable1 = pages.item(0).allPageItems[0].tables.item(4).cells.item(0).tables.item(0).appliedTableStyle = "Test Style";
    var specialTable2 = pages.item(0).allPageItems[0].tables.item(4).cells.item(1).tables.item(0).appliedTableStyle = "Test Style";
    for(var j = 0; j < numPages; j++){
       
        var firstPage = app.activeDocument.pages.item(j);
        var textFrame = firstPage.allPageItems[0];
        var tables = textFrame.tables;
        var numTables = tables.length;
        for(var i = 0; i < numTables; i++){
            var nTable = tables.item(i).appliedTableStyle = "Test Style";
        }
    }
  }

function createTableStyle(file, name){
    /******************* HERE ******************************************/
    var style = file.tableStyles.add({name: name});
    return style;
}


function createFromTemplate(template){
    //Creates a new document using the specified document preset.
    var myDocument = app.documents.add(true, app.documentPresets.item(template));
    return myDocument;
}

function placeWordDocument(file, text){
    //Get the current page.
    var myPage = file.pages.item(0);
    //Get the top and left margins to use as a place point.
    var myX = myPage.marginPreferences.left;
    var myY = myPage.marginPreferences.top;
    //Autoflow a text file on the current page.
    //Parameters for Page.place():
    //File as File object,
    //[PlacePoint as Array [x, y]]
    //[DestinationLayer as Layer object]
    //[ShowingOptions as Boolean = False]
    //[Autoflowing as Boolean = False]
    //You'll have to fill in your own file path.
    var myStory = myPage.place(File(text), [myX, myY], undefined, false, true)[0]; 
    //Note that if the PlacePoint parameter is inside a column, only the vertical (y) //coordinate will be honored--the text frame will expand horizontally to fit the column.
}

function getWordDocument(){

    var path = new File("~/desktop");
    var text = path.openDlg("Elige el archivo:");
    return text;
}

function set_Word_import_preferences ()
    {
    app.wordRTFImportPreferences.useTypographersQuotes = true;

    app.wordRTFImportPreferences.convertPageBreaks = ConvertPageBreaks.PAGE_BREAK;
    //~ app.wordRTFImportPreferences.convertPageBreaks = ConvertPageBreaks.columnBreak;
    //~ app.wordRTFImportPreferences.convertPageBreaks = ConvertPageBreaks.pageBreak;

    app.wordRTFImportPreferences.importEndnotes = true;
    app.wordRTFImportPreferences.importFootnotes = true;
    app.wordRTFImportPreferences.importIndex = true;
    app.wordRTFImportPreferences.importTOC = false;
    app.wordRTFImportPreferences.importUnusedStyles = false;
    app.wordRTFImportPreferences.preserveGraphics = true;
    app.wordRTFImportPreferences.convertBulletsAndNumbersToText = true;

    app.wordRTFImportPreferences.removeFormatting = false;
    // If removeFormatting is true, these two can be set as well:
    //~ app.wordRTFImportPreferences.convertTablesTo = ConvertTablesOptions.unformattedTabbedText;
    //~ app.wordRTFImportPreferences.convertTablesTo = ConvertTablesOptions.unformattedTable;
    //~ app.wordRTFImportPreferences.preserveLocalOverrides = true

    app.wordRTFImportPreferences.preserveTrackChanges = false;

    //~ app.wordRTFImportPreferences.resolveCharacterStyleClash = ResolveStyleClash.resolveClashAutoRename;
    //~ app.wordRTFImportPreferences.resolveCharacterStyleClash = ResolveStyleClash.resolveClashUseExisting;
    //~ app.wordRTFImportPreferences.resolveCharacterStyleClash = ResolveStyleClash.resolveClashUseNew;

    //~ app.wordRTFImportPreferences.resolveParagraphStyleClash = ResolveStyleClash.resolveClashAutoRename;
    //~ app.wordRTFImportPreferences.resolveParagraphStyleClash = ResolveStyleClash.resolveClashUseExisting;
    //~ app.wordRTFImportPreferences.resolveParagraphStyleClash = ResolveStyleClash.resolveClashUseNew;
    }
