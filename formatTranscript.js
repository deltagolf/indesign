﻿//Path for git repository//cd Volumes/HD3/Esto\ no\ es\ una\ escuela/Corriendo\ por\ las\ olas/Transcripciones/Maquetación/Scripts/main();function main(){    //Declares path to find documents ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++    var absolutePath = "/Volumes/HD3/Esto\ no\ es\ una\ escuela/Corriendo\ por\ las\ olas/Transcripciones/Maquetación/";    var pathPlaceDocs = absolutePath + "Pruebas";    //Path where template document is located ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++    var template = absolutePath + "Archivos\ InDesign/plantilla_maestra.indt";    //List of styles    var headerStyle = "Encabezado página";    //var titleCellStyle = "Título";    //var titleParagraphStyle = "Título";    var firstPageTableStyles = ["Título", "Datos entrevista", "Etiquetas", "Participantes", "Archivos", "Documentos"];    //Open document from template ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++    var myDoc = app.open(File(template), true, OpenOptions.DEFAULT_VALUE);    //Sets import preferences for Word docs ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++    set_Word_import_preferences ();    //Gets the contents for a Word Document ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++    var text = getWordDocument(pathPlaceDocs);    //Places the text in the created file ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++    placeWordDocument(myDoc, text);    //Replace With Text Variable ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++    var printingDate = myDoc.textVariables.itemByName("Printing Date");    var placeHolderPrintingDate = findText(myDoc, "%P_DATE%");    //Replace placeholders ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++    for(var i =0; i < placeHolderPrintingDate.length; i++){        placeHolderPrintingDate[i].texts[0].textVariableInstances.add({associatedTextVariable:printingDate});    }    findAndReplace(myDoc, "%P_DATE%", "");    //Gets the set of pages of the doc ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++    var pages = myDoc.pages;    //Stores number of pages    var numPages = pages.length;    var firstPage = pages[0];    var firstFrame = firstPage.textFrames[0];    var firstParagraph = firstFrame.paragraphs[0];    firstParagraph.appliedParagraphStyle = headerStyle;    var firstText = firstFrame.texts[0];    var firstContents = firstText.contents;    //Gets file name    var docName = firstContents.slice(0,12);    //Format tables ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++    for(var j = 0; j < numPages; j++){        var page = app.activeDocument.pages.item(j);        var textFrame = page.allPageItems[0];        var tables = textFrame.tables;        var numTables = tables.length;        var nTable = [];        var tempTable;        if(j == 0){            for(var i = 0; i < numTables; i++){                tempTable = tables.item(i);                tempTable.appliedTableStyle = firstPageTableStyles[i];                tempTable.clearTableStyleOverrides();                nTable.push(tempTable);            }            //Format first table HEADER ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++            tempTable = nTable[0];            var cell = tempTable.cells.item(0);            cell.height = 80;            cell.width = 161;            cell.appliedCellStyle = "Título";            cell.clearCellStyleOverrides();            //Format second table DATA ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++            tempTable = nTable[1];            for(var a = 0; a < tempTable.cells.length; a++){                cell = tempTable.cells[a];                cell.appliedCellStyle = "Datos entrevista";                cell.clearCellStyleOverrides();            }            tempTable.columns[0].width = 30;            tempTable.columns[1].width = 90;            //Format third table LABELS            tempTable = nTable[2];            tempTable.columns[0].width = 2;            tempTable.columns[1].width = 115;            tempTable.columns[2].width = 2;            var cellsStyles = ["Etiquetas izq", "Etiquetas", "Etiquetas der"];            for(var a = 0; a < tempTable.cells.length; a++){                cell = tempTable.cells.item(a);                cell.appliedCellStyle = cellsStyles[a];                cell.clearCellStyleOverrides();                }               //Format fourth table PARTICIPANTS            tempTable = nTable[3];            tempTable.rows[0].merge();            cell = tempTable.rows[0].cells[0];            cell.appliedCellStyle = "Participantes";            cell.clearCellStyleOverrides();            tempTable.rows[0].cells[0].bottomInset = 3;            var widths = [10, 90, 20];            for(var a = 0; a < 3; a++){                cell = tempTable.rows[1].cells.item(a);                cell.appliedCellStyle = "Nombres participantes";                cell.clearCellStyleOverrides();                tempTable.columns[a].width = widths[a];            }            var row;            for(var a = 2; a < tempTable.rows.length; a++){                row = tempTable.rows[a];                for(var b = 0; b < 3; b++){                    cell = row.cells[b];                    var text = row.cells[0].paragraphs[0].texts[0].contents;                    cell.appliedCellStyle = "Participantes";                    cell.clearCellStyleOverrides();                    if(text == "P"){                        cell.paragraphs[0].fontStyle = "Bold";                    }                }            }            //Format fifth table FILES            tempTable = nTable[4];            var specialTable;            for(b = 0; b < 2; b++){                tempTable.columns[b].width = 161/2;                var cell = tempTable.cells.item(b);                cell.paragraphs[0].remove();                cell.topInset = 0;                widths = [44, 12, 12, 12];                specialTable = cell.tables.item(0);                specialTable.appliedTableStyle = "Archivos";                specialTable.clearTableStyleOverrides();                var firstRow = specialTable.rows[0];                firstRow.merge();                firstRow.cells[0].appliedCellStyle = "Documentos";                firstRow.cells[0].clearCellStyleOverrides();                firstRow.bottomInset = 3;                                for(var z = 0; z < 4; z++){                    specialTable.columns[z].width = widths[z];                    specialTable.rows[1].cells[z].appliedCellStyle = "Archivos título";                    specialTable.rows[1].cells[z].clearCellStyleOverrides();                }                for(var a = 2; a < specialTable.rows.length; a++){                    row = specialTable.rows[a];                    for(z = 0; z < 4; z++){                        row.cells[z].appliedCellStyle = "Archivos";                        row.cells[z].clearCellStyleOverrides();                    }                }            }            //Format fifth table FILES            tempTable = nTable[5];            widths = [2, 115, 2];            cellsStyles[1] = "Documentos";            for(a = 0; a < 3; a++){                tempTable.columns[a].width = widths[a];                tempTable.cells.item(a).appliedCellStyle = cellsStyles[a];                tempTable.cells.item(a).clearCellStyleOverrides();            }                              } else {        }    }    //Place headers and footers ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++    var headers = createFrames(9, 22, 161, 7, myDoc);    var footers = createFrames(273, 22, 161, 24, myDoc);    //Place contents in headers and footers ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++    for(var i = 0; i < headers.length; i++){        firstParagraph.duplicate(LocationOptions.AT_BEGINNING, headers[i].texts[0]);    }    firstParagraph.remove();    firstParagraph = pages[0].textFrames[2].paragraphs[0].remove();    myDoc.save(new File(absolutePath + "Archivos\ InDesign/" + docName));    //Close document    //myDoc.close();  }function createFrames(x, y, h, w, file){    var pages = file.pages;    var i = 0, pagesLength = pages.length;    var frame, frames = [];    var properties = {        geometricBounds: [x, y, x + w, y + h]    };    for (i; i < pagesLength; i++){        frame = pages[i].textFrames.add(undefined, LocationOptions.AT_BEGINNING, pages[i], properties);        frames.push(frame);    }    return frames;}function selectText(file, text){}function findText(file, find){   //find/change text preferences    app.findTextPreferences = NothingEnum.nothing;     app.changeTextPreferences = NothingEnum.nothing;    app.findTextPreferences.findWhat = find;//    app.changeTextPreferences.changeTo = replace;    var foundItems = file.findText();//   //Clear find/change text preferences    app.findTextPreferences = NothingEnum.nothing;     app.changeTextPreferences = NothingEnum.nothing;    return foundItems;}function findAndReplace(file, find, replace){   //find/change text preferences    app.findTextPreferences = NothingEnum.nothing;     app.changeTextPreferences = NothingEnum.nothing;    app.findTextPreferences.findWhat = find;    app.changeTextPreferences.changeTo = replace;    var foundItems = file.changeText();//   //Clear find/change text preferences    app.findTextPreferences = NothingEnum.nothing;     app.changeTextPreferences = NothingEnum.nothing;}function createFromPreset(preset){    //Creates a new document using the specified document preset.    var myDocument = app.documents.add(true, app.documentPresets.item(preset));    return myDocument;}function placeWordDocument(file, text){    //Get the current page.    var myPage = file.pages.item(0);    //Get the top and left margins to use as a place point.    var myX = myPage.marginPreferences.left;    var myY = myPage.marginPreferences.top;    //Autoflow a text file on the current page.    //Parameters for Page.place():    //File as File object,    //[PlacePoint as Array [x, y]]    //[DestinationLayer as Layer object]    //[ShowingOptions as Boolean = False]    //[Autoflowing as Boolean = False]    //You'll have to fill in your own file path.    var myStory = myPage.place(File(text), [myX, myY], undefined, false, true)[0];     //Note that if the PlacePoint parameter is inside a column, only the vertical (y) //coordinate will be honored--the text frame will expand horizontally to fit the column.}function getWordDocument(path){    var myFile = new File(path);    var text = myFile.openDlg("Elige el archivo:");    return text;}function set_Word_import_preferences ()    {    app.wordRTFImportPreferences.useTypographersQuotes = true;    app.wordRTFImportPreferences.convertPageBreaks = ConvertPageBreaks.PAGE_BREAK;    //~ app.wordRTFImportPreferences.convertPageBreaks = ConvertPageBreaks.columnBreak;    //~ app.wordRTFImportPreferences.convertPageBreaks = ConvertPageBreaks.pageBreak;    app.wordRTFImportPreferences.importEndnotes = true;    app.wordRTFImportPreferences.importFootnotes = true;    app.wordRTFImportPreferences.importIndex = true;    app.wordRTFImportPreferences.importTOC = false;    app.wordRTFImportPreferences.importUnusedStyles = false;    app.wordRTFImportPreferences.preserveGraphics = true;    app.wordRTFImportPreferences.convertBulletsAndNumbersToText = true;    app.wordRTFImportPreferences.removeFormatting = true;    // If removeFormatting is true, these two can be set as well:    //~ app.wordRTFImportPreferences.convertTablesTo = ConvertTablesOptions.unformattedTabbedText;    app.wordRTFImportPreferences.convertTablesTo = ConvertTablesOptions.unformattedTable;    //~ app.wordRTFImportPreferences.preserveLocalOverrides = true    app.wordRTFImportPreferences.preserveTrackChanges = false;    //~ app.wordRTFImportPreferences.resolveCharacterStyleClash = ResolveStyleClash.resolveClashAutoRename;    //~ app.wordRTFImportPreferences.resolveCharacterStyleClash = ResolveStyleClash.resolveClashUseExisting;    //~ app.wordRTFImportPreferences.resolveCharacterStyleClash = ResolveStyleClash.resolveClashUseNew;    //~ app.wordRTFImportPreferences.resolveParagraphStyleClash = ResolveStyleClash.resolveClashAutoRename;    //~ app.wordRTFImportPreferences.resolveParagraphStyleClash = ResolveStyleClash.resolveClashUseExisting;    //~ app.wordRTFImportPreferences.resolveParagraphStyleClash = ResolveStyleClash.resolveClashUseNew;    }