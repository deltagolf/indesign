main();
function main(){
    var myDoc = app.activeDocument;
    var pages = app.activeDocument.pages;
    var numPages = pages.length;
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
    //Success!!!!!!!!
  }