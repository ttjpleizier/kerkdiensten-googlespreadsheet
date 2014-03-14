function preekplanning(){

//beginwaarden
var sheet = SpreadsheetApp.getActiveSheet();
var startRow = 2;  
var numRows = 124;   
sheet.getRange(2,2,numRows,1).setNumberFormat('@STRING@');

//initialiseer de selectie 
var monthNames = [ "januari", "februari", "maart", "april", "mei", "juni",
    "juli", "augustus", "september", "oktober", "november", "december" ];
var startdatum = Browser.inputBox('Begin-maand (nummer):', Browser.Buttons.OK_CANCEL);
    if (startdatum == 'cancel') return; 
var einddatum = Browser.inputBox('Eind-maand (nummer):', Browser.Buttons.OK_CANCEL);
    if(einddatum == 'cancel') return;
var datumnu = new Date();
var jaar = datumnu.getFullYear();
var plaats="PLAATSNAAM";
var laatsterij=0;
var eersterij=0;
var datumcolumn = sheet.getRange(startRow,1,numRows,1).getValues();
for(var ix = 0; ix < datumcolumn.length; ix++) {
  var datumrij=new Date(datumcolumn[ix]);
  var maand=datumrij.getMonth();
  if (maand==(startdatum-1) && eersterij ==0) {
    eersterij=(ix+2);
  }
  if (maand==(einddatum) && laatsterij ==0) {
    laatsterij=(ix+2);
  }
}
   if (einddatum==12){
     var aantalrijen = (numRows-eersterij);
     aantalrijen++;
   } else {
     var aantalrijen = (laatsterij-eersterij);
   }
laatsterij--;
var docrange=sheet.getRange(eersterij,1,aantalrijen,10).getValues();//dit is de range die in de tabel preekplanning terecht komt
var bestandnaam=("preekplanning_"+plaats+"_"+jaar+"_"+startdatum+"-"+einddatum);
  
//opbouwen array voor de tabel in het nieuwe bestand
var labels=[];
var vorigedatum = new Date();
vorigedatum.setFullYear(1975,1,1);
var z=0;
labels[0]=["datum","ochtenddienst","avonddienst", "",""]
for (y in docrange){
  var row=docrange[y];
  var gemeente=row[2];
  var datum=new Date(docrange[y][0]);
  Logger.log(z+" | "+gemeente+" | nu: "+datum+" vorigedatum: "+vorigedatum);
  if (gemeente==plaats){
    z++;
    var dag=datum.getDate();
    var maand=datum.getMonth();
    var sheettime=row[1];
    var n=String(sheettime).split(":");
    var uur=n[0];
    var ochtend="";
    var avond="";
    var ochtendplus="";
    var avondplus="";
    if (uur < 12){
       var ampm="ochtend";
       var ochtend=row[7];
       var ochtendplus=row[4];
       } else if (uur <18){//dit wordt alleen gebruikt als er géén ochtenddienst is
        var ampm="middag";
           var ochtend=row[7];
           var ochtendplus="N.B.middagdienst! "+row[4];
       } else {
        var ampm="avond";
        var avond=row[7];
        var avondplus=row[4];
        }
    if (datum > vorigedatum){
    labels[z]=[];
    labels[z].push(dag+" "+monthNames[maand],ochtend,avond,ochtendplus,avondplus);
    } else {//als er dezelfde dag nog één of meer diensten zijn
    var zz=--z;
      if (uur < 18){
        labels[zz][1]="vm. "+labels[zz][1]+" | nm. "+row[7];
        labels[zz][3]="vm. "+labels[zz][3]+" | nm. "+row[4];
      } else {
        labels[zz][2]=avond;
        labels[zz][4]=avondplus;
      }
   }
   var vorigedatum=new Date(docrange[y][0]);
   }
}
  
//opmaak van het nieuwe bestand
var style = {};
   style[DocumentApp.Attribute.FONT_SIZE] = 11;
   style[DocumentApp.Attribute.BOLD]=false;
   style[DocumentApp.Attribute.ITALIC]=false;
var toevoeging = {}; 
   toevoeging[DocumentApp.Attribute.BOLD]=false;
   toevoeging[DocumentApp.Attribute.ITALIC]=true;
   toevoeging[DocumentApp.Attribute.FONT_SIZE]=9;
   toevoeging[DocumentApp.Attribute.SPACING_AFTER]=0;
var cellstyle = {}; 
   cellstyle[DocumentApp.Attribute.ITALIC]=false;
   cellstyle[DocumentApp.Attribute.BOLD]=true;
   cellstyle[DocumentApp.Attribute.FONT_SIZE]=11;
var geenruimte ={};
   geenruimte[DocumentApp.Attribute.SPACING_AFTER]=0;

//creeer document
var doc = DocumentApp.create(bestandnaam);
if (startdatum==einddatum){
  var titleplanning ='Preekplanning voor '+monthNames[(startdatum-1)]+" "+jaar+" | "+plaats
} else {
  var titleplanning ='Preekplanning voor '+monthNames[(startdatum-1)]+" t/m "+monthNames[(einddatum-1)]+" "+jaar+" | "+plaats
}
var body = doc.getBody();
body.insertParagraph(0,titleplanning).setHeading(DocumentApp.ParagraphHeading.HEADING1);
body.insertParagraph(1,'naam predikant').setHeading(DocumentApp.ParagraphHeading.HEADING3);

//per cell vullen
var table=body.appendTable();
for (zzz in labels){ 
  var tr=table.appendTableRow();
  var td=[];
  
  td[0]=tr.appendTableCell(labels[zzz][0]);
   td[0].setAttributes(style); 
   var paraInCel = td[0].getChild(0).asParagraph();
   paraInCel.setAttributes(geenruimte);
  
  td[1]=tr.appendTableCell(labels[zzz][1]);
   td[1].setAttributes(cellstyle);
   var paraInCell = td[1].getChild(0).asParagraph();
   paraInCell.setAttributes(geenruimte);
   td[1].appendParagraph(labels[zzz][3]).setAttributes(toevoeging); 
  
  td[2]=tr.appendTableCell(labels[zzz][2]);
   td[2].setAttributes(cellstyle);
   var paraInCelll = td[2].getChild(0).asParagraph();  
     paraInCelll.setAttributes(geenruimte);
  td[2].appendParagraph(labels[zzz][4]).setAttributes(toevoeging);
  if (zzz == 0) {
    tr.setAttributes(style);
  }
  
}
//afhandelen bestand
var file = DocsList.getFileById(doc.getId());
file.removeFromFolder(DocsList.getRootFolder());
file.addToFolder(DocsList.getFolder("google drive map"));
}
