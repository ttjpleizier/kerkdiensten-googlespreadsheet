function Sheet2Cal() {
  var sheet = SpreadsheetApp.getActiveSheet();//een werkblad is in gdocs een sheet
  var startRow = 2;
  var numRows = 123;
  var dataRange = sheet.getRange(startRow, 1, numRows, 10);
  sheet.getRange(2,2,numRows,1).setNumberFormat('@STRING@');//de kolom met tijdaanduiding (kolom 2!) wordt gezet in plain text
  var data = dataRange.getValues();
  var cal = CalendarApp.getCalendarById("agendaid@google.com");
  var rij = startRow-1;

//nagaan of er wijzigingen zijn gemarkeerd in kolom 11 (edited)
  var wijziging=false;
  var wijzigcolumn = sheet.getRange(startRow,11,numRows,11).getValues();
  for(var ix = 0; ix < wijzigcolumn.length; ix++) {
  if(wijzigcolumn[ix] == '*') {
  var wijziging=true
  }
  }
  var editonly = 'no';
  if (wijziging) {
  var editonly = Browser.msgBox('Alleen gewijzigde rijen updaten?', Browser.Buttons.YES_NO)
  }

//centrale loop
  for (i in data) {
    var row = data[i];
    rij++; 
    Logger.log(rij+" "+i);
    var sheetdate = sheet.getRange(rij,1).getValue();
    var calendardate = new Date(sheetdate);
    var sheettime = sheet.getRange(rij,2).getValue();
    var n=String(sheettime).split(":");
    calendardate.setHours(calendardate.getHours()+n[0]);
    calendardate.setMinutes(calendardate.getMinutes()+n[1]);
    var tstop = new Date(calendardate);
    tstop.setMinutes(tstop.getMinutes()+90); //uitgaande van een kerkdienst van 90 minuten
    var gemeente = row[2]; 
    var contact = row[3];
    var desc = "Contact: "+ contact + " |  " + "\r\n\r\n(imported via gdocs)";  
    if (row[4] !=="") {
      var bijzonderheden = " ("+ row[4] + ")";
      } else {
        var bijzonderheden = "";
      }
    if (row[7] !== "") { 
      var tekst = " - " + row[7];
      } else { 
        var tekst ="";
      }
    var eventtitle= gemeente + tekst + bijzonderheden;
    //var thema = row[9];
    var eventidsheet=SpreadsheetApp.getActiveSheet().getRange(rij,10).getValue();
    var editedjn=SpreadsheetApp.getActiveSheet().getRange(rij,11).getValue();

    //test of EventId is aangemaakt, zo ja, verwijder event, 
    if (eventidsheet !== '' && (editonly !== 'yes'|| (editonly == 'yes' && editedjn !==''))) {
      var event = cal.getEventSeriesById(eventidsheet);
      try{
        event.deleteEventSeries(); 
      }
      catch(e){
        if (e.name.toString()!=="") // welke error dan ook 
        Browser.msgBox("Eventid " + eventidsheet + " met titel: " + event.getTitle() + " niet aanwezig in Google Calendar.");
          }    
    }

//maak nieuw event aan met de data uit de sheet
    if (editonly !== 'yes'|| (editonly == 'yes' && editedjn !=='')){
    var event=cal.createEvent(eventtitle, calendardate, tstop, {description:desc});//,location:loc});
    var eventid = event.getId();
    SpreadsheetApp.getActiveSheet().getRange(rij,10).setValue(eventid);
    if (editedjn !==''){SpreadsheetApp.getActiveSheet().getRange(rij,11).setValue('')}
    }
}
}
 
