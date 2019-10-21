//Author: MH Charles-Etuk
//Date: 10/2/2019
//Description: Google Script functions to read information from a Google Sheets spreadsheet then use it to send emails and create Google Calendar events as needed



function sendRequestResults() 
{
  var EMAIL_SENT = 'EMAIL_SENT'; //see use in for loop
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Responses")); //use correct worksheet
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;
  var dataRange = sheet.getRange("B2:I20"); //can be expanded to more rows as list grows, for now uses 20 rows
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; i++) //loop through each row in that range
  {
    var rowData = data[i]; //each element of data array contains 8 columns of data, from left to right
    var name = rowData[0]; //column B has names
    var reason = rowData[1]; //column C has reasons
    var date = rowData[2]; //column D has dates
    var startTime = rowData[3]; //column E has start times
    var endTime = rowData[4]; //column F has end times
    var emailAddress = rowData[5]; //column G has email addresses
    var response = rowData[6]; //column H has approval status
    var sent = rowData[7]; //confirmation column
    
    if (sent !== EMAIL_SENT) //prevents duplicate emails from being sent
    {
      if (response == 'yes') //if approved
      {
        var message = 'Hello, ' + name + '!\n\n' + 'Your request for time off on ' + date + ' has been granted.' + '\n\n' + 'In the Google Calendar you will find an all-day event representing your time off. Please modify it to match the time slot you requested. ' + '\n\n' + 'Have a good day.';
        var subject = 'Time off Request Granted';
        MailApp.sendEmail(emailAddress, subject, message);
        sheet.getRange(startRow, 9).setValue(EMAIL_SENT);
      }
      else if (response == 'no') //if denied
      {
        var message = 'Hello,' + name + '!\n\n' + 'Your request for time off on ' + date + ' was not granted.' + '\n\n' + 'Have a good day.';
        var subject = 'Time off Request Not Granted';
        MailApp.sendEmail(emailAddress, subject, message);
        sheet.getRange(startRow, 9).setValue(EMAIL_SENT);
      }
      else
        startRow++;
    }
    else
      startRow++;
  }
}

function sendToCalendar()
{
  var EVENT_CREATED = 'EVENT_CREATED'; //see use in for loop
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Responses"));
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;
  var dataRange = sheet.getRange("K2:P20"); //can also be expanded
  var data = dataRange.getValues();
  
  var calendar = CalendarApp.getCalendarById('ckcghealth.com_f3mjdqj3juime7dd23p0ftl82s@group.calendar.google.com') //id of specific Time Off calendar in CKCG calendar
  
  
  var numValues = 0;
  
  for (var i = 0; i < data.length; i++) //loop through data range like before
  {
    var name = data[i][0]; //column K has names
    var date = data[i][1]; //column L has dates
    var sTime = data[i][2]; //column M has start times
    var eTime = data[i][3]; //column N has end times
    var granted = data[i][4]; //column O has approval status
    var created = data[i][5]; //confirmation column
    
    if (name.length > 0) //if there's a name there
    {
       
      //check if it's been entered before          
      if (created !== 'EVENT_CREATED')
      {                       
        if (granted == 'yes') 
        { 
          var newTitle = 'Time Off: ' + name;
          var newEvent = calendar.createAllDayEvent(newTitle, date); //create event
          
          //var eventStart = new Date( date + " " + sTime); 
          //var eventEnd = new Date( date + " " + eTime);
          //var newTimedEvent = calendar.createEvent(newTitle, eventStart, eventEnd)
          
          sheet.getRange(startRow, 16).setValue(EVENT_CREATED); //mark as created
        }
        else 
         startRow++; //if created go to next row
      } 
      else
        startRow++;
    }
    numValues++;
  }
}
