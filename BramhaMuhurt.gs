/**
* @author: Rajesh Swarnkar <rjs.swarnkar@gmail.com>
 * Calculator made for Bramh Muhurt
 */

var my_recepient = "youremail@gmail.com"


function bramhaMuhurt() {


  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Muhurt").getRange("A2:I2").getValues(); 

  var place_name   = data[0][0]  // Static
  var lat          = data[0][1]  // Static
  var long         = data[0][2]  // Static
  var nextdate     = data[0][3]  // Static
  var sunrise      = data[0][4]
  var sunset       = data[0][5]
  var bm_start     = data[0][6]
  var bm_end       = data[0][7]
  var bm_duration  = data[0][8]


  // First calculate Sunrise and Sunset Time and set the value in Excel Sheet
  // Logic Copied from: https://stackoverflow.com/questions/48040528/import-sunrise-set-based-on-coordinates-into-google-sheet-using-api
  // https://api.sunrise-sunset.org/json?lat=19.139542640955355&lng=73.25242208466288&date=tomorrow&formatted=0
  var date_fmt = Utilities.formatDate(nextdate, "GMT+0530", "yyyy-MM-dd")
  Logger.log("Next Day       : "+date_fmt)



  var response = UrlFetchApp.fetch("https://api.sunrise-sunset.org/json?lat="+lat+"&lng="+long+"&date="+date_fmt+"&formatted=0");
  var json = response.getContentText();
  var url_data = JSON.parse(json);
  var sunrise = url_data.results.sunrise;
  var sunset = url_data.results.sunset;

  // Format the date objects: https://developers.google.com/apps-script/reference/utilities/utilities#formatDate(Date,String,String)
  sunrise = new Date(sunrise)
  sunset  = new Date(sunset)

  sunrise_temp = Utilities.formatDate(sunrise, "GMT+0530", "HH:mm:ss")
  sunset_temp  = Utilities.formatDate(sunset, "GMT+0530", "HH:mm:ss")

  Logger.log("Sunrise      : "+sunrise_temp)
  Logger.log("Sunset       : "+sunset_temp)

  var values = [
       [sunrise,sunset]
   ];

  var range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Muhurt").getRange("E2:F2");
  range.setValues(values);

// Now calculate Bramha Muhurt times and set the value in Excel Sheet
// https://en.wikipedia.org/wiki/Brahmamuhurtha
// Brahmamuhurtha (time of Brahma) is a period (muhurta) that begins 96 minutes before sunrise, and ends 48 minutes later.  
// One muhurtha is a period of 48 minutes.

  var sunrise_millis = sunrise.getTime(); // in milis

  var before96min_sunrise =  -1 * 96 * 60 * 1000  // -96 minutes in milliseconds
  var after48min_sunrise  =       48 * 60 * 1000  // +48 minutes in milliseconds


  bm_start = sunrise_millis + before96min_sunrise; // Bramha Muhurt Start
  bm_end   = sunrise_millis + after48min_sunrise; // Bramha Muhurt End
  bm_duration = ( (bm_end - bm_start) / 1000.0 ) / 60.0 // Duration in Minutes

  bm_start = new Date(bm_start)
  bm_end   = new Date(bm_end)

  bm_start_temp = Utilities.formatDate(bm_start, "GMT+0530", "HH:mm:ss")
  bm_end_temp   = Utilities.formatDate(bm_end, "GMT+0530", "HH:mm:ss")
  bm_duration_temp = bm_duration

  var values = [
       [bm_start,bm_end, bm_duration]
   ];

  Logger.log("Bramha Muhurt Start      : "+bm_start_temp)
  Logger.log("Bramha Muhurt End       : "+bm_end_temp)
  Logger.log("Bramha Muhurt Duration (mins)       : "+bm_duration)

  var range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Muhurt").getRange("G2:I2");
  range.setValues(values);

/// ---------------------- Send via email ----------------------

var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Muhurt").getRange("A2:I2").getValues(); 

  var place_name   = data[0][0]  // Static
  var lat          = data[0][1]  // Static
  var long         = data[0][2]  // Static
  var nextdate     = data[0][3]  // Static
  var sunrise      = data[0][4]
  var sunset       = data[0][5]
  var bm_start     = data[0][6]
  var bm_end       = data[0][7]
  var bm_duration  = data[0][8]

  my_subject = "Brahma Muhurt | " +  Utilities.formatDate(nextdate, "GMT+0530", "dd-MMM-yyyy")

  var table_format = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'

  var htmltable = '<br /> <table ' + table_format +' ">';

  htmltable += '<tr>';
  htmltable += '<th>Place Name</th>';
  htmltable += '<th>Date</th>';
  htmltable += '<th>Sunrise Time</th>';
  htmltable += '<th>Bramha Muhurt Start</th>';
  htmltable += '<th>Bramha Muhurt End</th>';
  htmltable += '<th>Bramha Muhurt Duration</th>';
  htmltable += '</tr>';

  htmltable += '<tr>';
  htmltable += '<td>' + place_name + '</td>';
  htmltable += '<td>' + Utilities.formatDate(nextdate, "GMT+0530", "dd-MMM-yyyy") + '</td>';
  htmltable += '<td>' + Utilities.formatDate(sunrise, "GMT+0530", "HH:mm:ss") + '</td>';
  htmltable += '<td>' + Utilities.formatDate(bm_start, "GMT+0530", "HH:mm:ss") + '</td>';
  htmltable += '<td>' + Utilities.formatDate(bm_end, "GMT+0530", "HH:mm:ss") + '</td>';
  htmltable += '<td>' + bm_duration + '</td>';
  htmltable += '</tr>';

  htmltable += '</table>';

  mainMessageBody = "Brahma Muhurt begins 96 minutes before sunrise, and ends 48 minutes after. <br />" + htmltable +  "<br />"

  MailApp.sendEmail(my_recepient, my_subject,'' ,{htmlBody: mainMessageBody})
  Logger.log("Mail Sent.");


}
 
