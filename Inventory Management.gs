var workBookID = [
  "1IGoJhc-hhxmVksMir8gwlPS3zEPB9py1r8ZMlX5puZQ",
  "1kyg_Z5lamXhuDDHnQj-iiAHsKMc1HoDqz_OIhXaojwQ",
  "14HXIgAZ2JcHOW_rAvP5Gc2JvUuxZK65G1qa9CgTB058",
  "1jTNT0JpybB9WP9xTDQ0eDYud-sidQkM7EIdEoso6MLc",
  "1E_r3uwfkzXTE0vk4z9HKNwzmodgKYzCheGIEniMe14o",
  "1Z8_hF1NFfnogLGVRpLCi9EE_0iMtzbOMUROzsUCOfCo",
  "1rbJZ1pkco2hIUaXy2B1PLX2OBNuZiwTiRcsGZMfW3Uc",
  "1RsPXzspjC6FyiAr9zF0aKTAc2oFU0TZwW-SdcKkTAFc",
  "1Pw08EGjv-EhS-BPs3p210wI0IHUwf3cUSHYzwq1-DlA",
  "1_ms8NMA4znClCjjw4_91cbw1A8QUtiQ_phz4T3mUzcI",
                 ];
var summarySS = SpreadsheetApp.openById("17grMQcnqDpvOycMAUNkiMvpan3LRDUU9a6uM4r1-rQA").getSheets();

function main() {
  //variable determination
  var workBookCounter;
  var sheetCounter;
  var totalGross = 0;
  var totalNet = 0;
  var summaryRowNumber = 7;

  //clear summary sheet
  for(var i = 0; i < summarySS.length; i++){
    summarySS[i].getRange("A7:F1002").clearContent();
  }


  //loop through workbooks
  for(workBookCounter=0; workBookCounter<workBookID.length; workBookCounter++){
    var ss = SpreadsheetApp.openById(workBookID[workBookCounter]).getSheets();
    var summarySheet = summarySS[workBookCounter];
    var grandGross = 0;
    var grandNet = 0;
    //clear the summary sheet


    //loop through sheets
    for(sheetCounter=0; sheetCounter<ss.length; sheetCounter++){
      var sheet = ss[sheetCounter];
      var data = sheet.getDataRange().getValues();
      var boxCount = 0;

      for(var i=5; i<data.length; i++){
        //if no data in gross, then break the loop and run next sheet
        if(data[i][2] == ""){
          break;
        }//end if
        var itemNet = data[i][2] - 75;
        totalGross = totalGross + data[i][2];
        totalNet = totalNet + itemNet;
        sheet.getRange(i+1, 4).setValue(itemNet);
        boxCount++;
        //Logger.log("finished loop reset and go to next sheet");
      }//end for reset variable then go to next sheet
      grandGross = grandGross + totalGross;
      grandNet = grandNet + totalNet;
      sheet.getRange(3, 3).setValue(totalGross);
      sheet.getRange(4, 3).setValue(totalNet);

      //set inventory summary data
      summarySheet.getRange(summaryRowNumber, 1).setValue(data[0][2]);
      summarySheet.getRange(summaryRowNumber, 2).setValue(totalGross);
      summarySheet.getRange(summaryRowNumber, 3).setValue(totalNet);
      summarySheet.getRange(summaryRowNumber, 4).setValue(boxCount);
      summarySheet.getRange(summaryRowNumber, 6).setValue(data[5][5].concat(" ",data[1][2]));//set source and remark

      //reset data
      totalGross = 0;
      totalNet = 0;
      summaryRowNumber++;
    }//end for go to next sheet
    summarySheet.getRange(3, 2).setValue(Date());
    summarySheet.getRange(4, 2).setValue(grandGross);
    summarySheet.getRange(5, 2).setValue(grandNet);
    grandGross = 0;
    grandNet = 0;
    summaryRowNumber = 7;
  }//end for, go to next workbook



}

/*
  for (var i = 2; i < data.length; i++){
    if(data[i-1][2] == "Total"){
      //find names
      for(var j = startPosition; j < i; j++){
        if(data[j-1][3].indexOf("/")>-1){
          tempNames = data[j-1][3].split("/");
          for(var k=0; k<tempNames.length; k++){
            names.push(tempNames[k]);
          }
          tempNames = [];
        }else if(data[j-1][3] !=""){
          names.push(data[j-1][3]);
        }
      }

      names = removeDups(names);
      //find work time for each name
      //Logger.log(names);
      for(var z = 0; z <names.length; z++){
       for(var j = startPosition; j < i; j++){
         if(data[j-1][3].indexOf(names[z])>-1 && data[j-1][10] != null ){
           time = time + data[j-1][10];
            }
       }

        workHourSS.getRange(row, 1).setValue(data[j-1][1]);
        workHourSS.getRange(row, 2).setValue(names[z]);
        workHourSS.getRange(row, 3).setValue(time);
        dailyTimeTotal = dailyTimeTotal + time;


        if(names[z] == "Maurilia"){
         time = 5.5 - time;
        }else if(names[z] == "Saul"){
         time = 5.5 - time;
        }else if(names[z] == "Jose"||names[z] == "Maria"){
          time = 1 - time;
        }else{
         time = 7-time;
        }

        if(time > 0){
          workHourSS.getRange(row, 4).setValue(time);
          missTimeTotal = missTimeTotal + time;
        }else{
          workHourSS.getRange(row, 4).setValue(0);
        }
        row++;
        time = 0;
      }
      names = [];
      startPosition = i + 1;
      workHourSS.getRange(row, 1).setValue(data[j-1][1]);
      workHourSS.getRange(row, 2).setValue("Total");
      workHourSS.getRange(row, 3).setValue(dailyTimeTotal);
      workHourSS.getRange(row, 4).setValue(missTimeTotal);
      row=row+2;
      monthTimeTotal = parseFloat(monthTimeTotal) + parseFloat(dailyTimeTotal);
      monthMissTimeTotal = monthMissTimeTotal + missTimeTotal;
      Logger.log(monthTimeTotal);
      Logger.log(dailyTimeTotal);
      Logger.log('--');
      dailyTimeTotal=0;
      missTimeTotal = 0;

    }



  }
      workHourSS.getRange(row, 1).setValue("Month");
      workHourSS.getRange(row, 2).setValue("Total");
      workHourSS.getRange(row, 3).setValue(monthTimeTotal);
      workHourSS.getRange(row, 4).setValue(monthMissTimeTotal);

}
*/
