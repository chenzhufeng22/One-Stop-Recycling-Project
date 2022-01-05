function calculateGrossProfit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = ss.getDataRange().getValues();

  for(var i = 2; i<=data.length;i++){
    if(data[i-1][2] == "G1" ||data[i-1][2] == "G2" || data[i-1][2] == "G4"){   //for G1,G2,G4
      if(data[i-1][8].toLowerCase().indexOf("hd200")>-1){
        ss.getRange(i,16).setValue(0.04);
        ss.getRange(i,17).setValue(data[i-1][11]*0.13);
      }else if(data[i-1][8].toLowerCase().indexOf("ag")>-1&&data[i-1][8].toLowerCase().indexOf("hd")>-1){
        ss.getRange(i,16).setValue(0.15);
        ss.getRange(i,17).setValue(data[i-1][11]*0.15);
      }else if(data[i-1][8].toLowerCase().indexOf("hd")>-1||data[i-1][8].toLowerCase().replace(/\s/g,'').indexOf("pp")>-1||data[i-1][8].toLowerCase().replace(/\s/g,'').indexOf("pmma")>-1||
        data[i-1][8].toLowerCase().replace(/\s/g,'').indexOf("acrylic")>-1){
        ss.getRange(i,16).setValue(0.13);
        ss.getRange(i,17).setValue(data[i-1][11]*0.13);
        //need to get PET price
      }else if(data[i-1][8].indexOf("402")>-1||data[i-1][8].indexOf("805")>-1||data[i-1][8].indexOf("806")>-1){
         ss.getRange(i,16).setValue(0.12);
         ss.getRange(i,17).setValue(data[i-1][11]*0.12);
      }else if(data[i-1][8].indexOf("403")>-1){
         ss.getRange(i,16).setValue(0.05);
         ss.getRange(i,17).setValue(data[i-1][11]*0.05);
      }else if(data[i-1][8].toLowerCase().indexOf("pete")>-1||data[i-1][8].toLowerCase().indexOf("pet")>-1){
          if(data[i-1][8].toLowerCase().indexOf("104")>-1){
            ss.getRange(i,16).setValue(0.1);
            ss.getRange(i,17).setValue(data[i-1][11]*0.13);
          }else if(data[i-1][8].toLowerCase().indexOf("mc")>-1||data[i-1][8].toLowerCase().indexOf("cl")>-1||data[i-1][8].toLowerCase().indexOf("clear")>-1){
            ss.getRange(i,16).setValue(0.2);
            ss.getRange(i,17).setValue(data[i-1][11]*0.2);
          }else if(data[i-1][8].toLowerCase().indexOf("106")>-1||data[i-1][8].toLowerCase().indexOf("smoke")>-1||data[i-1][8].toLowerCase().indexOf("110")>-1){
            ss.getRange(i,16).setValue(0.12);
            ss.getRange(i,17).setValue(data[i-1][11]*0.12);
          }else if(data[i-1][8] != ""){
            ss.getRange(i,16).setValue(0);
            ss.getRange(i,17).setValue(0);
          }
      }else if(data[i-1][8].indexOf("CMI")>-1){
          ss.getRange(i,16).setValue(0.075);
          ss.getRange(i,17).setValue(data[i-1][11]*0.075);
      }else if(data[i-1][8].indexOf("Mul")>-1){
          ss.getRange(i,16).setValue(0.17);
          ss.getRange(i,17).setValue(data[i-1][11]*0.17);
       }
    }else if(data[i-1][2].indexOf("G5")>-1){
      ss.getRange(i,16).setValue(0.11);
      ss.getRange(i,17).setValue(data[i-1][11]*0.11);
    }else if(data[i-1][2].indexOf("S1/G3")>-1){
      if(data[i-1][8].indexOf("HD200")>-1){
        ss.getRange(i,16).setValue(0.08);
        ss.getRange(i,17).setValue(data[i-1][11]*0.08);
      }else if(data[i-1][8].indexOf("209")>-1){
        ss.getRange(i,16).setValue(0.1);
        ss.getRange(i,17).setValue(data[i-1][11]*0.1);
      }

    }else if(data[i-1][2]=="S1"){
      if(data[i-1][8].indexOf("LD")>-1){
        ss.getRange(i,16).setValue(0.05);
        ss.getRange(i,17).setValue(data[i-1][11]*0.05);
      }else if(data[i-1][8].indexOf("CMI")>-1){
        ss.getRange(i,16).setValue(0.075);
        ss.getRange(i,17).setValue(data[i-1][11]*0.075);
      }else{
        ss.getRange(i,16).setValue(0.04);
        ss.getRange(i,17).setValue(data[i-1][11]*0.04);
      }
    }else if(data[i-1][2] == "B1"||data[i-1][2] == "B2"){
      if(data[i-1][8].toLowerCase().indexOf("abs")>-1||data[i-1][8].toLowerCase().indexOf("pc")>-1||data[i-1][8].toLowerCase().indexOf("pmma")>-1||data[i-1][8].toLowerCase().indexOf("acrylic")>-1){
        ss.getRange(i,16).setValue(0.08);
        ss.getRange(i,17).setValue(data[i-1][11]*0.08);
      }else if(data[i-1][8].toLowerCase().indexOf("pet")>-1){
        ss.getRange(i,16).setValue(0.05);
        ss.getRange(i,17).setValue(data[i-1][11]*0.05);
      }else if(data[i-1][8].toLowerCase().indexOf("film")>-1){
        ss.getRange(i,16).setValue(0.04);
        ss.getRange(i,17).setValue(data[i-1][11]*0.04);
      }else if(data[i-1][8] != ""){
        ss.getRange(i,16).setValue(0);
        ss.getRange(i,17).setValue(0);
      }
    }

  }//close for(var i = 2; i<=data.length;i++)

  //calculate daily/monthly total
  var startPosition = 2;
  var monthTotal = 0;
  var dayTotal = 0;
  data = ss.getDataRange().getValues();
  for(var i = 2; i<=data.length;i++){
    if(data[i-1][2] == "Total"){
      for(var j = startPosition; j < i; j++){
        if(data[j-1][16] != ""){
          dayTotal = data[j-1][16] + dayTotal;
        }
      }
      ss.getRange(i,17).setValue(dayTotal).setFontColor("Red");;
      monthTotal = monthTotal + dayTotal;
      startPosition = i + 1;
      dayTotal = 0;
    }
  }
  ss.getRange(data.length,16).setValue("Month Total").setFontColor("Red");
  ss.getRange(data.length,17).setValue(monthTotal).setFontColor("Red");
}

function makeYearlySummary(){
  var ssSummary = SpreadsheetApp.getActive().getSheetByName("Year Summary");
  var sheetNames =["January", "Feburary", "March", "April"];
  var ss;
  var data;
  var startPosition = 3;
  for (var i = 0; i < sheetNames.length; i++){
    ss = SpreadsheetApp.getActive().getSheetByName(sheetNames[i]);
    data = ss.getDataRange().getValues();
    for(var j = 2; j<=data.length;j++){
      if(data[j-1][2] == "Total"){
        ssSummary.getRange(startPosition, (i+1)*2).setValue(data[j-1][16]);
        ssSummary.getRange(startPosition, (i+1)*2-1).setValue(data[j-1][1]);
        startPosition++;
      }
    }
    startPosition = 3;
  }

}
