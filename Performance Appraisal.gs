var dailyProductionRatingSheet = SpreadsheetApp.getActive().getSheetByName('Daily Production Rating');
var productionSummarySheet = SpreadsheetApp.getActive().getSheetByName('Production Summary');
var data = dailyProductionRatingSheet.getDataRange().getValues();

var wage = 18;

var ABS_O_BLK_G2 = 2500;
var ABS_O_BLK_S2 = 2800;
var ABS_O_BLK_S2G5 = 2600;
var ABS_Chrome_B2_1W = 1100;
var ABS_822_AC_WHT_S2G5 = 2400;
var ABS_Contaminated = 1300;
var ABS_Fenders = 1200;
var ABS_pieces_B1_1W = 1000;
var ABS_wCoating_B2_1W = 1400;
var ABSOS_PMMA_S2G5_2W = 2500;
var ABSPC805INJNAT_G2_1W = 1200;
var Acrylic_Cast = 1280;
var Acrylic_Chrome_B2_1W = 1400;
var Acrylic_WithTape_B1_1W = 1100;

var cores = 1100;

var Film_Pepsi = 1100;
var Film_Rebale = 1500;
var Film_B1_2W = 1200;

var HD200_G1_1W = 1900;
var HD200_G2_1W = 2000;
var HD200_shred_s1s2_2W = 2800;
var HD200_S1G3_1W = 1600;
var HDPE_207_INJ = 1000;
var HIPS_604 = 2600;
var HDPE_202_BM_NAT_G2_1W = 500;
var HPDE_203_INJ_MC_S2G5 = 2500;
var HDPE_203_INJ_MC_G1_1W = 1600;
var HDPE_203_INJ_MC_G2_1W = 1650;
var HDPE_204_BM_MC_G2_1W = 221;
var HDPE_204_BM_MC_G1_1W = 600;
var HDPE_205_WF_MC_S1G3_1W = 600;


var LUZ_HD_S1G3_2W = 1500;
var LDPE402INJCAPWHT_G1_1W = 700;
var LDPE_403_EXT_PRG_MC_G2_1W = 1000;
var LDPE_403_EXT_PRG_MC_S1G3_1W = 800;
var LDPE_403_EXT_PRG_MC_S1_1W = 1400;


var MS_905_INJ_CL_G1_1W = 1500;
var multiflexa6221_G1_1W = 600;

var Nylon = 800;

var OCC = 1200;

var PMMA = 1800;
var PMMA_IMP_B1_2W = 1300;
var PMMA_901_IMP_G1G2_2W = 800;
var PMMA_902_G1_2W = 1100;
var PMMA_904_INJ_CL_G2_1W = 450;
var PMMA_WTape_B1 = 900;

var PETE_106 = 1300;
var PET_Green_Strapping = 1500;
var PETE_110_AG_G4_1W= 450;

var PVC_Tubing = 2000;
var Painted_ABS = 1300;
var PP_Craes_MC = 600;
var PP_502_BM_Nat = 1000;
var PP_CR8_MC = 1300;
var PP_511_cr8_MC = 1300;
var PP_INJ_MC_S2G5 = 2000;
var PP_503_INJ_CAPWT_G1_1W = 1843;
var PP_506_INJ_MC_G2_1W = 1700;
var PET_102_BM_MC_G4_1W = 700;
var PET_101_G4_1W = 425;
var PETE_101_BM_CLR = 620;
var PET_104_Tray_G4_1W = 500;
var PP510EXTWT_G1_1W = 150;
var PCABS806INJMC_G2_1W = 800;
var PCABS805INJNAT_G1_1W = 1300;

var sonocoShred_S2G5_2W = 1500;

var TPU_B1 = 1100;
var TPU_B2 = 2000;
var TPO_Nibs = 1050;
var TPO_Contaminated_B2_1W = 1400;
var TPO_3400_S1 = 2183;
var TPO_3400_S2G5_2W = 2000;


function main(){
  setGoal();
  changeColor();
  makeProductionSummary();
  makePersonalRating();
  makeMachineAnalysis();
  makePersonalPerformanceAnalysis();
  makeWorkerHour();

}

/*
Auto fill in production goal base on material & machine
*/
function setGoal() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Daily Production Rating');
  //sheet.getRange(3,3).setValue(data[62][8].toLowerCase().replace(/\s/g,'').replace(/-/g,'').replace(/\//g,''));

  /*
      case '***':
        if(data[i-1][2] == '***'){
          sheet.getRange(i,8).setValue(oneWorkerData(*****, i));
        }
        break;
        */


  for(var i = 2; i<=data.length;i++){

    switch(data[i-1][8].toLowerCase().replace(/\s/g,'').replace(/-/g,'').replace(/\//g,'')){

        case 'pete110ag':
        if(data[i-1][2] == 'G4'){
          sheet.getRange(i,8).setValue(oneWorkerData(PETE_110_AG_G4_1W, i));
        }
        break;

        case 'pp506injmc':
        if(data[i-1][2] == 'G2'){
          sheet.getRange(i,8).setValue(oneWorkerData(PP_506_INJ_MC_G2_1W, i));
        }
        break;


        case 'hdpe205wfmc':
        if(data[i-1][2] == 'S1/G3'){
          sheet.getRange(i,8).setValue(oneWorkerData(HDPE_205_WF_MC_S1G3_1W, i));
        }
        break;

        case 'ldpe403extprgmc':
        if(data[i-1][2] == 'G2'){
          sheet.getRange(i,8).setValue(oneWorkerData(LDPE_403_EXT_PRG_MC_G2_1W, i));
        }else if(data[i-1][2] == 'S1/G3'){
          sheet.getRange(i,8).setValue(oneWorkerData(LDPE_403_EXT_PRG_MC_S1G3_1W, i));
        }else if(data[i-1][2] == 'S1'){
          sheet.getRange(i,8).setValue(oneWorkerData(LDPE_403_EXT_PRG_MC_S1_1W, i));
        }
        break;

      case 'hdpe204bmmc':
        if(data[i-1][2] == 'G2'){
          sheet.getRange(i,8).setValue(oneWorkerData(HDPE_204_BM_MC_G2_1W, i));
        }else if(data[i-1][2] == 'G1'){
          sheet.getRange(i,8).setValue(oneWorkerData(HDPE_204_BM_MC_G1_1W, i));
        }
        break;

      case 'multiflexa6221':
        if(data[i-1][2] == 'G1'){
          sheet.getRange(i,8).setValue(oneWorkerData(multiflexa6221_G1_1W, i));
        }
        break;

      case 'pcabs805injnat':
        if(data[i-1][2] == 'G1'){
          sheet.getRange(i,8).setValue(oneWorkerData(PCABS805INJNAT_G1_1W, i));
        }
        break;

      case 'pcabs806injmc':
      case 'pcabsmcgerhardi':
        if(data[i-1][2] == 'G2'||data[i-1][2] == 'G1'){
          sheet.getRange(i,8).setValue(oneWorkerData(PCABS806INJMC_G2_1W, i));
        }
        break;

      case 'abspc805injnat':
        if(data[i-1][2] == 'G2'){
          sheet.getRange(i,8).setValue(oneWorkerData(ABSPC805INJNAT_G2_1W, i));
        }
        break;

      case 'pp510extwt':
        if(data[i-1][2] == 'G1'){
          sheet.getRange(i,8).setValue(oneWorkerData(PP510EXTWT_G1_1W, i));
        }
        break;

       case 'ldpe402injcapwht':
        if(data[i-1][2] == 'G1'){
          sheet.getRange(i,8).setValue(oneWorkerData(LDPE402INJCAPWHT_G1_1W, i));
        }
        break;

      case 'hdpe202bmnat':
        if(data[i-1][2] == 'G1'){
          sheet.getRange(i,8).setValue(oneWorkerData(HDPE_202_BM_NAT_G2_1W, i));
        }
        break;

      case 'sonocoshred':
        if(data[i-1][2] == 'S2/G5'){
          sheet.getRange(i,8).setValue(twoWorkerData(sonocoShred_S2G5_2W, i));
        }
        break;

      case 'osabspmma':
        if(data[i-1][2] == 'S2/G5'){
          sheet.getRange(i,8).setValue(twoWorkerData(ABSOS_PMMA_S2G5_2W, i));
        }
        break;

      case 'acrylicwtape':
      case 'acrylicwredtape':
        if(data[i-1][2] == 'B1'){
          sheet.getRange(i,8).setValue(oneWorkerData(Acrylic_WithTape_B1_1W, i));
        }
        break;

        case 'acrylicchrome':
        if(data[i-1][2] == 'B2'){
          sheet.getRange(i,8).setValue(oneWorkerData(Acrylic_Chrome_B2_1W, i));
        }
        break;

      case 'pmma902imp':
      case 'pmma902impext':
        if(data[i-1][2] == 'G1'||data[i-1][2] == 'G2'){
          sheet.getRange(i,8).setValue(twoWorkerData(PMMA_902_G1_2W, i));
        }
        break;

      case 'abswcoating':
      case 'contaminategrayabs':
      case 'whiteabswcoating':
      case 'abswcoatingwhite':
        if(data[i-1][2] == 'B2'||data[i-1][2] == 'B1'){
          sheet.getRange(i,8).setValue(twoWorkerData(ABS_wCoating_B2_1W, i));
        }
        break;

      case 'pete101bmmc':
        if(data[i-1][2] == 'G4'){
          sheet.getRange(i,8).setValue(oneWorkerData(PET_101_G4_1W, i));
        }
        break;

      case 'pete104tray':
      case 'pete104traytn':
      case 'pete104traymc':
      if(data[i-1][2] == 'G4'){
        sheet.getRange(i,8).setValue(oneWorkerData(PET_104_Tray_G4_1W, i));
      }
      break;

      case 'cores':
        sheet.getRange(i,8).setValue(cores);
        break;

      case 'absoblk':
      case 'absoblack':
      case 'absblk':
      case 'osabs':
      case 'absosblk':
      case 'abschrome':
      case 'abscrome':
      case 'absosgray':
      case 'abs822':
      case 'absogray':
      case 'abs825extblkdm':
      case 'abs826extmcos':
      case'abs822&abs828':
      case 'abs828extpntos':

        if(data[i-1][2] == 'G2'){
          sheet.getRange(i,8).setValue(ABS_O_BLK_G2);
        }else if(data[i-1][2] == 'S2'){
          sheet.getRange(i,8).setValue(ABS_O_BLK_S2);
        }else if(data[i-1][2] == 'S1'){
          if(data[i-1][9] == 1){
            sheet.getRange(i,8).setValue(1200);
          }else{
            sheet.getRange(i,8).setValue(1600);
          }
        }else if(data[i-1][2] == 'S2/G5'){
          sheet.getRange(i,8).setValue(twoWorkerData(ABS_O_BLK_S2G5, i));
        }else if(data[i-1][2] == 'B1'){
              sheet.getRange(i,8).setValue(oneWorkerData(ABS_Chrome_B2_1W, i));
        }else if(data[i-1][2] == 'B2'){
          if(data[i-1][9] == 1){
            sheet.getRange(i,8).setValue(1200);
          }else{
            sheet.getRange(i,8).setValue(1600);
          }
        }
        break;

      case 'absgray':
      case 'grayabs':

        if(data[i-1][2] == 'S2'||data[i-1][2] == 'B2'){
          if(data[i-1][9] == 1){
            sheet.getRange(i,8).setValue(1200);
          }else{
            sheet.getRange(i,8).setValue(1600);
          }
        }

        break;

      case 'hd200':
      case 'hdpe200':
      case 'hd200shred':
        if(data[i-1][2] == 'G1'){
          sheet.getRange(i,8).setValue(oneWorkerData(HD200_G1_1W, i));
        }else if(data[i-1][2] == 'G2'){
          sheet.getRange(i,8).setValue(oneWorkerData(HD200_G2_1W, i));
        }else if(data[i-1][2] == 'S1'||data[i-1][2] == 'S2'){
          sheet.getRange(i,8).setValue(twoWorkerData(HD200_shred_s1s2_2W, i));
        }else if(data[i-1][2] == 'S1/G3'){
          sheet.getRange(i,8).setValue(oneWorkerData(HD200_S1G3_1W, i));
        }
        break;

      case 'hdpe207inj':
      case 'hdpe207':
      case 'hdpe207injnat':
        sheet.getRange(i,8).setValue(oneWorkerData(HDPE_207_INJ, i));
        break;

      case 'occ':
        sheet.getRange(i,8).setValue(oneWorkerData(OCC, i));
        break;

      case 'pete101bmclr':
      case 'pete101':
        sheet.getRange(i,8).setValue(PETE_101_BM_CLR);
        break;

      case 'pmma':
      case 'whitepmma':
      case 'pmmamix':
        sheet.getRange(i,8).setValue(twoWorkerData(PMMA, i));
          break;

      case 'ppcratesmc':
        sheet.getRange(i,8).setValue(PP_Craes_MC);
        break;

      case 'tpu':
        if(data[i-1][2] == 'B1'){
          sheet.getRange(i,8).setValue(TPU_B1);
        }else if(data[i-1][2] == 'B2'){
          sheet.getRange(i,8).setValue(TPU_B2);
        }
        break;

      case 'filmpepsi':
      case 'pepsifilm':
        sheet.getRange(i,8).setValue(oneWorkerData(Film_Pepsi, i));
        break;

      case 'abs822acoatwhtos':
      case 'abs822acoatingwhiteos':
      case 'abs822acoatblkos':
        sheet.getRange(i,8).setValue(ABS_822_AC_WHT_S2G5);
        break;

      case 'pete106':
      case 'pete106bmtn':
      case 'pete106bmtncl':
        sheet.getRange(i,8).setValue(PETE_106);
        break;

      case 'acryliccast':
      case 'acrylic':
        sheet.getRange(i,8).setValue(Acrylic_Cast);
        break;

      case 'petgreenstrapping':
        sheet.getRange(i,8).setValue(PET_Green_Strapping);
        break;

      case 'filmrebale':
      case 'quakerfilmrebale':
        sheet.getRange(i,8).setValue(Film_Rebale);
        break;

      case 'pvctubing':
        sheet.getRange(i,8).setValue(PVC_Tubing);
        break;

      case 'paintedabs':
        sheet.getRange(i,8).setValue(Painted_ABS);
        break;

      case 'abscontaminated':
        sheet.getRange(i,8).setValue(ABS_Contaminated);
        break;

      case 'nylon':
        sheet.getRange(i,8).setValue(Nylon);
        break;

      case 'tponibs':

          if(data[i-1][9] == 2){
            sheet.getRange(i,8).setValue(2500);
          }else{
            sheet.getRange(i,8).setValue(1500);
          }
        break;


      case 'absfenders':
      case 'absfender':
        if(data[i-1][2] == 'B1'){
          sheet.getRange(i,8).setValue(ABS_Fenders);
        }else{
          sheet.getRange(i,8).setValue(ABS_Fenders);
        }
        break;

      case 'pp502bmnat':
        sheet.getRange(i,8).setValue(PP_502_BM_Nat);
        break;

      case 'ppcr8mc' :
        sheet.getRange(i,8).setValue(PP_CR8_MC);
        break;

      case 'hips604':
        sheet.getRange(i,8).setValue(HIPS_604);
        break;

      case 'hdpe202':
      case 'hdpe202bmnat':
        sheet.getRange(i,1,1,15).setFontColor('gray');
        break;

      case 'pp511cr8mc':
        sheet.getRange(i,8).setValue(PP_511_cr8_MC);
        break;

      case 'hdpe203injmc':
        if(data[i-1][2] == 'S2/G5'){
        sheet.getRange(i,8).setValue(HPDE_203_INJ_MC_S2G5);
        }
        break;

      case 'pmma901imp' :
      case 'pmma901impinj':
      case 'pmmaimp':
        if(data[i-1][2] == 'G1'||data[i-1][2] == 'G2'){
          sheet.getRange(i,8).setValue(oneWorkerData(PMMA_901_IMP_G1G2_2W, i));
        }else if(data[i-1][2] == 'B1'){
          sheet.getRange(i,8).setValue(twoWorkerData(PMMA_IMP_B1_2W, i));
        }
        break;

      case 'ppinjmc':
        if(data[i-1][2] == 'S2/G5'){
          if(data[i-1][9] == 1){
            sheet.getRange(i,8).setValue(PP_INJ_MC_S2G5*0.7);
          }else{
            sheet.getRange(i,8).setValue(PP_INJ_MC_S2G5);
          }
        }
        break;

      case 'pmmawtape':
        sheet.getRange(i,8).setValue(PMMA_WTape_B1);
        break;

      case 'tpo3400':
        if(data[i-1][2] == 'S2/G5'){

            sheet.getRange(i,8).setValue(twoWorkerData(TPO_3400_S2G5_2W,i));

        }else if(data[i-1][2] == 'S1'){
          sheet.getRange(i,8).setValue(TPO_3400_S1);
        }
        break;

        case 'abspieces':
        if(data[i-1][2] == 'B1'){
          if(data[i-1][9] == 1){
            sheet.getRange(i,8).setValue(ABS_pieces_B1_1W);
          }else{
            sheet.getRange(i,8).setValue(ABS_pieces_B1_1W*1.3);
          }
        }
          break;


        case 'film':
          if(data[i-1][2] == 'B1'){
            if(data[i-1][9] == 2){
              sheet.getRange(i,8).setValue(Film_B1_2W);
            }else{
              sheet.getRange(i,8).setValue(Film_B1_2W*0.75);
            }
          }
            break;

      case 'tpocontaminated' :
      case 'contaminatedtpo':
        if(data[i-1][2] == 'B2'){
            if(data[i-1][9] == 2){
              sheet.getRange(i,8).setValue(TPO_Contaminated_B2_1W * 1.25);
            }else{
              sheet.getRange(i,8).setValue(TPO_Contaminated_B2_1W);
            }
        }
          break;

      case 'luzhd':
        if(data[i-1][2] == 'S1/G3'){
            if(data[i-1][9] == 2){
              sheet.getRange(i,8).setValue(LUZ_HD_S1G3_2W);
            }else{
              sheet.getRange(i,8).setValue(LUZ_HD_S1G3_2W * 0.75);
            }
        }
        break;


      case 'pp503injcapwt':
      case 'ppinj':
        if(data[i-1][2] == 'G1'){
          sheet.getRange(i,8).setValue(oneWorkerData(PP_503_INJ_CAPWT_G1_1W, i));
        }

        break;

      case 'hdpe203injmc':
      case 'hdpe203':
        if(data[i-1][2] == 'G1'){
          sheet.getRange(i,8).setValue(oneWorkerData(HDPE_203_INJ_MC_G1_1W, i));
        }else if(data[i-1][2] == 'G2'){
          sheet.getRange(i,8).setValue(oneWorkerData(HDPE_203_INJ_MC_G2_1W, i));
        }
        break;


      case 'ms905injcl':
        if(data[i-1][2] == 'G1'){
          sheet.getRange(i,8).setValue(oneWorkerData(MS_905_INJ_CL_G1_1W, i));
        }
        break;

      case 'pet102bmmc':
      case 'pete102':
      case 'pete102bmmc':
        if(data[i-1][2] == 'G4'){
          sheet.getRange(i,8).setValue(oneWorkerData(PET_102_BM_MC_G4_1W, i));
        }
        break;

      case 'pmma904injcl':
        if(data[i-1][2] == 'G2'){
          sheet.getRange(i,8).setValue(oneWorkerData(PMMA_904_INJ_CL_G2_1W, i));
        }
        break;


        default:
        break;

    }
    }

  data = dailyProductionRatingSheet.getDataRange().getValues();
  for(i = 2; i<=data.length;i++){
    //calculate data
    if(data[i-1][3] != ""){
      sheet.getRange(i,13).setValue(data[i-1][9]*data[i-1][10]*wage);
      sheet.getRange(i,14).setValue(data[i-1][12]/data[i-1][11]);
      sheet.getRange(i,7).setValue(data[i-1][11]/data[i-1][10]);
      if(data[i-1][7] != ""){
        sheet.getRange(i,15).setValue((data[i-1][7]*data[i-1][10]));
        sheet.getRange(i,5).setValue((data[i-1][11]/data[i-1][10])/data[i-1][7]);
        if(((data[i-1][11]/data[i-1][10])/data[i-1][7])>=1.1){
          sheet.getRange(i,6).setValue("A+");
        }else if(((data[i-1][11]/data[i-1][10])/data[i-1][7])>=0.9){
          sheet.getRange(i,6).setValue("A");
        }else if(((data[i-1][11]/data[i-1][10])/data[i-1][7])>=0.8){
          sheet.getRange(i,6).setValue("B");
        }else{
          sheet.getRange(i,6).setValue("C");
        }
      }
    }
}

}



/**
Change the font color base on the rating
**/
function changeColor(){
  var dataa = dailyProductionRatingSheet.getDataRange().getValues();
  for(var i = 1; i<dataa.length;i++){
    if(dataa[i][5] == 'A+'||dataa[i][5] == 'A'){
      dailyProductionRatingSheet.getRange(i+1,1,1,15).setFontColor('green');
    }else if(dataa[i][5] == 'B'){
      dailyProductionRatingSheet.getRange(i+1,1,1,15).setFontColor('blue');
    }else if(dataa[i][5] == 'C'){
      dailyProductionRatingSheet.getRange(i+1,1,1,15).setFontColor('red');
    }else{
      dailyProductionRatingSheet.getRange(i+1,1,1,15).setFontColor('black');
    }
  }

  dailyProductionRatingSheet.getRange(2,1,dataa.length,1).setFontColor('white');
}
/**
Import total data to Proudction Summary
**/
function makeProductionSummary(){
  var positionCounter = 2 ;
  var sum = 0;
  var average = 0;
  var idealSum = 0;
  var idealAverage = 0;
  var dataa = dailyProductionRatingSheet.getDataRange().getValues();

  //loop through data in daily production rating, if value of the column is total, then import to proudction summary
  for(var i = 1; i<dataa.length;i++){
    if(dataa[i][2] == 'Total'||dataa[i][2] == 'total'){
      productionSummarySheet.getRange(positionCounter,1).setValue(dataa[i][1]);
      productionSummarySheet.getRange(positionCounter,2).setValue(dataa[i][11]);
      productionSummarySheet.getRange(positionCounter,3).setValue(dataa[i][14]);
      positionCounter ++ ;
    }
  }

  for(var j = 2; j<=24; j++){
    sum = productionSummarySheet.getRange(j,2).getValue() + sum;
    idealSum = productionSummarySheet.getRange(j,3).getValue() + idealSum;
  }

  average = sum/(positionCounter - 2);
  idealAverage = idealSum/(positionCounter - 2);

  for(var k = 2; k<=24; k++){
    productionSummarySheet.getRange(k,4).setValue(average);
    productionSummarySheet.getRange(k,5).setValue(70000);
  }

  productionSummarySheet.getRange(25,2).setValue(sum);
  productionSummarySheet.getRange(26,2).setValue(average);
  productionSummarySheet.getRange(25,3).setValue(idealSum);
  productionSummarySheet.getRange(26,3).setValue(idealAverage);

  //change font color
  dataa = productionSummarySheet.getDataRange().getValues();
  for(i = 1; i<24;i++){
    if(dataa[i][1]> 70000){
      productionSummarySheet.getRange(i+1,1,1,3).setFontColor('green');
    }else{
      productionSummarySheet.getRange(i+1,1,1,3).setFontColor('red');
    }
  }

}


function calculateGrossProfit(){
  for(var i = 2; i<=data.length;i++){
    switch(data[i-1][8].toLowerCase().replace(/\s/g,'').replace(/-/g,'').replace(/\//g,'')){
      case 'pmma':
        dailyProductionRatingSheet.getRange(i,15).setValue(0.12);
        dailyProductionRatingSheet.getRange(i,16).setValue(0.12*data[i-1][11]);
        break;

      case 'hd200':
      case 'hdpe200':
        if(data[i-1][2] == 'S1/G3'){
          dailyProductionRatingSheet.getRange(i,15).setValue(0.08);
          dailyProductionRatingSheet.getRange(i,16).setValue(0.08*data[i-1][11]);
        }else if(data[i-1][2] == 'S1'||data[i-1][2] == 'G1'||data[i-1][2] == 'G2'){
          dailyProductionRatingSheet.getRange(i,15).setValue(0.04);
          dailyProductionRatingSheet.getRange(i,16).setValue(0.04*data[i-1][11]);
        }else if(data[i-1][2] == 'G3'){
          dailyProductionRatingSheet.getRange(i,15).setValue(0.08);
          dailyProductionRatingSheet.getRange(i,16).setValue(0.08*data[i-1][11]);
        }
        break;

      case 'hips604':
        if(data[i-1][2] == 'S2/G5'){
          dailyProductionRatingSheet.getRange(i,15).setValue(0.15);
          dailyProductionRatingSheet.getRange(i,16).setValue(0.15*data[i-1][11]);
        }
        break;

      case 'pete101bmclr':
        if(data[i-1][2] == 'G4'){
          dailyProductionRatingSheet.getRange(i,15).setValue(0.25);
          dailyProductionRatingSheet.getRange(i,16).setValue(0.25*data[i-1][11]);
        }
        break;

      case 'filmpepsi':
        if(data[i-1][2] == 'B1'){
          dailyProductionRatingSheet.getRange(i,15).setValue(0.05);
          dailyProductionRatingSheet.getRange(i,16).setValue(0.05*data[i-1][11]);
        }
        break;

      case 'abs822acoatingwhite':
        if(data[i-1][2] == 'S2/G5'){
          dailyProductionRatingSheet.getRange(i,15).setValue(0.2);
          dailyProductionRatingSheet.getRange(i,16).setValue(0.2*data[i-1][11]);
        }

        break;

      case 'pete106bmtn' :
      case 'pete106':
      case 'pete106bmtncl' :
        if(data[i-1][2] == 'G4'){
          dailyProductionRatingSheet.getRange(i,15).setValue(0.13);
          dailyProductionRatingSheet.getRange(i,16).setValue(0.13*data[i-1][11]);
        }
        break;

      case 'acryliccast' :
        if(data[i-1][2] == 'B1'){
          dailyProductionRatingSheet.getRange(i,15).setValue(0.08);
          dailyProductionRatingSheet.getRange(i,16).setValue(0.08*data[i-1][11]);
        }
        break;

      case 'abs':
        if(data[i-1][2] == 'S2'){
          dailyProductionRatingSheet.getRange(i,15).setValue(0.15);
          dailyProductionRatingSheet.getRange(i,16).setValue(0.15*data[i-1][11]);
        }
        break;



      default:
        break;
    }
}
}


function makePersonalRating(){
  var ratingDataSS = SpreadsheetApp.getActive().getSheetByName('Rating Data');
  var ratingSummarySS = SpreadsheetApp.getActive().getSheetByName('Rating Summary');
  var employeeNames = ["brayan", "ismael", "duglas", "saul", "mario", "carlos", "maurilia", "marisol"];
  var rowNumber;
  var tempName;
  var columnNumber = -2;
  var total;
  var counter;
  var average;
  var summaryRowNumber = 2;
  for (var j = 0; j < employeeNames.length; j++){
    rowNumber = 2;
    columnNumber = columnNumber + 3;
    total = 0;
    counter = 0;
    for(var i = 1; i< data.length;i++){
      tempName = data[i][3].toLowerCase();
      if(tempName.indexOf(employeeNames[j]) > -1 && data[i][5] != ""){
        ratingDataSS.getRange(rowNumber, columnNumber).setValue(data[i][1]);
        ratingDataSS.getRange(rowNumber, columnNumber+1).setValue(data[i][5]);
        ratingDataSS.getRange(rowNumber, columnNumber+2).setValue(data[i][4]).setNumberFormat("0.0;%");
        if(data[i][4]!=""){
          total = total + data[i][4];
          counter++;
        }
        rowNumber++;
      }
    }
    average = total/counter;
    ratingDataSS.getRange(60, columnNumber).setValue(capitalizeFirstLetter(employeeNames[j]));
    ratingDataSS.getRange(60, columnNumber + 1).setValue(average);
    ratingSummarySS.getRange(summaryRowNumber, 1).setValue(capitalizeFirstLetter(employeeNames[j]));
    ratingSummarySS.getRange(summaryRowNumber, 2).setValue(average);
    summaryRowNumber++;
  }
}

function makeMachineAnalysis(){
  var machineAnalysisSS = SpreadsheetApp.getActive().getSheetByName('Machine Analysis');
  var machineNames =['s1/g3', 's2/g5', 's1', 'g1', 'g2', 'g4', 'b1', 'b2'];
  var materials = [];
  //variables for calcualting average production pcercentage per material on each machine base on number of workers
  var total_1W = 0;
  var counter_1W = 0;
  var total_2W = 0;
  var counter_2W = 0;
  var rowNumber = 2;
  var hasMachine = false;

  //Variables for calcualting average production numbers per material on each machine base on number of workers
  var materialTotal_1W = 0;
  var materialCounter_1W = 0;
  var materialTotal_2W = 0;
  var materialCounter_2W = 0;

  for (var j = 0; j < machineNames.length; j++){
    for(var i = 1; i < data.length;i++){
      if(data[i][2].toLowerCase().indexOf(machineNames[j]) > -1){
        hasMachine = true;
        materials.push(data[i][8].replace(/\s/g,'').replace(/-/g,'').toLowerCase());
      }
    }
    if(hasMachine){

    materials = removeDups(materials);

    for(k = 0; k< materials.length; k++){
      for(i = 1; i< data.length;i++){
        if(data[i][2].toLowerCase().indexOf(machineNames[j]) > -1
           && data[i][8].toLowerCase().replace(/\s/g,'').replace(/-/g,'').indexOf(materials[k].toLowerCase()) > -1){

          //get total and times of each material base one number of workers
          if(data[i][9] == 1){
            materialTotal_1W = materialTotal_1W + data[i][6];
            materialCounter_1W++;
            total_1W = total_1W + data[i][4];
            counter_1W = counter_1W + 1;
          }else{
            total_2W = total_2W + data[i][4];
            counter_2W = counter_2W + 1;
            materialTotal_2W = materialTotal_2W + data[i][6];
            materialCounter_2W++;
          }

          hasMachine = true;
        }
      }
      machineAnalysisSS.getRange(rowNumber, 1).setValue(machineNames[j].toUpperCase());
      machineAnalysisSS.getRange(rowNumber, 2).setValue(materials[k].toUpperCase());
      //put in percentage
      if(total_1W != 0){
        machineAnalysisSS.getRange(rowNumber, 3).setValue(total_1W/counter_1W);
      }else{
        machineAnalysisSS.getRange(rowNumber, 3).setValue("NA");
      }
      if(total_2W != 0){
        machineAnalysisSS.getRange(rowNumber, 4).setValue(total_2W/counter_2W);
      }else{
        machineAnalysisSS.getRange(rowNumber, 4).setValue("NA");
      }
      //put in numbers
      if(materialTotal_1W != 0){
        machineAnalysisSS.getRange(rowNumber, 5).setValue(materialTotal_1W/materialCounter_1W);
      }else{
        machineAnalysisSS.getRange(rowNumber, 5).setValue("NA");
      }
      if(materialTotal_2W != 0){
        machineAnalysisSS.getRange(rowNumber, 6).setValue(materialTotal_2W/materialCounter_2W);
      }else{
        machineAnalysisSS.getRange(rowNumber, 6).setValue("NA");
      }

      total_1W = 0;
      total_2W = 0;
      counter_1W = 0;
      counter_2W = 0;
      materialTotal_1W = 0;
      materialCounter_1W = 0;
      materialTotal_2W = 0;
      materialCounter_2W = 0;
      rowNumber ++;
      machineAnalysisSS.getRange(rowNumber, 1).setValue('');
      machineAnalysisSS.getRange(rowNumber, 2).setValue('');
      machineAnalysisSS.getRange(rowNumber, 3).setValue('');
      machineAnalysisSS.getRange(rowNumber, 4).setValue('');
    }
    materials = [];
    rowNumber++;
  }
    hasMachine = false;
  }
  //change color
  var machineData = machineAnalysisSS.getDataRange().getValues();
  for(var i = 1; i<machineData.length;i++){
    if(machineData[i][2] >= 0.9){
      machineAnalysisSS.getRange(i+1,3,1,1).setFontColor('green');
    }else if(machineData[i][2] >= 0.8){
      machineAnalysisSS.getRange(i+1,3,1,1).setFontColor('blue');
    }else if(machineData[i][2] <0.8){
      machineAnalysisSS.getRange(i+1,3,1,1).setFontColor('red');
    }else{
      machineAnalysisSS.getRange(i+1,3,1,1).setFontColor('gray');
    }

    if(machineData[i][3] >= 0.9){
      machineAnalysisSS.getRange(i+1,4,1,1).setFontColor('green');
    }else if(machineData[i][3] >= 0.8){
      machineAnalysisSS.getRange(i+1,4,1,1).setFontColor('blue');
    }else if(machineData[i][3] <0.8){
      machineAnalysisSS.getRange(i+1,4,1,1).setFontColor('red');
    }else{
      machineAnalysisSS.getRange(i+1,4,1,1).setFontColor('gray');
    }
  }
}

function makePersonalPerformanceAnalysis(){
  var personalPerformanceAnalysisSS = SpreadsheetApp.getActive().getSheetByName('Personal Performance Analysis');
  var machineNames =['s1/g3', 's2/g5', 'g1', 'g2', 'g4', 'b1', 'b2'];
  var materials = [];
  var person = [];
  var total = 0;
  var counter = 0;
  var rowNumber = 2;
  var hasMachine = false;

  for (var j = 0; j < machineNames.length; j++){
    hasMachine = false;
    for(var i = 1; i < data.length;i++){
      if(data[i][2].toLowerCase().indexOf(machineNames[j]) > -1){
        hasMachine = true;
        materials.push(data[i][8].replace(/\s/g,'').toLowerCase());
      }
    }
    if(!hasMachine){
     continue;
    }
    materials = removeDups(materials);

    for(k = 0; k< materials.length; k++){
      //find person of this machine -> material
      for(i = 1; i< data.length;i++){
        if(data[i][2].toLowerCase().indexOf(machineNames[j]) > -1
           && data[i][8].toLowerCase().replace(/\s/g,'').indexOf(materials[k].toLowerCase()) > -1){
          person.push(data[i][3].toLowerCase());
        }
      }
      person = removeDups(person);
      for(var q = 0; q < person.length; q++){
        for(i = 1; i< data.length;i++){
        if(data[i][2].toLowerCase().indexOf(machineNames[j]) > -1
           && data[i][8].toLowerCase().replace(/\s/g,'').indexOf(materials[k].toLowerCase()) > -1
           && data[i][3].toString().toLowerCase() == person[q].toLowerCase()){
            total = total + data[i][4];
            counter = counter + 1;
        }
      }
      personalPerformanceAnalysisSS.getRange(rowNumber, 1).setValue(machineNames[j].toUpperCase());
      personalPerformanceAnalysisSS.getRange(rowNumber, 2).setValue(materials[k].toUpperCase());
      personalPerformanceAnalysisSS.getRange(rowNumber, 3).setValue(capitalizeFirstLetter(person[q]));
      personalPerformanceAnalysisSS.getRange(rowNumber, 4).setValue(total/counter);
      total = 0 ;
      counter = 0;
      rowNumber ++;
      personalPerformanceAnalysisSS.getRange(rowNumber, 1).setValue('');
      personalPerformanceAnalysisSS.getRange(rowNumber, 2).setValue('');
      personalPerformanceAnalysisSS.getRange(rowNumber, 3).setValue('');
      personalPerformanceAnalysisSS.getRange(rowNumber, 4).setValue('');
      }
      person = [];
      rowNumber++;

    }
    materials = [];
    rowNumber++;
}
  //change color
  var machineData = personalPerformanceAnalysisSS.getDataRange().getValues();
  for(var i = 1; i<machineData.length;i++){
    if(machineData[i][3] >= 0.9){
      personalPerformanceAnalysisSS.getRange(i+1,1,1,15).setFontColor('green');
    }else if(machineData[i][3] >= 0.8){
      personalPerformanceAnalysisSS.getRange(i+1,1,1,15).setFontColor('blue');
    }else{
      personalPerformanceAnalysisSS.getRange(i+1,1,1,15).setFontColor('red');
    }
  }
}


function makeWorkerHour(){
  var workHourSS = SpreadsheetApp.getActive().getSheetByName('WorkHour');
  var startPosition = 2;
  var names = [];
  var tempNames = [];
  var time = 0;
  var dailyTimeTotal = 0;
  var missTimeTotal = 0;
  var row = 2;
  var monthTimeTotal = 0;
  var monthMissTimeTotal = 0;


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







function oneWorkerData( goal, i){
  if(data[i-1][9] == 2){
    goal = goal *1.3;
  }
  return(goal);
}

function twoWorkerData( goal, i){
  if(data[i-1][9] == 1){
    goal = goal *0.7;
  }
  return(goal);
}

function capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
}

function removeDups(array) {
  var outArray = [];
  array.sort();
  outArray.push(array[0]);
  for(var n in array){
    if(outArray[outArray.length-1]!=array[n].trim()){
      outArray.push(array[n].trim());
    }
  }
  return outArray;
}
