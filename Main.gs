var queryString = Math.random();
var ss = SpreadsheetApp.getActiveSpreadsheet();

// Get Sheet
var ssWalletsSheet = ss.getSheetByName('Wallets');
var ssTotalSheet = ss.getSheetByName('Total');
var ssBuysSheet = ss.getSheetByName('Buys');
var ssLoanSheet = ss.getSheetByName('Loan');
var targetCurrency = 'usd'

function getTotalData() {
  
  // Grabs all CoinMarketCap data
  if (typeof targetCurrency == 'undefined' || targetCurrency == '') {targetCurrency = 'usd'};
  //try {
  var coins = getCoins();
  //} catch(e if e instanceof SyntaxError) {
  //  // Most likely being rate limited by CMC
  //  console.error('CMC sent back html about being rate limited. Try again later. Error: ' + e);
  //
  //  // Kill function
  //  return; 
  //} catch(e) {
  //  // Logs an unknown or new ERROR message.
  //  console.error('There was an error in getTotalData(): ' + e);
  //}
  
  // Make sure to get sheet
  ssWalletSheet = ss.getSheetByName('Wallets');
  ssTotalSheet = ss.getSheetByName('Total');
  //ssTESTSheet = ss.getSheetByName('TEST');
  
  var myCoins = ssWalletSheet.getRange("A8:C56").getValues();
  var buyPrice = ssBuysSheet.getRange("C2:C55").getValues();
  
  var totRangeAC = ssTotalSheet.getRange("A2:C52"); //A-C
  var totRangeGL = ssTotalSheet.getRange("G2:L52"); //G-L
  var TotOutputArrayAC = totRangeAC.getValues();  //New output array for performance purposes according to google apps scripts best practices to use setvalues as few as possible
  var TotOutputArrayGL = totRangeGL.getValues();
  //values[6][3] = "This is D7"; A0 B1 C2 D3 E4 F5 G6 H7 I8 J9 K10 L11 M12... because 7 row = 6 and 3 col = 4 = D
  
  // Creating new Object with our coins for later use.  
  // Each Object's key is the coin symbol
  var myCoinsObj = {};
  var myCoinsCount = myCoins.length;
  var n = 0;
  for (var i = 0; i < myCoinsCount; i++) {
    if (myCoins[i][0] !== '') {
      
      try {    
        n = 0;  
      
        while (coins['data'][n]['symbol'] !== myCoins[i][0]) {
          n++;
        }
        myCoinsObj[coins['data'][n]['symbol']] = coins['data'][n];
              
      
        ssWalletSheet.getRange('B'+(i+8).toString()).setValue(myCoinsObj[myCoins[i][0]]['name']);
      
        //ssTotalSheet.getRange('A'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['name']);
        if (myCoinsObj[myCoins[i][0]]['name'] == 0) {
          TotOutputArrayAC[i][0] = "Error? - " & myCoins[i][0]['name'];
        } else {
          TotOutputArrayAC[i][0] = myCoinsObj[myCoins[i][0]]['name'];
        }
        //ssTotalSheet.getRange('B'+(i+2).toString()).setValue(myCoins[i][2]);
        TotOutputArrayAC[i][1] = myCoins[i][2];
        //ssTotalSheet.getRange('C'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['quote']['USD']['price']);
        TotOutputArrayAC[i][2] = myCoinsObj[myCoins[i][0]]['quote']['USD']['price'];
    
        //ssTotalSheet.getRange('G'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['cmc_rank']);
        TotOutputArrayGL[i][6-6] = myCoinsObj[myCoins[i][0]]['cmc_rank'];
        //ssTotalSheet.getRange('H'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['quote']['USD']['percent_change_1h']);
        TotOutputArrayGL[i][7-6] = myCoinsObj[myCoins[i][0]]['quote']['USD']['percent_change_1h'];
        //ssTotalSheet.getRange('I'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['quote']['USD']['percent_change_24h']);
        TotOutputArrayGL[i][8-6] = myCoinsObj[myCoins[i][0]]['quote']['USD']['percent_change_24h'];
        //ssTotalSheet.getRange('J'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['quote']['USD']['percent_change_7d']);      
      
        //ssTotalSheet.getRange('J'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['symbol']);
        TotOutputArrayGL[i][9-6] = myCoinsObj[myCoins[i][0]]['symbol'];
        //ssTotalSheet.getRange('K'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['price_btc']);                                                  FIX?
        //ssTotalSheet.getRange('L'+(i+2).toString()).setValue(buyPrice[i][0]);
        TotOutputArrayGL[i][11-6] = buyPrice[i][0];
      
        //ssTotalSheet.getRange(''+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['id']);
        //ssTotalSheet.getRange(''+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['24h_volume_usd']);
        //ssTotalSheet.getRange(''+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['market_cap_usd']);
        //ssTotalSheet.getRange(''+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['available_supply']);
        //ssTotalSheet.getRange(''+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['total_supply']);
        //ssTotalSheet.getRange(''+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['max_supply']);
      
        //ssTotalSheet.getRange(''+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['last_updated']);
      
        //if (typeof targetCurrency !== 'usd') {
        //ssTotalSheet.getRange(''+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['price_' + targetCurrency]);
        //ssTotalSheet.getRange(''+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['24h_volume_' + targetCurrency]);
        //ssTotalSheet.getRange(''+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['market_cap_' + targetCurrency]);
        //};
      
      
        //Also do the Loan Sheet
        ssLoanSheet.getRange('A'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['name']);
        //ssLoanSheet.getRange('B'+(i+2).toString()).setValue(myCoins[i][2]);
        ssLoanSheet.getRange('C'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['quote']['USD']['price']);
    
        ssLoanSheet.getRange('G'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['cmc_rank']);
        ssLoanSheet.getRange('H'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['quote']['USD']['percent_change_1h']);
        ssLoanSheet.getRange('I'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['quote']['USD']['percent_change_24h']);
        ssLoanSheet.getRange('J'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['quote']['USD']['percent_change_7d']);      
      
        ssLoanSheet.getRange('K'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['symbol']);
        //ssLoanSheet.getRange('L'+(i+2).toString()).setValue(myCoinsObj[myCoins[i][0]]['price_btc']);                                                  FIX?
        //ssLoanSheet.getRange('M'+(i+2).toString()).setValue(buyPrice[i][0]);
        
      } catch(e if e instanceof TypeError) {
        // Most likely a coin symbol was changed. Need to output which one was changed to make easier to fix
        console.error('Symbol ' + myCoins[i][0] + ' was probably changed on CMC. Error: ' + e);
        
        //ssTotalSheet.getRange('A'+(i+2).toString()).setValue("");
        if (myCoins[i][0] !== "") {
          TotOutputArrayAC[i][0] = "Error? - " + myCoins[i][0]; //concat is + sign
        } else {
          TotOutputArrayAC[i][0] = "";
        }
        //ssTotalSheet.getRange('B'+(i+2).toString()).setValue("");
        TotOutputArrayAC[i][1] = "";
        //ssTotalSheet.getRange('C'+(i+2).toString()).setValue("");
        TotOutputArrayAC[i][2] = "";
    
        //ssTotalSheet.getRange('G'+(i+2).toString()).setValue("");
        TotOutputArrayGL[i][6-6] = "";
        //ssTotalSheet.getRange('H'+(i+2).toString()).setValue("");
        TotOutputArrayGL[i][7-6] = "";
        //ssTotalSheet.getRange('I'+(i+2).toString()).setValue("");   
        TotOutputArrayGL[i][8-6] = "";
      
        //ssTotalSheet.getRange('J'+(i+2).toString()).setValue("");
        TotOutputArrayGL[i][9-6] = "";
        //ssTotalSheet.getRange('L'+(i+2).toString()).setValue(""); 
        TotOutputArrayGL[i][11-6] = "";
        
      } catch(e) {
        // Logs an unknown or new ERROR message.
        console.error('There was an error in getTotalData(): ' + e);
      }
      
    } else {
      //ssTotalSheet.getRange('A'+(i+2).toString()).setValue("");
      if (myCoins[i][0] !== "") {
        TotOutputArrayAC[i][0] = "Error? - " + myCoins[i][0];
      } else {
        TotOutputArrayAC[i][0] = "";
      }
      //ssTotalSheet.getRange('B'+(i+2).toString()).setValue("");
      TotOutputArrayAC[i][1] = "";
      //ssTotalSheet.getRange('C'+(i+2).toString()).setValue("");
      TotOutputArrayAC[i][2] = "";
    
      //ssTotalSheet.getRange('G'+(i+2).toString()).setValue("");
      TotOutputArrayGL[i][6-6] = "";
      //ssTotalSheet.getRange('H'+(i+2).toString()).setValue("");
      TotOutputArrayGL[i][7-6] = "";
      //ssTotalSheet.getRange('I'+(i+2).toString()).setValue("");   
      TotOutputArrayGL[i][8-6] = "";
      
      //ssTotalSheet.getRange('J'+(i+2).toString()).setValue("");
      TotOutputArrayGL[i][9-6] = "";
      //ssTotalSheet.getRange('L'+(i+2).toString()).setValue(""); 
      TotOutputArrayGL[i][11-6] = "";
    }
  }
  
  //Output the array to the sheet
  totRangeAC.setValues(TotOutputArrayAC);
  totRangeGL.setValues(TotOutputArrayGL);
  //ssTESTSheet.getRange("A1:C52").setValues(TotOutputArrayAC);
  //ssTESTSheet.getRange("G1:L52").setValues(TotOutputArrayGL);
    
  //Add total to 'Daily Balance' list
  var ssTimelineSheet = ss.getSheetByName('Timeline');
  
  var column = ssTimelineSheet.getRange('G:G');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct] && values[ct][0] != "" ) {
    ct++;
  }
  
  var _MS_PER_DAY = 1000 * 60 * 60 * (24-.5); //Every 23.5 hours instead of 24 to try and make the times line up better by day
  var timelineSheetCell = ssTimelineSheet.getRange('G' + (ct - 0));
  var today = new Date();
  var dateCell = timelineSheetCell.getValue();
  
  if ((today.valueOf() - dateCell.valueOf()) > _MS_PER_DAY) {
    ssTimelineSheet.getRange('G' + (ct + 1)).setValue(new Date()).setNumberFormat("MM/dd/YYYY");
    ssTimelineSheet.getRange('H' + (ct + 1)).setValue(ssTotalSheet.getRange("C53").getValue());
  }
  
  //Sort the Data
  SpreadsheetApp.flush();
  ssTotalSheet.getRange("A2:M52").sort({column: 5, ascending: false});

}

function getCoins() {

  var url = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest?limit=1500';
  var response = UrlFetchApp.fetch(url, {'headers': {'X-CMC_PRO_API_KEY' : 'INSERT API KEY HERE','Accept' : 'application/json', 'Accept-Encoding' : 'deflate, gzip'}, 'muteHttpExceptions': true});
  //ssTotalSheet.getRange('A'+(100).toString()).setValue(response);
  Logger.log(response);
  var json = response.getContentText();
  ssTotalSheet.getRange('A'+(100).toString()).setValue(json.substring(0, 50000));
  var data = JSON.parse(json);
    
  return data;
}

function HIDE() {
  var ssWhatIfSheet = ss.getSheetByName('What If');
  var ssTimelineSheet = ss.getSheetByName('Timeline');
  var ssCoinDescriptionsSheet = ss.getSheetByName('Coin Descriptions');
  
  var totalSheetValues = ssTotalSheet.getRange("D2:D52");
  var totalSheetValues2 = ssTotalSheet.getRange("B55:B57");
  var coinDescriptionsSheetValues = ssCoinDescriptionsSheet.getRange("B2:B100");
  var walletSheetValues = ssWalletsSheet.getRange("C6:AY7");
  var whatIfSheetValues = ssWhatIfSheet.getRange("D1:D10");
  var timelineSheetValues = ssTimelineSheet.getRange("E1:E900");
  var timelineSheetValues2 = ssTimelineSheet.getRange("H2:H900");
  
  ssWhatIfSheet.getRange("C9").setValue("1");
  ssTimelineSheet.getRange("F2").setValue("1");
  totalSheetValues.setFontColor("white");
  totalSheetValues2.setFontColor("white");
  ssTotalSheet.getRange("C53").setFontColor("white");
  totalSheetValues.setBackground("white");
  coinDescriptionsSheetValues.setFontColor("white");
  walletSheetValues.setFontColor("white");
  whatIfSheetValues.setFontColor("white");
  timelineSheetValues.setFontColor("white");
  timelineSheetValues2.setFontColor("white");
}

function UNHIDE() {
  var ssWhatIfSheet = ss.getSheetByName('What If');
  var ssTimelineSheet = ss.getSheetByName('Timeline');
  var ssCoinDescriptionsSheet = ss.getSheetByName('Coin Descriptions');
  
  var totalSheetValues = ssTotalSheet.getRange("D2:D52");
  var totalSheetValues2 = ssTotalSheet.getRange("B55:B57");
  var coinDescriptionsSheetValues = ssCoinDescriptionsSheet.getRange("B2:B100");
  var walletSheetValues = ssWalletsSheet.getRange("C6:AY7");
  var whatIfSheetValues = ssWhatIfSheet.getRange("D1:D10");
  var timelineSheetValues = ssTimelineSheet.getRange("E1:E900");
  var timelineSheetValues2 = ssTimelineSheet.getRange("H2:H900");
  
  ssWhatIfSheet.getRange("C9").setValue("");
  ssTimelineSheet.getRange("F2").setValue("");
  totalSheetValues.setFontColor("black");
  totalSheetValues2.setFontColor("black");
  ssTotalSheet.getRange("C53").setFontColor("black");
  totalSheetValues.clearFormat(); //clearing format sets it back to using alternating colors
  coinDescriptionsSheetValues.setFontColor("black");
  walletSheetValues.setFontColor("black");
  whatIfSheetValues.setFontColor("black");
  timelineSheetValues.setFontColor("black");
  timelineSheetValues2.setFontColor("black");
}
