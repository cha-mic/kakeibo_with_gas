function perDay(){

    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2023').activate();
    const mysheet = SpreadsheetApp.getActiveSheet();
  
    //２日前の日付の取得
    const today     = new Date();
    const yesterday = new Date(today); 
    yesterday.setDate(today.getDate() - 2);
  
    let month       = yesterday.getMonth()+1;
    let date        = yesterday.getDate();
    let date_value = month*100+date;
    console.log(`${month}月${date}日`);
  
    //データを日付でソート
    let lastrow = mysheet.getLastRow();
    let i = 3;
    while( mysheet.getRange(i,1).getValue() < date_value - 1){
      i++;
    }
    console.log(`${i}列目から${lastrow}列目をソート`);
    let data = mysheet.getRange(i,1,lastrow - 2,4);
    data.sort(1); 
    console.log(`ソート完了`)
  
    //一日あたりの収支の計算（食費）
    const limit_of_day = mysheet.getRange(1,6).getValue();
    let cost_of_day = 0;
    let row_day;
    i = 3;
  
    let str;
    let term = /外食/;
  
    while( mysheet.getRange(i,1).getValue() < date_value + 1 ){
  
      if( mysheet.getRange(i,1).isBlank() == true ) break; //空白判定
  
      str = mysheet.getRange(i,4).getValue();
  
      if( mysheet.getRange(i,1).getValue() == date_value ){
        if( term.test(str) == true )cost_of_day += mysheet.getRange(i,3).getValue();
      }
  
      row_day = i;
      i++;
    }
    console.log(cost_of_day);
  
    mysheet.insertRowAfter(row_day);//一行挿入
    mysheet.getRange(row_day + 1,1).setValue(date_value);
    mysheet.getRange(row_day + 1,2).setValue(yesterday.getDay()); 
   
    mysheet.getRange(row_day + 1,5).setValue(cost_of_day);
    mysheet.getRange(row_day + 1,6).setValue(limit_of_day - cost_of_day);
    mysheet.getRange(row_day + 1,1,1,6).setBackground("lime");
    if( (limit_of_day - cost_of_day) < 0){
      mysheet.getRange(row_day+1,6).setFontColor('red');
    }
    else{
      mysheet.getRange(row_day+1,6).setFontColor('black');
    }
  
  }
  