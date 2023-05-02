function perWeek() {

    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2023').activate();
    const mysheet = SpreadsheetApp.getActiveSheet();
  
    //2日前の日付の取得
    let month,date;
    const today     = new Date();
    const yesterday = new Date(today); 
    yesterday.setDate(today.getDate() - 2);
  
    month       = yesterday.getMonth() + 1;
    date        = yesterday.getDate();
    let date_value = month * 100 + date;
    console.log(`${month}月${date}日`);
  
    let i = mysheet.getLastRow();
    let row_week;
  
    while(i => 3){
      if(mysheet.getRange(i,1).getValue() == date_value){
        row_week = i;
        break;
      }
      i--;
    }
    console.log(`${i}列目以降`);
  
    //週当たりの収支の計算(毎週月曜朝に実行)
  
    const limit_of_week = mysheet.getRange(1,8).getValue();
    let   cost_of_week  = 0;
    let   cost_array    = [0,0,0,0,0,0,0];
  
    let day_frag = 0;//曜日が変わる回数をカウントする
  
    i = row_week;
  
    while(day_frag < 7 && i > 2){
  
      if(mysheet.getRange(i,3).getValue() != "")cost_of_week += mysheet.getRange(i,3).getValue();
  
      if(mysheet.getRange(i,4).getValue() == '食費（外食or中食）'       )cost_array[0] += mysheet.getRange(i,3).getValue();
      if(mysheet.getRange(i,4).getValue() == '食費（買いだめ）'         )cost_array[1] += mysheet.getRange(i,3).getValue();
      if(mysheet.getRange(i,4).getValue() == '生活用品（消耗品）'       )cost_array[2] += mysheet.getRange(i,3).getValue();
      if(mysheet.getRange(i,4).getValue() == '生活用品（消耗品でない）'  )cost_array[3] += mysheet.getRange(i,3).getValue();
      if(mysheet.getRange(i,4).getValue() == '交通費'                 )cost_array[4] += mysheet.getRange(i,3).getValue();
      if(mysheet.getRange(i,4).getValue() == 'な'                    )cost_array[5] += mysheet.getRange(i,3).getValue();
      if(mysheet.getRange(i,4).getValue() == 'その他'                 )cost_array[6] += mysheet.getRange(i,3).getValue();
  
      if(mysheet.getRange(i,2).getValue() != mysheet.getRange(i-1,2).getValue()) day_frag++;
      i--;
    }
  
    mysheet.insertRowAfter(row_week);//一行挿入
    mysheet.getRange(row_week+1,1).setValue(date_value);
    mysheet.getRange(row_week+1,2).setValue(yesterday.getDay()); 
  
    mysheet.getRange(row_week+1,7).setValue(cost_of_week);
    mysheet.getRange(row_week+1,8).setValue(limit_of_week - cost_of_week);
    mysheet.getRange(row_week+1,1,1,8).setBackground("aqua");
    if((limit_of_week - cost_of_week) < 0){
      mysheet.getRange(row_week+1,8).setFontColor('red');
    }
    else{
      mysheet.getRange(row_week+1,8).setFontColor('black');
    }
  
    //メール送信
    const recipient = ''; //gmailのアドレス
    const subject = '【from SpreadSheet】先週の支出';
  
    let body ='先週の支出:'+String(cost_of_week)+'\n'
            +'目安額との差:'+String(limit_of_week - cost_of_week)+'\n'
            +'食費：'+String(cost_array[0] + cost_array[1])+'\n\n'
            +'内訳は\n'+'---------------------\n'
            +'食費（外食or中食）:'+String(cost_array[0])+'\n'
            +'食費（買いだめ）:'+String(cost_array[1])+'\n'
            +'生活用品（消耗品）:'+String(cost_array[2])+'\n'
            +'生活用品（消耗品でない）:'+String(cost_array[3])+'\n'
            +'交通費:'+String(cost_array[4])+'\n'
            +'な:'+String(cost_array[5])+'\n'
            +'その他:'+String(cost_array[6])+'\n'
            +'---------------------\n';
            
    const options = {name:'支出計算シート'};
  
    GmailApp.sendEmail(recipient,subject,body,options);
    
  }
  