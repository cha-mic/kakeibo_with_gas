function perMonth() {

    const year = '2023';
  
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(year).activate();
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
    let row_month = 0;
  
    while(i >= 3){
  
      if((mysheet.getRange(i,1).getValue() - (mysheet.getRange(i,1).getValue() % 100)) / 100 == month ){
        row_month = i;
        break;
      }
  
      i--;
    }
    console.log(`${i}列目以降`);
  
    //月あたりの支出の計算(毎月二日に計算)
  
    let cost_of_month = 0;
    let cost_array    = [0,0,0,0,0,0,0];
  
    let month_frag = 0;
  
    i = row_month;
    while( month_frag == 0 && i >= 3){
  
      if(mysheet.getRange(i,4).getValue() != "")cost_of_month += mysheet.getRange(i,3).getValue();
  
      if(mysheet.getRange(i,4).getValue() == '食費（外食or中食）'       )cost_array[0] += mysheet.getRange(i,3).getValue();
      if(mysheet.getRange(i,4).getValue() == '食費（買いだめ）'         )cost_array[1] += mysheet.getRange(i,3).getValue();
      if(mysheet.getRange(i,4).getValue() == '生活用品（消耗品）'       )cost_array[2] += mysheet.getRange(i,3).getValue();
      if(mysheet.getRange(i,4).getValue() == '生活用品（消耗品でない）'  )cost_array[3] += mysheet.getRange(i,3).getValue();
      if(mysheet.getRange(i,4).getValue() == '交通費'                 )cost_array[4] += mysheet.getRange(i,3).getValue();
      if(mysheet.getRange(i,4).getValue() == 'な'                    )cost_array[5] += mysheet.getRange(i,3).getValue();
      if(mysheet.getRange(i,4).getValue() == 'その他'                 )cost_array[6] += mysheet.getRange(i,3).getValue();
  
      if((mysheet.getRange(i,1).getValue() - (mysheet.getRange(i,1).getValue() % 100)) / 100
       != (mysheet.getRange(i-1,1).getValue() - (mysheet.getRange(i-1,1).getValue() % 100)) / 100){
         month_frag++;
       }
      i--;
  
    }
  
    mysheet.insertRowAfter(row_month);//一行挿入
    mysheet.getRange(row_month+1,1).setValue(date_value);
    mysheet.getRange(row_month+1,2).setValue(yesterday.getDay());
      
    mysheet.getRange(row_month+1,9).setValue(cost_of_month);
    
    mysheet.getRange(row_month+1,1,1,9).setBackground("yellow");
  
    sendEmail(cost_of_month, cost_array);
    makeSummary(year, date_value, cost_of_month, cost_array);
  
  }
  
  function sendEmail(cost_of_month, cost_array){
  
    //メール送信
    const recipient = ''; //gmailのアドレス
    const subject = '【from SpreadSheet】先月の支出';
  
    let body ='先月の支出:'+String(cost_of_month)+'\n'
            +'食費：'+String(cost_array[0]+cost_array[1])+'\n\n'
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
  
    console.log(`メール送信`);
  
  }
  
  function makeSummary(year, date_value, cost_of_month, cost_array){
  
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Summary').activate();
    const mysheet = SpreadsheetApp.getActiveSheet();
  
    let i = mysheet.getLastRow() + 1;
  
    mysheet.getRange(i, 1).setValue(Number(year));
  
    let month = (date_value - (date_value % 100)) / 100;
  
    mysheet.getRange(i, 2).setValue(month);
  
    let official = mysheet.getRange(i, 10).getValue() + mysheet.getRange(i, 11).getValue() 
    + mysheet.getRange(i, 12).getValue() + mysheet.getRange(i, 13).getValue();
    cost_of_month = cost_of_month + official; 
    mysheet.getRange(i, 3).setValue(cost_of_month);
    
    for(let j = 4; j < 9; j++){
      mysheet.getRange(i, j).setValue(cost_array[j - 4]);
    }
    mysheet.getRange(i, 9).setValue(cost_array[6]);
  
    console.log(`Summary作成`);
  
  }
  
  
  
  
  
  