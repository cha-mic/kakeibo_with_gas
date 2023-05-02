function SetResponses(e) {

    var responses = e.response.getItemResponses();//回答の収集
  
    const sheet_id = ''
    const sheet_name = '2023'
  
    const sheet = SpreadsheetApp.openById(sheet_id).getSheetByName(sheet_name);//スプレッドシートを開く
  
    const today_or_yesterday =  responses[0].getResponse() ; //昨日or今日
    const price              =  responses[1].getResponse(); //金額
    const type               =  responses[2].getResponse() ; //種類
    
    //日付と曜日を計算する（0:sunday 6:saturday）
    let day_of_week;
    let month,date;
    let today     = new Date();
    let yesterday = new Date(today); 　
    yesterday.setDate(today.getDate()-1);
   
    if( today_or_yesterday == '今日' ){
      day_of_week = today.getDay();
      month       = today.getMonth()+1;
      date        = today.getDate();
    }
    else if( today_or_yesterday == '昨日' ){
      day_of_week = yesterday.getDay();
      month       = yesterday.getMonth()+1;
      date        = yesterday.getDate();
    }
  
    // シートに記入
    const addarray = [month*100+date,day_of_week,price,type];
    sheet.appendRow(addarray);
  
  
  }
  