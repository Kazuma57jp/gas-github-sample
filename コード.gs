function EmailExtractTransfer() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var thds = GmailApp.search('newer_than:3d l:^ss_sr');
  var row = 2;
  
  for(var n in thds){
      var thd = thds[n]; // インデックスn番のスレッドを変数thdに代入する
      var msgs = thd.getMessages(); // 変数thd のメッセージを変数 msgs に代入する  
    
    for(m in msgs){
      var msg = msgs[m]; // 取り出したメッセージのインデックス0番（最初のメッセージ）を変数msgに代入する
      var Date = msg.getDate(); // 変数msg に代入されている日付データを変数Dateに代入する
      var From = msg.getFrom(); // 変数msg に代入されている差出人データを変数Fromに代入する
      var subject= msg.getSubject();　//変数msgに代入されている件名データを編集subjectに代入する
      
      sheet.getRange(row,1).setValue(Date);
      sheet.getRange(row,2).setValue(From);
      sheet.getRange(row,3).setValue(subject);
      row++;
    }
      
  }
}
