/** @OnlyCurrentDoc */

function addCalenderCol() {
  // 翌月1日を取得
  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth()+2;
  var date = 1;

  // 1日の曜日を取得
  var aryDay = new Array('日', '月', '火', '水', '木', '金', '土');
  var dayNum = new Date(year + "/" + month + "/" + date).getDay();
  var nowWeek = aryDay[dayNum];
  
  // シートコピー
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmp = ss.getSheetByName("y年m月");
  var nn = tmp.copyTo(ss);
  nn.setName(year + '年' + month + '月');
  ss.setActiveSheet(nn);
  ss.moveActiveSheet(1);
  
  // 日付連続データ
  ss.getRange('J4').activate();
  ss.getCurrentCell().setValue(month + '/' + date);
  ss.getActiveRange().autoFill(ss.getRange('J4:AN4'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // 曜日連続データ
  ss.getRange('J3').activate();
  ss.getCurrentCell().setValue(nowWeek);
  ss.getActiveRange().autoFill(ss.getRange('J3:AN3'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // 背景色
  var sh = ss.getActiveSheet();
  for(let i = 10; i <= 40; i++) {
    //sh.getRange({行番号},{列番号},{行数},{列数});
    if(sh.getRange(3, i).getValue() == '土'){
      sh.getRange(4, i, 31, 1).activate();
      ss.getActiveRangeList().setBackground('#87ceeb');
    }else if(sh.getRange(3, i).getValue() == '日'){
      sh.getRange(4, i, 31, 1).activate();
      ss.getActiveRangeList().setBackground('#ffb9cb');
    }
  }
  
  // 2月は28日、4,6,9,11月は30日
  if(month == 2){
    sh.deleteColumns(38, 3);
  }else if(month == 4 || month == 6 || month == 9 || month == 11){
    sh.deleteColumn(40);
  }

};

