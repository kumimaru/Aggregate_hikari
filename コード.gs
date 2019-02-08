/* 指定月の特定カレンダーからイベントすべてを取得してスプレッドシートに書き出す */
function getCalendar() {

  var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得
//  var RANGE = mySheet.getRange("A:A").getValues().filter(String).length;  // スプレッドシート：開始位置
//  mySheet.getRange((RANGE + 1), 1).activate();
  var RANGE = 1;
  var FORMAT_TIME = 'mm/dd hh:mm';  // スプレッドシート

  var myCal=CalendarApp.getCalendarById('bgq6sq6oh7l7ptkig2lmboulq8@group.calendar.google.com'); //特定IDのカレンダーを取得

  var date='2018/10/01 00:00:00'; //対象月を指定
  var startDate=new Date(date); //取得開始日
  var endDate=new Date(date);
  endDate.setMonth(endDate.getMonth()+5);　//取得終了日（自動計算） ※一ヶ月単位

  var schedules = myCal.getEvents(startDate,endDate);  //予定オブジェクトの生成

  // 予定を繰り返し出力する
  for(var index = 0; index < schedules.length; index++) {
    var range = RANGE + index;
    // カレンダー名を出力
    mySheet.getRange(range, 1).setValue(myCal.getName());
    // 予定名を出力
    mySheet.getRange(range, 2).setValue(schedules[index].getTitle());
    // 開始時間を出力
    mySheet.getRange(range, 3).setValue(schedules[index].getStartTime()).setNumberFormat(FORMAT_TIME);
    // 終了時間を出力
    mySheet.getRange(range, 4).setValue(schedules[index].getEndTime()).setNumberFormat(FORMAT_TIME);
    // イベント内容を出力
    mySheet.getRange(range, 6).setValue(schedules[index].getDescription()).setNumberFormat(FORMAT_TIME);

  }

} // end function
