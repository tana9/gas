// LAVAの予約完了メールから予定を自動で登録する

var EVENT_TITLE = "ヨガ"
var MAIL_FILTER = "in:inbox from:reserve@yoga-lava.com"

// 「11月14日(土) 11:30～12:30」を開始日と終了日に分割する
function startAndEndDate_(date){
  var d = date.match(/\d./g); // 日付中の数値のみを抽出する
  var year = new Date().getFullYear();
  var date = year + "/" + d[0] + "/" + d[1]
  var startTime = d[2] + ":" + d[3]
  var endTime = d[4] + ":" + d[5]
  var startDate = new Date(date + " " + startTime);
  var endDate = new Date(date + " " + endTime);
  return [startDate, endDate];
}

// イベント追加
function addEvent_(title, startTime, endTime, options){
  // イベントが登録済みかを確認する
  var events = CalendarApp.getEvents(startTime, endTime);
  for(var i=0; i<events.length; i++){
    if(events[i].getTitle() === EVENT_TITLE){
      return false;
    }
  }
  CalendarApp.createEvent(title, startTime, endTime, options);
  return true;
}

// 登録したイベントをスプレッドシートに記録する
function addLog_(startTime, endTime, tenpo, cource, tanto){
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow([startTime, endTime, tenpo, cource, tanto])
}

function addLAVAEvent(){
  var threads = GmailApp.search(MAIL_FILTER)
  for(var i=0; i<threads.length; i++){
    var subject = threads[i].getFirstMessageSubject(); // メールのタイトル取得
    
    if(0 < subject.indexOf("予約完了")){
      var rows = threads[i].getMessages()[0].getPlainBody().split("\n");
      for(var j=0; j<rows.length; j++){
        var row = rows[j];
        if(0 < row.indexOf("レッスンの日時は以下となります。")){
          var tenpo = rows[j+2] // 店舗名
          var date = rows[j+3]  // 日付
          var dates = startAndEndDate_(rows[j+3]); // 開始日と終了日
          var course = rows[j+4] // ヨガのコース
          var tanto = rows[j+5]  // レッスン担当
          if(addEvent_(EVENT_TITLE, dates[0], dates[1], {location: tenpo})){
            addLog_(dates[0], dates[1], tenpo, course, tanto);
          }
          break;
        }
      }
    }
  }
}

function test_startAndEndDate_(){
  var date = startAndEndDate("11月07日(土) 16:00～17:00");
  Logger.log(date);
}




