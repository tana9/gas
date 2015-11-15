// LAVA�̗\�񊮗����[������\��������œo�^����

var EVENT_TITLE = "���K"
var MAIL_FILTER = "in:inbox from:reserve@yoga-lava.com"

// �u11��14��(�y) 11:30�`12:30�v���J�n���ƏI�����ɕ�������
function startAndEndDate_(date){
  var d = date.match(/\d./g); // ���t���̐��l�݂̂𒊏o����
  var year = new Date().getFullYear();
  var date = year + "/" + d[0] + "/" + d[1]
  var startTime = d[2] + ":" + d[3]
  var endTime = d[4] + ":" + d[5]
  var startDate = new Date(date + " " + startTime);
  var endDate = new Date(date + " " + endTime);
  return [startDate, endDate];
}

// �C�x���g�ǉ�
function addEvent_(title, startTime, endTime, options){
  // �C�x���g���o�^�ς݂����m�F����
  var events = CalendarApp.getEvents(startTime, endTime);
  for(var i=0; i<events.length; i++){
    if(events[i].getTitle() === EVENT_TITLE){
      return false;
    }
  }
  CalendarApp.createEvent(title, startTime, endTime, options);
  return true;
}

// �o�^�����C�x���g���X�v���b�h�V�[�g�ɋL�^����
function addLog_(startTime, endTime, tenpo, cource, tanto){
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow([startTime, endTime, tenpo, cource, tanto])
}

function addLAVAEvent(){
  var threads = GmailApp.search(MAIL_FILTER)
  for(var i=0; i<threads.length; i++){
    var subject = threads[i].getFirstMessageSubject(); // ���[���̃^�C�g���擾
    
    if(0 < subject.indexOf("�\�񊮗�")){
      var rows = threads[i].getMessages()[0].getPlainBody().split("\n");
      for(var j=0; j<rows.length; j++){
        var row = rows[j];
        if(0 < row.indexOf("���b�X���̓����͈ȉ��ƂȂ�܂��B")){
          var tenpo = rows[j+2] // �X�ܖ�
          var date = rows[j+3]  // ���t
          var dates = startAndEndDate_(rows[j+3]); // �J�n���ƏI����
          var course = rows[j+4] // ���K�̃R�[�X
          var tanto = rows[j+5]  // ���b�X���S��
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
  var date = startAndEndDate("11��07��(�y) 16:00�`17:00");
  Logger.log(date);
}




