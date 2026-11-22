/**
 * 업무자동화 — 폰 입력 웹앱 (Google Apps Script)
 *
 * 하는 일
 *  - 구글 드라이브 '내 드라이브/FormFiller_inbox/phone_form.json' (PC 프로그램이 써 둔 입력 양식)을 읽어 폰 화면을 만든다.
 *  - [올리기] 하면 inbox 안에 '날짜_시각_기기' 폴더를 만들고 사진 → input.json 순서로 저장한다.
 *  - 출근해서 켜 둔 PC 가 1분 안에 읽어서 문서를 만들고 result.json 을 남긴다. 폰 화면은 그 결과를 보여 준다.
 *
 * 보안
 *  - 배포할 때 '실행: 나', '액세스: 나만' 으로 한다. 내 구글 계정으로 로그인한 폰에서만 열린다.
 *  - 이 스크립트는 FormFiller_inbox 폴더만 읽고 쓴다.
 */

/** PC 설정(config.json 의 phoneInboxFolder)을 바꿨다면 여기도 같은 이름으로 */
var INBOX_NAME = 'FormFiller_inbox';
var FORM_FILE = 'phone_form.json';
var MAX_PHOTO_BYTES = 8 * 1024 * 1024;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('업무자동화 폰 입력')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function inbox_() {
  var it = DriveApp.getRootFolder().getFoldersByName(INBOX_NAME);
  if (!it.hasNext()) throw new Error("드라이브에 '" + INBOX_NAME + "' 폴더가 없습니다. PC 프로그램을 한 번 켜 주세요.");
  return it.next();
}

function readJson_(folder, name) {
  var files = folder.getFilesByName(name);
  if (!files.hasNext()) return null;
  return JSON.parse(files.next().getBlob().getDataAsString('UTF-8'));
}

/** 폰 화면 양식 + 최근 올린 것들의 처리 결과 */
function getForm() {
  var box = inbox_();
  var form = readJson_(box, FORM_FILE);
  if (!form) throw new Error('PC 가 아직 입력 양식을 만들지 않았습니다. PC 에서 업무자동화를 켜 주세요.');
  return JSON.stringify({ form: form, recent: recent_(box, 6) });
}

function recent_(box, count) {
  var list = [];
  var it = box.getFolders();
  while (it.hasNext()) {
    var f = it.next();
    if (/^\d{4}-\d{2}-\d{2}_\d{6}_/.test(f.getName())) list.push(f);
  }
  list.sort(function (a, b) { return a.getName() < b.getName() ? 1 : -1; });
  return list.slice(0, count).map(function (f) {
    return { name: f.getName(), result: readJson_(f, 'result.json') };
  });
}

/** 올릴 폴더를 만든다. 돌려준 이름으로 사진·입력을 올린다 */
function startSubmission(date, device) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('날짜 형식이 틀렸습니다');
  var dev = String(device || 'phone').replace(/[^A-Za-z0-9]/g, '').slice(0, 20) || 'phone';
  var time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HHmmss');
  var name = date + '_' + time + '_' + dev;
  inbox_().createFolder(name);
  return name;
}

function submissionFolder_(name) {
  if (!/^\d{4}-\d{2}-\d{2}_\d{6}_[A-Za-z0-9]+$/.test(name)) throw new Error('폴더 이름이 올바르지 않습니다');
  var it = inbox_().getFoldersByName(name);
  if (!it.hasNext()) throw new Error('올릴 폴더가 없습니다: ' + name);
  return it.next();
}

/** 사진 한 장 (base64 JPEG). 파일 이름은 p1.jpg 같은 단순한 이름만 */
function uploadPhoto(folderName, fileName, base64) {
  if (!/^p\d{1,3}\.jpg$/.test(fileName)) throw new Error('사진 이름이 올바르지 않습니다');
  var bytes = Utilities.base64Decode(base64);
  if (bytes.length > MAX_PHOTO_BYTES) throw new Error('사진이 너무 큽니다');
  submissionFolder_(folderName).createFile(Utilities.newBlob(bytes, 'image/jpeg', fileName));
  return true;
}

/** 마지막에 input.json 을 쓴다 (PC 는 이 파일이 있어야 처리한다) */
function finishSubmission(folderName, inputJson) {
  var input = JSON.parse(inputJson);   // 형식 확인 (잘못된 글자면 여기서 오류)
  if (input.schemaVersion !== 2) throw new Error('형식 버전이 다릅니다');
  submissionFolder_(folderName).createFile('input.json', JSON.stringify(input), 'application/json');
  return true;
}

/** 처리 결과 (PC 가 만들기 전이면 null) */
function getResult(folderName) {
  return JSON.stringify(readJson_(submissionFolder_(folderName), 'result.json'));
}
