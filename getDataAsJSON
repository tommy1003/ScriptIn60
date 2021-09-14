// Please follow this if you know JSON and want to convert 
// sheet data to JSON

// Nếu bạn biết dữ liệu JSON và muốn chuyển đổi dữ liệu sheet 
// thành JSON thì làm theo các bước sau nhé

// Follow ScriptIn60 để biết thêm nhiều code hay nhé

// Lập trình theo yêu cầu
// Firebase, VBA, Apps Script, Python, MongoDB, tự động hóa Google Sheets, tạo API và Zalo OA chatbot
// HOTLINE: 078 600 5534 (Zalo)

function getDataAsJSON(){
  var ss = SpreadsheetApp.getActive();
  var sht = ss.getSheetByName('Sheet1');
  var [header, ...data] = sht.getDataRange().getValues();
  var dataArr = [];
  for (let row of data) {
    let obj = {};
    for (let i=0; i<header.length; i++) {
      obj[header[i]] = row[i];
    }
    dataArr.push(obj);
  }
  console.log(dataArr);
}
