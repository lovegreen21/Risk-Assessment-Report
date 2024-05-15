
 /*
************************************************
************** I. Chức năng chính **************
************************************************
*/
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var dDataValidSheet = spreadsheet.getSheetByName("Input");
var QLGE_SHEET_DAUTHANG = spreadsheet.getSheetByName("Quản lý gửi Email Đầu tháng");
var QLGE_SHEET = spreadsheet.getSheetByName("Quản lý gửi Email Giữa tháng");
var MSRR_SHEET = spreadsheet.getSheetByName("THMHSRR");
var QLPQF_SHEET = spreadsheet.getSheetByName("Quản lý phân quyền File");


/**
 * xóa dữ liệu sheet[THMHSRR], copy dữ liệu, format các file trong sheet[Input] các file 
 * @returns dán vào sheet[THMHSRR] và hiện thông báo
 */
function triggerImportMHSRR() {
    // 1.1 Đọc sheet[dDataValid1] 

dDataValidRange =  dDataValidSheet.getRange(4, 1, dDataValidSheet.getLastRow() - 3, 11 );
dDataValidValues = dDataValidRange.getValues();

// Xóa dữ liệu, format hiện tại 
destination = spreadsheet.getSheetByName("THMHSRR");
destination.getRange(8, 2, destination.getLastRow() , destination.getLastColumn() ).clear({contentsOnly: true, formatOnly: true})
var countCopy = 1;

for ( i = 0; i < dDataValidValues.length; i++) {
// lấy tên MHS của phòng ban

nameMHS = dDataValidValues[i][4];
// Lấy URL các File
  if (dDataValidValues[i][10] == "") { continue }
  else if( dDataValidValues[i][10] !== "") {
   
    fileHSRR = SpreadsheetApp.openByUrl(dDataValidValues[i][10]);
    sheets = fileHSRR.getSheets();

    // Loop sheetName 
    for (j = 0; j < sheets.length; j ++) { 
      var nameMHSArray = []
      // Kiểm tra sheet[F.02-QT.02.KSNB]
      if (sheets[j].getName().includes("F.02-QT.02.KSNB")) {
          var sheet =   sheets[j];
          var sheetLastRow = sheet.getLastRow();

          var numberRow = sheetLastRow - 18 - 12;
          var sheetLastColumn = sheet.getLastColumn();

          sourceRange = sheet.getRange(13, 1, numberRow, sheetLastColumn );
          for ( g = 0; g < numberRow; g++  ) {
            nameMHSArray.push([nameMHS]);
          }       
          
          var sourceValues = sourceRange.getValues();
          var sourceFormats = sourceRange.getNumberFormats();
          var sourceBackgrounds = sourceRange.getBackgrounds();
          var sourceFontColors = sourceRange.getFontColors();
          var sourceWrap = sourceRange.getWrap();

          if ( countCopy == 1) {

              destinationActual = destination.getRange( 8, 3, numberRow , sheetLastColumn );
              destinationActual.setValues(sourceValues);
              destinationActual.setNumberFormats(sourceFormats);
              destinationActual.setBackgrounds(sourceBackgrounds);
              destinationActual.setFontColors(sourceFontColors);
              destinationActual.setWrap(sourceWrap);
              destination.getRange( 8, 2, numberRow , 1 ).setValues(nameMHSArray);
              countCopy = countCopy + 1
              break
          }

          else if (countCopy > 1){
            destinationLastRow =  destination.getLastRow(); 
            destinationRangeValues = destination.getRange(8, 2, destinationLastRow - 7 , 2).getValues();

            var lastRowDestinationRange = 0;

             for (var k = destinationRangeValues.length - 1; k >= 0; k--) {
              var rowValues = destinationRangeValues[k];

              if(rowValues && rowValues.length > 0) {
                var isEmpty = rowValues.every(function (cellValue) {
                    return cellValue === '' ;
                });
              }
              if (!isEmpty) {
                  lastRowDestinationRange = k + 1; // Adjust to 1-based index
                  break;
              }
            }
              destinationActual = destination.getRange( lastRowDestinationRange + 8, 3, numberRow , sheetLastColumn );
              destinationActual.setValues(sourceValues);
              destinationActual.setNumberFormats(sourceFormats);
              destinationActual.setBackgrounds(sourceBackgrounds);
              destinationActual.setFontColors(sourceFontColors);
              destinationActual.setWrap(sourceWrap);
              destination.getRange( lastRowDestinationRange + 8, 2, numberRow , 1 ).setValues(nameMHSArray);
              break;  
          }
                
      }
    }
  }
}

}

/**
 * tìm hàng/cột cuối cùng các file
 * @returns điền công thức query vào sheet[Quản lý gửi Email Đầu tháng] và thông báo
 */

function triggerImportMSRRChuaDienDauThang() {
    c5Cell = QLGE_SHEET_DAUTHANG.getRange("C5");
  // Lấy cột cuối cùng
    var lastColumn = MSRR_SHEET.getLastColumn();
  // Lấy cột gần xếp hàng rủi ro
    var columnXephangHSRR =  lastColumn - 2 ; 
  // Chuyển thành tên
    var lastColumnLetterXephangHSRR = getColumnLetter(columnXephangHSRR);

    // Công thức formular
    var formulaC5 = `= ARRAYFORMULA ( 
                          IF (
                            B5:B <> "", 
                            COUNTIFS (THMHSRR!$B$8:$B,B5:B, THMHSRR!$${lastColumnLetterXephangHSRR}$8:$${lastColumnLetterXephangHSRR}, "=#N/A"),
                            ""
                          ) 
                      )`;   

    c5Cell.setFormula (formulaC5)

}

/**
 * tìm hàng/cột cuối cùng các file
 * @returns điền công thức query vào sheet[Quản lý gửi Email Giữa tháng] và thông báo
 */
function triggerImportMSRRChuaDienGiuaThang() {
    b4Cell = QLGE_SHEET.getRange("B4");
  
    var lastColumn = MSRR_SHEET.getLastColumn();
    var lastColumnLetter = getColumnLetter(lastColumn);
    // Tính từ cột B
    var columnXephangHSRR =  "Col" + (lastColumn - 2); 
    var columnMHS = "Col1" 
    var columnNSPT = "Col3"
  
    // Công thức query
    var formulaB4 = `=QUERY(THMHSRR!$B$8:${lastColumnLetter}, \"SELECT ${columnMHS}, COUNT(${columnMHS}) WHERE ${columnNSPT} IS NOT NULL AND (${columnXephangHSRR} = '#N/A' or ${columnXephangHSRR} is null) GROUP BY Col1 LABEL ${columnMHS} 'Mã Hồ sơ', COUNT(${columnMHS}) 'Số Mã số rủi ro không điền'\")`;   
    b4Cell.setFormula (formulaB4)
}

/**
 * Trigger chạy mỗi ngày/lần vào lúc 8h sáng , ktra ngày hnay
 * @returns kích hoạt gửi Email btnGuiDauThang() btnGuiGiuaThang() btnGuiCuoiThang theo điều kiện
 */
// Trigger chạy mỗi ngày/lần vào lúc 8h sáng 
function triggerGuiEmail() {

    var currentDate = new Date();
    var date = currentDate.getDate();
    var month = currentDate.getMonth();
    var year = currentDate.getFullYear();
  
    firstDateofCurrentMonth = new Date(year, month, 1);
    firstDateofCurrentMonthInt = firstDateofCurrentMonth.getDate();
    
    midDateofCurrentMonth = new Date(year, month, 15);
    midDateofCurrentMonthInt = midDateofCurrentMonth.getDate();
  
    // 3daybeforeLastdate
    lastDateofCurrentMonth = new Date(year, month + 1, -3);
    lastDateofCurrentMonthInt = lastDateofCurrentMonth.getDate();
  
    date24CurrentMonth = new Date (year, month, 24)
    date24CurrentMonthInt = date24CurrentMonth.getDate();
  
    if (date == firstDateofCurrentMonthInt) {
        return btnGuiDauThang()
    }
    else if (date == midDateofCurrentMonthInt) {
        return btnGuiGiuaThang()
    }
  
    else if (date == lastDateofCurrentMonth) {
        return btnGuiCuoiThang()
    }
  }
  /**
 * Phân quyền cho user vào các file input
 * @returns cấp quyền cho user , trả thông báo
 */
function btnCapNhatQuyen() {
    // Lấy dòng cuối cùng sheet[DANHMUC]
    lastRowDanhMuc = QLPQF_SHEET.getLastRow() ;
    // Lấy email và Permission tương ứng
    emailAndViewEditPermission = QLPQF_SHEET.getRange(6, 7, lastRowDanhMuc-5, 5).getValues() ;
    // Lấy email
    linkFile = QLPQF_SHEET.getRange(6,5,lastRowDanhMuc-5,1).getValues()
    
    // Số Row được thỏa mãn điều kiện
    var countEmailError = [];
    var countEmailSuccesful = []
    for( i = 0; i < emailAndViewEditPermission.length; i++ ) {
      if ( emailAndViewEditPermission[i][4] === true ) {
        for( j= 0; j < 4; j++){
          // Check email lỗi thì thông báo, vẫn tiếp tục cập nhật các email còn lại, sau đó phải set Manual
          try {
            if ( emailAndViewEditPermission[i][j] !== "" ) {
            // Logger.log(emailAndViewEditPermission[i][j])
            DriveApp.getFileById(SpreadsheetApp.openByUrl(linkFile[i][0]).getId()).addEditor(emailAndViewEditPermission[i][j])
            countEmailSuccesful.push(emailAndViewEditPermission[i][j])
            }
          }
          catch {
            countEmailError.push(emailAndViewEditPermission[i][j])
          }
        }
      }
    }

    if(countEmailSuccesful.length > 0 && countEmailError.length > 0) {
      SpreadsheetApp.getUi().alert("Đã cập nhật quyền chỉnh sửa thành công " + countEmailSuccesful.length + " Email. Vui lòng kiểm tra [Link], [Email]  [" + countEmailError + "]. Vui lòng cấp quyền lại thủ công hoặc chỉ tick Row đó. Rồi nhấn nút")
    }

    else if ( countEmailSuccesful.length > 0) {
      SpreadsheetApp.getUi().alert("Đã cập nhật quyền chỉnh sửa thành công " + countEmailSuccesful.length + " Email.")
    }

    else if ( countEmailError.length > 0) {
      SpreadsheetApp.getUi().alert("Không có email nào đúng, Vui lòng kiểm tra lại Email và Link File" )
    }
}



/*
************************************************
************** II. Helper functions ************
************************************************
*/
function getColumnLetter(columnNumber) {
    var columnLetter = "";
    while (columnNumber > 0) {
      var remainder = (columnNumber - 1) % 26;
      columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
      columnNumber = Math.floor((columnNumber - 1) / 26);
    }
    return columnLetter;
  }
  
  /**
   * Gửi Email
   */
  function mySendMail(mailTo, subject, body, cc) {
    MailApp.sendEmail ( 
      {
        to: mailTo,
        cc: cc,
        subject: subject,
        htmlBody: body,   
      }
    )
  }

/**
 * ktra bộ check MHSRR sheet[Quản lý gửi Email] cột F
 * @returns gửi email theo template MailDauThang.html
 */
function btnGuiDauThang(){

    var lastRow = QLGE_SHEET_DAUTHANG.getLastRow();
    var numberRow = lastRow - 4;
    var sourceValues = QLGE_SHEET_DAUTHANG.getRange(5,2,numberRow, 2).getValues();
    
    var timeZone = 'Asia/Ho_Chi_Minh';
    var currentDate = new Date 
    var currentMonth = currentDate.getMonth();
    var currentYear = currentDate.getFullYear();
  
    var lastDayOfMonth = getLastDayOfMonth(currentMonth, currentYear);
    var lastMonth =  Utilities.formatDate(new Date(currentYear, currentMonth , 0), timeZone, 'MM/yyyy' );
    var currentMonthmmyyyy =  Utilities.formatDate(new Date(currentYear, currentMonth ), timeZone, 'MM/yyyy' );
    
    var lastRowSourceRange = 0;
  
    for (var i = numberRow - 1; i >= 0; i--) {
        var rowValues = sourceValues[i];
  
        if(rowValues && rowValues.length > 0) {
          var isEmpty = rowValues.every(function (cellValue) {
              return cellValue === '' ;
          });
        }
        if (!isEmpty) {
            lastRowSourceRange = i + 1; // Adjust to 1-based index
            break;
        }
    }
  
    var dataMHSRR = []
    sourceRange = QLGE_SHEET_DAUTHANG.getRange(5,2, lastRowSourceRange, 6).getValues(); 
    
    // Kiểm tra MHS có tick thì run
    for ( i = 0; i < sourceRange.length; i++ ) {
      if ( sourceRange[i][3] === true) {
        dataMHSRR.push([sourceRange[i][0], sourceRange[i][1], sourceRange[i][2]])
      }
    }
  
    index = [];
    dataCheck = QLPQF_SHEET.getRange(6, 3, QLPQF_SHEET.getLastRow() - 5, 8).getValues();
    for ( i = 0; i < dataMHSRR.length; i++ ){
      var target = dataMHSRR[i][0];
      for (var j = 0; j < dataCheck.length; j++) {
        if (dataCheck[j][0] == target) {
          // sourceRange[[i][2]
          index.push( [j, dataMHSRR[i][1], dataMHSRR[i][2]] );
        }
      }
    } 
    dataSendEmail = [];
    for ( i = 0; i < index.length; i ++){
      dataCheck[index[i][0]].push( index[i][1], index[i][2] )
      dataSendEmail.push( dataCheck[index[i][0]] )
    }
  
    for ( i = 0; i < dataSendEmail.length; i ++){
      data = dataSendEmail[i]
      // Thêm ngày cuối cùng của tháng, tháng trước, tháng hiện tại 
      // 10,11,12
      data.push(lastDayOfMonth, lastMonth, currentMonthmmyyyy )
      let recipients =  [ data[5], data[6] ,data[7] ].filter(email => email);
      let mailTo = recipients.join(",");
      let ccEr = data[4];
      var nameDep = data[1];
      let subject = "NHẮC NHỞ THỰC HIỆN ĐÁNH GIÁ KẾ HOẠCH HÀNH ĐỘNG - HỒ SƠ RỦI RO - THÁNG " + (currentMonth + 1) + " - [" + nameDep + "]";
     
      const temp = HtmlService.createTemplateFromFile('MailDauThang');
      temp.data = data;
      message= temp.evaluate().getContent();
      let body = message;
  
    }
  }
  
  /**
   * ktra bộ check MHSRR sheet[Quản lý gửi Email] cột F
   * @returns gửi email theo template MailGiua/Cuoithang.html
   */
  function btnGuiGiuaThang(){
  
    var lastRow = QLGE_SHEET.getLastRow();
    var numberRow = lastRow - 4;
    var sourceValues = QLGE_SHEET.getRange(5,2,numberRow, 2).getValues();
    
    var timeZone = 'Asia/Ho_Chi_Minh';
    var currentDate = new Date 
    var currentMonth = currentDate.getMonth();
    var currentYear = currentDate.getFullYear();
  
    var lastDayOfMonth = getLastDayOfMonth(currentMonth, currentYear);
    var lastMonth =  Utilities.formatDate(new Date(currentYear, currentMonth , 0), timeZone, 'MM/yyyy' );
    var currentMonthmmyyyy =  Utilities.formatDate(new Date(currentYear, currentMonth ), timeZone, 'MM/yyyy' );
    
    var lastRowSourceRange = 0;
  
    for (var i = numberRow - 1; i >= 0; i--) {
        var rowValues = sourceValues[i];
  
        if(rowValues && rowValues.length > 0) {
          var isEmpty = rowValues.every(function (cellValue) {
              return cellValue === '' ;
          });
        }
  
        if (!isEmpty) {
            lastRowSourceRange = i + 1; // Adjust to 1-based index
            break;
        }
    }
    var dataMHSRR = []
    sourceRange = QLGE_SHEET.getRange(5,2, lastRowSourceRange, 6).getValues(); 
    
    // Kiểm tra MHS có tick thì run
    for ( i = 0; i < sourceRange.length; i++ ) {
      if ( sourceRange[i][3] === true) {
        dataMHSRR.push([sourceRange[i][0], sourceRange[i][1], sourceRange[i][2]])
      }
    }
  
    index = [];
  
    dataCheck = QLPQF_SHEET.getRange(6, 3, QLPQF_SHEET.getLastRow() - 5, 8).getValues();
  
    for ( i = 0; i < dataMHSRR.length; i++ ){
      var target = dataMHSRR[i][0];
      for (var j = 0; j < dataCheck.length; j++) {
        if (dataCheck[j][0] == target) {
          // sourceRange[[i][2]
          index.push( [j, dataMHSRR[i][1], dataMHSRR[i][2]] );
        }
      }
    }
  
    dataSendEmail = [];
    for ( i = 0; i < index.length; i ++){
      dataCheck[index[i][0]].push( index[i][1], index[i][2] )
      dataSendEmail.push( dataCheck[index[i][0]] )
    }
    //  Logger.log(dataSendEmail[0][8])
    for ( i = 0; i < dataSendEmail.length; i ++){
      data = dataSendEmail[i]
      // Thêm ngày cuối cùng của tháng, tháng trước, tháng hiện tại 
      data.push(lastDayOfMonth, lastMonth, currentMonthmmyyyy )
  
      let recipients =  [ data[5], data[6] ,data[7] ].filter(email => email);
      let mailTo = recipients.join(",");
      let ccEr = data[4];
      var nameDep = data[1];
      let subject = "NHẮC NHỞ THỰC HIỆN ĐÁNH GIÁ KẾ HOẠCH HÀNH ĐỘNG - HỒ SƠ RỦI RO - THÁNG " + (currentMonth + 1) + " - [" + nameDep + "]";
      
      const temp = HtmlService.createTemplateFromFile('MailGiua/Cuoithang');
      temp.data = data;
      message= temp.evaluate().getContent();
      let body = message;
      // Logger.log(mailTo)
      // Logger.log(ccEr)
  
      mySendMail(mailTo , subject, body, ccEr)
  
    }
  }