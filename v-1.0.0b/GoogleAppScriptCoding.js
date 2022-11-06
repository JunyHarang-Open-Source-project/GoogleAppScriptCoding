function Initalize() {                                          // 설문지 작성 시 Trigger 작동
    let triggers = ScriptApp.getProjectTriggers();

    for (let triggers in triggers) {
        ScriptApp.deleteTrigger(trigger[i]);
    }

    ScriptApp.newTrigger("mainFunction")
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onFormSubmit()
        .create();
} // Initalize() 끝

function mainFunction() {
    const activeSheet = SpreadsheetApp.getActiveSheet();
    const rangeData = activeSheet.getDataRange();
    const lastRow = rangeData.getLastRow();                                                                                                                                   // 마지막 행
    const lastCol = rangeData.getLastColumn();                                                                                                                                // 마지막 열
    const firstCol = lastCol - (lastCol - 1);                                                                                                                                 // 첫번째 열
    const detailViewURL = '{}';                                     // 구글 스프레드 시트 URL
    const slackURL = '{}';                                                                     // Slack Web Hook URL
    const alarmBotVersion = '1.0.0';                                                                                                                                          // App Script Version
    let responseValue = [];

    let slackMessage = '📢 꼼꼼한 친구들 Perfact Care 상담 / 견적 요청서 알림이 도착했어요.'
        + '\n ⓒ 2022. 주니하랑 All Rights Reserved.'
        + '\n 상담 / 견적 요청 알림 Bot Version : ' + alarmBotVersion
        + '\n 요청서 상세 보기 URL : ' + detailViewURL + '\n\n\n\n';

    let errorSlackMessage = '';

    responseValue = spreadSheetsLoop(activeSheet, firstCol, lastRow, slackURL, slackMessage, detailViewURL, alarmBotVersion);

    errorCheck(responseValue, detailViewURL, alarmBotVersion);
} // mainFunction() 끝

function spreadSheetsLoop(activeSheet, firstCol, lastRow,slackURL, slackMessage, detailViewURL, alarmBotVersion) {

    for (row = 2; row < lastRow + 1; row++) {                                                                                                                                 // 스프레드 시트 2번째 행 부터 한 줄씩 반복

        let timeStamp = activeSheet.getRange(row, firstCol).getValue();                                                                                                           // 2행 1열(A2) 부터 값을 가져와서 요청서 작성 일자를 변수에 저장

        if(timeStamp === undefined || timeStamp === '') {                                                                                                                         // timestamp 값이 비어 있는가?
            break;                                                                                                                                                                  // 반복문 종료
        }

        let date = Utilities.formatDate(timeStamp, "GMT+9", "yyyy년MM월dd일 HH시MM분");                                                                                              // 2022년 11월 04일 09시00분과 같은 형식으로 일시 처리를 위한 변수
        let privacyInfoAgree = activeSheet.getRange(row, firstCol + 1).getValue();                                                                                                // 개인정보 수집 동의 여부
        let email = activeSheet.getRange(row, firstCol + 17).getValue();                                                                                                          // 고객 Email 주소
        let name = activeSheet.getRange(row, firstCol + 2).getValue();                                                                                                            // 고객 이름
        let phoneNumber = activeSheet.getRange(row, firstCol + 3).getValue();                                                                                                     // 고객 연락처
        let address = activeSheet.getRange(row, firstCol + 4).getValue();                                                                                                         // 작업 요청 현장 주소
        let jobType = activeSheet.getRange(row, firstCol + 5).getValue();                                                                                                         // 작업 종류
        let cleanType = activeSheet.getRange(row, firstCol + 6).getValue();                                                                                                       // 청소 종류
        let jobDateTimeStamp = activeSheet.getRange(row, firstCol + 7).getValue();                                                                                                // 작업 희망일 TimeStamp Value
        let jobTimeTimeStamp = activeSheet.getRange(row, firstCol + 8).getValue();                                                                                                // 작업 희망 시간 TimeStamp Value
        let jobDate = Utilities.formatDate(jobDateTimeStamp, "GMT+9", "yyyy년MM월dd일");                                                                                            // 작업 희망일 TimeStamp 값 형식 변경
        let jobTime = Utilities.formatDate(jobTimeTimeStamp, "GMT+9", "HH시MM분");                                                                                                 // 작업 희망 시간 TimeStamp 값 형식 변경
        let spotType = activeSheet.getRange(row, firstCol + 9).getValue();                                                                                                        // 현장 종류
        let spotFloor = activeSheet.getRange(row, firstCol + 10).getValue();                                                                                                      // 현장 층 수
        let isLift = activeSheet.getRange(row, firstCol + 11).getValue();                                                                                                         // 승강기 존재 유무
        let callDateTimeStamp = activeSheet.getRange(row, firstCol + 12).getValue();                                                                                              // 상담 가능 일자 TimeStamp Value
        let callTimeTimeStamp = activeSheet.getRange(row, firstCol + 13).getValue();                                                                                              // 상담 가능 시간 TimeStamp Value
        let callDate = Utilities.formatDate(callDateTimeStamp, "GMT+9", "yyyy년MM월dd일");                                                                                          // 상담 가능 일자 TimeStamp 값 형식 변경
        let callTime = Utilities.formatDate(callTimeTimeStamp, "GMT+9", "HH시MM분");                                                                                               // 상담 가능 시간 TimeStamp 값 형식 변경
        let spotDetail = activeSheet.getRange(row, firstCol + 14).getValue();                                                                                                     // 현장 상세
        let etc = activeSheet.getRange(row, firstCol + 15).getValue();                                                                                                            // 기타
        let sendEmailStatus = activeSheet.getRange(row, firstCol + 16).getValue();                                                                                                // 고객 접수 완료 Email 발송 여부

        const emailTest = 0;

        slackMessage += row - 1 + '번째 상담 / 견적 요청 고객 정보 입니다.'
            + '\n -------------------'
            + '\n 상담 / 견적 요청 일시 : ' + date
            + '\n 개인 정보 수집 동의 여부 : ' + privacyInfoAgree
            + '\n 고객 Email 주소 : ' + email
            + '\n 고객 이름 : ' + name
            + '\n 고객 연락처 : ' + phoneNumber
            + '\n 현장 주소 : ' + address
            + '\n 요청 작업 종류 : ' + jobType
            + '\n 요청 청소 종류 : ' + cleanType
            + '\n 요청 작업 일시 : ' + jobDate + ' ' + jobTime
            + '\n 현장 종류 : ' + spotType
            + '\n 현장 층 수 : ' + spotFloor
            + '\n 승강기 존재 여부 : ' + isLift
            + '\n 상담 요청 일시 : ' + callDate + ' ' + callTime
            + '\n 현장 상세 : ' + spotDetail
            + '\n 기타 사항 : ' + etc
            + '\n -------------------'
            + '\n\n';

        let slackSendReturnValue = sendSlack(slackURL, slackMessage, detailViewURL, alarmBotVersion);
        let emailSendReturnValue = '';

        if (sendEmailStatus === undefined || sendEmailStatus === '' || sendEmailStatus === '미발송') {
            emailSendReturnValue = sendMail(activeSheet,row, firstCol, email, emailTest, name, date);

            Logger.log('Email 전송 Error 내용 : ' + emailSendReturnValue);
        }

        let returnValue = new Array();

        returnValue.push(slackSendReturnValue);
        returnValue.push(emailSendReturnValue);

        return returnValue;
    } // for문 끝
} // spreadsheetsLoop(activeSheet, firstCol, lastRow,slackURL, slackMessage, detailViewURL, alarmBotVersion) 끝

function sendSlack(slackURL, slackMessage, detailViewURL, alarmBotVersion) {
    slackMessage += '\n\n 상담 / 견적 요청 알림 Bot Version : ' + alarmBotVersion
        + '\n요청서 상세 보기 URL : ' + detailViewURL + '\n';
    + 'ⓒ 2022. 주니하랑 All Rights Reserved. \n\n\n\n';

    let slackErrorCheck = '';

    let payload = {                                                                                                                                                                 // Slack으로 보낼 Message 준비
        'text' : slackMessage
    };

    let option = {                                                                                                                                                                  // Slack Web Hook 정보
        'method' : 'post',
        'contentType' : 'application/json',
        'payload' : JSON.stringify(payload)
    };

    try {
        UrlFetchApp.fetch(slackURL, option);                                                                                                                                           // Slack에 보내기기
        Logger.log('꼼꼼한 친구들 Perfact Care 고객 상담 / 견적 요청 정보를 Slack에 성공적으로 보냈습니다.');

    } catch(err) {
        slackErrorCheck = '꼼꼼한 친구들 Perfact Care 고객 상담 / 견적 요청 정보를 Slack 전송 실패 하였습니다.\n'
            + '문제 정보 : ' + err
            + '\n\n\n\n';

        Logger.log(slackErrorCheck);

    } finally {
        Logger.log('꼼꼼한 친구들 Perfact Care 고객 상담 / 견적 요청 정보 Slack 작업 처리 완료 되었습니다.');

        return slackErrorCheck;
    }
}  // sendSlack(slackURL, slackMessage, detailViewURL, alarmBotVersion) 끝

function sendMail(activeSheet,row, firstCol, email, emailTest, name, date) {
    const emailSubject = '[꼼꼼한 친구들 - Perfact Care] 상담 / 견적 요청서 정상 접수';
    const emailBody = '안녕하십니까? ' + name + ' 고객님'
        +'\n 꼼꼼한 친구들 Perfect Care 입니다.'
        +'\n' + date + '에 요청 주신 상담 / 견적 요청서 정상 접수 되었습니다.'
        +'\n 담당 영업 대표가 해당 내용을 확인하고,'
        +'\n 작성해 주신 상담 일시에 맞추어 연락드리겠습니다.'
        +'\n 추가 문의 사항이 있으시다면 꼼꼼한 친구들 카카오 채팅 고객 센터 혹은 010-4828-2711로 연락 주시면 친절하고, 빠르게 안내 도와드리겠습니다.'
        +'\n 언제나 좋은 일만 가득하시기 바라겠고, 저희 꼼꼼한 친구들을 찾아주셔서 대단히 감사드립니다.'
        +'\n 그럼 빠른 시일 내에 인사 드리도록 하겠습니다.'
        +'\n 감사합니다. :)'
        +'\n\n\n\n'
        +'본 Mail은 자동화 Bot에 의해 발송 되었습니다.';

    let clientEmailSendErrorCheck = '';

    if(emailTest === 1) {                                                                                                                                                        // Mail Test 여부 확인(0 = False, 1 = True)
        email = 'junyharang8592@gmail.com';                                                                                                                                        // Mail Test 시 고객에게 Mail 전송을 막기 위해 내부 이용자 Mail 주소로 치환
    }

    try {
        GmailApp.sendEmail(email, emailSubject, emailBody);
        activeSheet.getRange(row, firstCol + 16).setValue('발송');                                                                                                                    // Mail을 보내게 되면 16번째 스프레드 시트 열에 '발송'이라고 기재
        Logger.log('꼼꼼한 친구들 상담 / 견적 요청 고객에게 Email이 발송 되었어요.');

    } catch(err) {
        clientEmailSendErrorCheck = '꼼꼼한 친구들 Perfact Care 고객 상담 / 견적 요청 완료 및 감사 인사 관련 고객 Email 전송 실패 하였습니다.\n'
            + '문제 정보 : ' + err
            + '\n\n\n\n';

        Logger.log(clientEmailSendErrorCheck);

    } finally {
        Logger.log('꼼꼼한 친구들 Perfact Care 고객 상담 / 견적 요청 완료 및 감사 인사 관련 Email 전송 작업 처리 되었습니다.');

        return clientEmailSendErrorCheck;
    }
} // sendMail(activeSheet,row, firstCol, email, emailTest, name, date) 끝

function errorCheck(responseValue, detailViewURL, alarmBotVersion) {

    errorSlackMessage = '🆘 자동화 처리 중 문제 발생하였습니다.'
        + '\n 상담 / 견적 요청 알림 Bot Version : ' + alarmBotVersion
        + '\n\n 요청서 상세 보기 URL : ' + detailViewURL + '\n';
    + 'ⓒ 2022. 주니하랑 All Rights Reserved. \n\n\n\n';

    for(idx = 0; idx < responseValue.length; idx++) {
        if (responseValue[idx] !== '') {
            errorSlackMessage += responseValue[idx];
            sendSlackError(errorSlackMessage);
        }
    }
} // errorCheck(responseValue, detailViewURL, alarmBotVersion) 끝

function sendSlackError(errorSlackMessage) {

    let sendSlackErrorURL = '{}';

    let payload = {                                                                                                                                                                 // Slack으로 보낼 Message 준비
        'text' : errorSlackMessage
    };

    let option = {                                                                                                                                                                  // Slack Web Hook 정보
        'method' : 'post',
        'contentType' : 'application/json',
        'payload' : JSON.stringify(payload)
    };

    try {
        UrlFetchApp.fetch(sendSlackErrorURL, option);                                                                                                                                           // Slack에 보내기기
        Logger.log('꼼꼼한 친구들 Perfact Care 고객 상담 / 견적 요청 정보를 Slack에 성공적으로 보냈습니다.');

    } catch(err) {
        Logger.log('꼼꼼한 친구들 Perfact Care 고객 상담 / 견적 요청 정보를 Slack 전송 실패 하였습니다.\n');
        Logger.log('문제 정보 : ' + err);
    } finally {
        Logger.log('꼼꼼한 친구들 Perfact Care 고객 상담 / 견적 요청 정보를 Slack 작업이 처리 되었습니다.');
    }
} // sendSlackError(errorSlackMessage) 끝