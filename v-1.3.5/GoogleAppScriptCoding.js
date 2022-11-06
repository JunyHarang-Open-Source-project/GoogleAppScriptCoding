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
    Logger.log('mainFunction 작동 되었습니다.');

    const activeSheet = SpreadsheetApp.getActiveSheet();
    const rangeData = activeSheet.getDataRange();
    const lastRow = rangeData.getLastRow();                                                                                                                                   // 마지막 행
    const lastCol = rangeData.getLastColumn();                                                                                                                                // 마지막 열
    const firstCol = lastCol - (lastCol - 1);                                                                                                                                  // 첫번째 열
    const alarmBotVersion = '1.1.0';                                                                                                                                          // App Script Version
    let responseValue = [];
    const detailViewURL = '';                                                                                                                                                 // 구글 스프레드 시트 URL

    responseValue = spreadSheetsLoop(activeSheet, firstCol, lastRow, detailViewURL, alarmBotVersion);

    if (responseValue !== undefined || responseValue !== '' || responseValue !== ',,') {
        errorCheck(responseValue, detailViewURL, alarmBotVersion);
    }
} // mainFunction() 끝

function spreadSheetsLoop(activeSheet, firstCol, lastRow, detailViewURL, alarmBotVersion) {

    let date = '';
    let privacyInfoAgree = '';
    let email = '';
    let name = '';
    let phoneNumber = '';
    let address = '';
    let jobType = '';
    let cleanType = '';
    let jobDateTimeStamp = '';
    let jobTimeTimeStamp = '';
    let jobDate = '';
    let jobTime = '';
    let spotType = '';
    let spotFloor = '';
    let isLift = '';
    let callDateTimeStamp = '';
    let callTimeTimeStamp = '';
    let callDate = '';
    let callTime = '';
    let spotDetail = '';
    let etc = '';

    const emailTest = 0;
    let loopCount = 0;
    let spreadSheetsLoopErrorMessage = '';

    for (row = 2; row < lastRow + 1; row++) {                                                                                                                                     // 스프레드 시트 2번째 행 부터 한 줄씩 반복
        loopCount += 1;

        let timeStamp = activeSheet.getRange(row, firstCol).getValue();                                                                                                           // 2행 1열(A2) 부터 값을 가져와서 요청서 작성 일자를 변수에 저장

        if(timeStamp === undefined || timeStamp === '') {                                                                                                                         // timestamp 값이 비어 있는가?
            break;                                                                                                                                                                // 반복문 종료
        }

        date = Utilities.formatDate(timeStamp, "GMT+9", "yyyy년MM월dd일 HH시MM분");                                                                                              // 2022년 11월 04일 09시00분과 같은 형식으로 일시 처리를 위한 변수
        privacyInfoAgree = activeSheet.getRange(row, firstCol + 1).getValue();                                                                                                // 개인정보 수집 동의 여부
        email = activeSheet.getRange(row, firstCol + 17).getValue();                                                                                                          // 고객 Email 주소
        name = activeSheet.getRange(row, firstCol + 2).getValue();                                                                                                            // 고객 이름
        phoneNumber = activeSheet.getRange(row, firstCol + 3).getValue();                                                                                                     // 고객 연락처
        address = activeSheet.getRange(row, firstCol + 4).getValue();                                                                                                         // 작업 요청 현장 주소
        jobType = activeSheet.getRange(row, firstCol + 5).getValue();                                                                                                         // 작업 종류
        cleanType = activeSheet.getRange(row, firstCol + 6).getValue();                                                                                                       // 청소 종류
        jobDateTimeStamp = activeSheet.getRange(row, firstCol + 7).getValue();                                                                                                // 작업 희망일 TimeStamp Value
        jobTimeTimeStamp = activeSheet.getRange(row, firstCol + 8).getValue();                                                                                                // 작업 희망 시간 TimeStamp Value
        jobDate = Utilities.formatDate(jobDateTimeStamp, "GMT+9", "yyyy년MM월dd일");                                                                                            // 작업 희망일 TimeStamp 값 형식 변경
        jobTime = Utilities.formatDate(jobTimeTimeStamp, "GMT+9", "HH시MM분");                                                                                                 // 작업 희망 시간 TimeStamp 값 형식 변경
        spotType = activeSheet.getRange(row, firstCol + 9).getValue();                                                                                                        // 현장 종류
        spotFloor = activeSheet.getRange(row, firstCol + 10).getValue();                                                                                                      // 현장 층 수
        isLift = activeSheet.getRange(row, firstCol + 11).getValue();                                                                                                         // 승강기 존재 유무
        callDateTimeStamp = activeSheet.getRange(row, firstCol + 12).getValue();                                                                                              // 상담 가능 일자 TimeStamp Value
        callTimeTimeStamp = activeSheet.getRange(row, firstCol + 13).getValue();                                                                                              // 상담 가능 시간 TimeStamp Value
        callDate = Utilities.formatDate(callDateTimeStamp, "GMT+9", "yyyy년MM월dd일");                                                                                          // 상담 가능 일자 TimeStamp 값 형식 변경
        callTime = Utilities.formatDate(callTimeTimeStamp, "GMT+9", "HH시MM분");                                                                                               // 상담 가능 시간 TimeStamp 값 형식 변경
        spotDetail = activeSheet.getRange(row, firstCol + 14).getValue();                                                                                                     // 현장 상세
        etc = activeSheet.getRange(row, firstCol + 15).getValue();                                                                                                            // 기타
        sendEmailStatus = activeSheet.getRange(row, firstCol + 16).getValue();                                                                                                // 고객 접수 완료 Email 발송 여부

        let bodyMessage = row - 1 + '번째 상담 / 견적 요청 고객 정보 입니다.'
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

        spreadSheetsLoopErrorMessage = sendBridge(
            bodyMessage,
            detailViewURL,
            alarmBotVersion,
            activeSheet,
            row,
            lastRow,
            firstCol,
            email,
            emailTest,
            name,
            date,
            loopCount
        );
    } // for문 끝

    return spreadSheetsLoopErrorMessage

} // spreadSheetsLoop(activeSheet, firstCol, lastRow, detailViewURL, alarmBotVersion) 끝

function sendBridge(bodyMessage, detailViewURL, alarmBotVersion, activeSheet, row, lastRow, firstCol, email, emailTest, name, date, loopCount) {

    let headerMessage = '📢' + date
        + '\n 꼼꼼한 친구들 Perfect Care 상담 / 견적 요청서 알림이 도착했어요. 📢 \n';

    let commonMessage = '\n ⓒ 2022. 주니하랑(junyharang8592@gmail.com) All Rights Reserved. Blog : https://junyharang.tistory.com/'
        + '\n 상담 / 견적 요청 알림 Bot Version : ' + alarmBotVersion
        + '\n 요청서 상세 보기 URL : ' + detailViewURL + '\n\n\n\n';

    const slackURL = '';                                                                     // Slack Web Hook URL
    let returnValue = new Array();
    let allInfoSlackSendReturnValue = sendSlack(slackURL, bodyMessage, row, lastRow, loopCount, headerMessage, commonMessage);
    returnValue.push(allInfoSlackSendReturnValue);

    let slackEmployeeEmailSendReturnValue = '';

    let emailSendReturnValue = '';

    if (sendEmailStatus === undefined || sendEmailStatus === '' || sendEmailStatus === '미발송') {
        emailSendReturnValue = sendMail(activeSheet,row, firstCol, email, emailTest, name, date);
    }

    returnValue.push(emailSendReturnValue);

    if ((lastRow - 1) < row && row < (lastRow + 1)) {
        slackEmployeeEmailSendReturnValue = slackEmployeeEmail(headerMessage, commonMessage, bodyMessage);
    }

    returnValue.push(slackEmployeeEmailSendReturnValue);

    return returnValue;
} // sendBridge(bodyMessage, detailViewURL, alarmBotVersion, activeSheet, row, lastRow, firstCol, email, emailTest, name, date, loopCount) 끝

function sendMail(activeSheet,row, firstCol, email, emailTest, name, date) {
    Logger.log('sendBridge 작동 되었습니다.');

    const emailSubject = '[꼼꼼한 친구들 - Perfect Care] 상담 / 견적 요청서 정상 접수';
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
        clientEmailSendErrorCheck = '꼼꼼한 친구들 Perfect Care 고객 상담 / 견적 요청 완료 및 감사 인사 관련 고객 Email 전송 실패 하였습니다.\n'
            + '문제 정보 : ' + err
            + '\n\n\n\n';

        Logger.log(clientEmailSendErrorCheck);

    } finally {
        Logger.log('꼼꼼한 친구들 Perfect Care 고객 상담 / 견적 요청 완료 및 감사 인사 관련 Email 전송 작업 처리 되었습니다.');

        return clientEmailSendErrorCheck;
    }
} // sendMail(activeSheet,row, firstCol, email, emailTest, name, date) 끝

function sendSlack(slackURL, bodyMessage, row, lastRow, loopCount, headerMessage, commonMessage) {

    let message = '';
    let slackErrorCheck = '';

    Logger.log(typeof loopCount);

    if (row === 2) {
        Logger.log('row === 2 검증 : ' + row === 2);
        message = headerMessage + commonMessage + bodyMessage

    } else if (lastRow === row) {
        let imsicheck = lastRow === row;
        Logger.log('(lastRow === row) 검증 : ' + imsicheck);
        message = bodyMessage + commonMessage;

    } else {
        message = bodyMessage;
    }

    let payload = {                                                                                                                                                                 // Slack으로 보낼 Message 준비
        'text' : message
    };

    let option = {                                                                                                                                                                  // Slack Web Hook 정보
        'method' : 'post',
        'contentType' : 'application/json',
        'payload' : JSON.stringify(payload)
    };

    try {
        UrlFetchApp.fetch(slackURL, option);                                                                                                                                           // Slack에 보내기기
        Logger.log('꼼꼼한 친구들 Perfect Care 고객 상담 / 견적 요청 정보를 Slack에 성공적으로 보냈습니다.');

    } catch(err) {
        slackErrorCheck = '꼼꼼한 친구들 Perfect Care 고객 상담 / 견적 요청 정보를 Slack 전송 실패 하였습니다.\n'
            + '문제 정보 : ' + err
            + '\n\n\n\n';

        Logger.log(slackErrorCheck);

    } finally {
        Logger.log('꼼꼼한 친구들 Perfect Care 고객 상담 / 견적 요청 정보 Slack 작업 처리 완료 되었습니다.');

        return slackErrorCheck;
    }
}  // sendSlack(slackURL, bodyMessage, row, lastRow, loopCount, headerMessage, commonMessage) 끝

function slackEmployeeEmail(headerMessage, commonMessage, bodyMessage) {

    let message = headerMessage + commonMessage + bodyMessage;

    let slackURL = '';

    let errorMessage = [];

    const emailSubject = '[꼼꼼한 친구들 - Perfect Care] 상담 / 견적 요청서가 접수 되었습니다.';
    let emailBody = message
        +'\n\n\n\n'
        +'본 Mail은 자동화 Bot에 의해 발송 되었습니다.';

    let employeeEmailSendErrorCheck = '';

    let employeeEmail = {
        ceo : 'abc@hanmail.net',
        coCeo : 'def@naver.com',
        cto : 'hijk@gmail.com'
    };

    let slackErrorCheck = '';

    let payload = {                                                                                                                                                                 // Slack으로 보낼 Message 준비
        'text' : message
    };

    let option = {                                                                                                                                                                  // Slack Web Hook 정보
        'method' : 'post',
        'contentType' : 'application/json',
        'payload' : JSON.stringify(payload)
    };

    for(let [key, value] of Object.entries(employeeEmail)) {                                                                                                                        // employeeEmail 객체의 값을 하나씩 꺼내 Logic 처리
        let email = value;

        try {
            GmailApp.sendEmail(email, emailSubject, emailBody);
            Logger.log('꼼꼼한 친구들 상담 / 견적 요청 접수 보고 직원에게 Email이 발송 되었어요.');

        } catch(err) {
            employeeEmailSendErrorCheck = '꼼꼼한 친구들 Perfect Care 고객 상담 / 견적 요청 접수 보고 직원 Email 전송 실패 하였습니다.\n'
                + '문제 정보 : ' + err
                + '\n\n\n\n';

            Logger.log(employeeEmailSendErrorCheck);
            errorMessage.push(employeeEmailSendErrorCheck);
        }
    }

    try {
        UrlFetchApp.fetch(slackURL, option);                                                                                                                                           // Slack에 보내기기
        Logger.log('꼼꼼한 친구들 Perfect Care 고객 상담 / 견적 요청 정보를 Slack에 성공적으로 보냈습니다.');

    } catch(err) {
        slackErrorCheck = '꼼꼼한 친구들 Perfect Care 고객 상담 / 견적 요청 정보를 Slack 전송 실패 하였습니다.\n'
            + '문제 정보 : ' + err
            + '\n\n\n\n';

        Logger.log(slackErrorCheck);
        errorMessage.push(slackErrorCheck);

    } finally {
        Logger.log('꼼꼼한 친구들 Perfect Care 고객 상담 / 견적 요청에 대한 한 건 Slack 보고 및 직원 Email 발송 작업 처리 되었습니다.');
        return errorMessage;
    }
} // slackEmployeeEmail(headerMessage, commonMessage, bodyMessage) 끝

function errorCheck(responseValue, detailViewURL, alarmBotVersion) {
    Logger.log('errorCheck 작동 되었습니다.');
    Logger.log('errorCheck 함수 매개 변수 검증 : ' + responseValue);

    errorSlackMessage = '🆘 자동화 처리 중 문제 발생하였습니다.'
        + '\n 상담 / 견적 요청 알림 Bot Version : ' + alarmBotVersion
        + '\n\n 요청서 상세 보기 URL : ' + detailViewURL + '\n'
        + 'ⓒ 2022. 주니하랑 All Rights Reserved. \n\n\n\n'
        + '각 Logic 반환 문제 내용 : ' + responseValue
        + '\n\n\n\n';

    for(idx = 0; idx < responseValue.length; idx++) {
        if (responseValue[idx] !== undefined || responseValue[idx] !== '') {
            errorSlackMessage += responseValue[idx];
        }
    }
    sendSlackError(errorSlackMessage);
} // errorCheck(responseValue, detailViewURL, alarmBotVersion) 끝

function sendSlackError(errorSlackMessage) {

    let sendSlackErrorURL = '';

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
        Logger.log('꼼꼼한 친구들 Perfect Care 고객 상담 / 견적 요청 정보를 Slack에 성공적으로 보냈습니다.');

    } catch(err) {
        Logger.log('꼼꼼한 친구들 Perfect Care 고객 상담 / 견적 요청 정보를 Slack 전송 실패 하였습니다.\n');
        Logger.log('문제 정보 : ' + err);
    } finally {
        Logger.log('꼼꼼한 친구들 Perfect Care 고객 상담 / 견적 요청 정보를 Slack 작업이 처리 되었습니다.');
    }
} // sendSlackError(errorSlackMessage) 끝