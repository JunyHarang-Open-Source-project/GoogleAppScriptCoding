# 🚀 Source Code 소개

---
기깔나는 사람들 Sub Project에서 꼼꼼한 친구들 - Perfect Care에서 이용할
구글 설문지를 통해 고객의 요청 정보를 받고, 그것을 Slack과 Email 등을 통해 자동화 하는 Code 입니다 :)

<br><br>

## 주의

---
본 Repository Source Code는 내려받기(Download) 등을 통해 바로 실행할 수 있는 Program 이 아니며,
Google App Script를 이용하면서 사용할 수 있는 Srouce Code 입니다.

## CAUTION

---
This Repository Source Code is not a program that can be directly executed through download, etc.
Srouce Code that can be used while using Google App Script.

## 변경 사항(Update Information)

---
* Version 1.3.5 - 2022년 11월 07
  * Slack 추가 기능
    * 하나의 채널에는 구글 스프레드 시트에 모든 정보를 보여주고(sendSlack(slackURL, bodyMessage, row, lastRow, loopCount, headerMessage, commonMessage))
      또 다른 채널에는 최신 정보 한 건만 보여준다. (slackEmployeeEmail(headerMessage, commonMessage, bodyMessage))
  * 직원 Email 발송 기능
    * Slack에서 받을 수 있는 한 건에 대한 정보를 지정한 직원 Email로도 발송할 수 있게 처리 (slackEmployeeEmail(headerMessage, commonMessage, bodyMessage))
  * Error Slack Send 기능 Bug Fix
    * 아무 이상이 없는데도 Slack으로 문제 알림이 오는 문제 해결 (mainFunction(), errorCheck(responseValue, detailViewURL, alarmBotVersion))

## 기깔나는 사람들 크루 모집 공고

![](https://img1.daumcdn.net/thumb/R1280x0/?scode=mtistory2&fname=https%3A%2F%2Fk.kakaocdn.net%2Fdn%2Fdbrw12%2FbtrQqmSAmQ0%2FKB9EfCFR13MOruUYhkdSGk%2Fimg.jpg)

[크루 모집 공고](https://productive-ornament-cad.notion.site/ff54d02eadd346b488c5e761414bf87f)

## 👨‍👨‍👧‍👧 참여자

| 이름     | Blog                            | Instagram                             |
| ---------- | --------------------------------- | --------------------------------------- |
| 주니하랑 | https://junyharang.tistory.com/ | https://www.instagram.com/junyharang/ |

<br><br><br>

## Project Code 정리

https://junyharang.tistory.com/351

<br><br><br>