# slackbot_memoir
성장하고자 하는 사람들의 회고모임 플랫폼, 메모어 서비스를 위한 Slack Bot 프로젝트

메모어 : https://www.memoirapp.com/

※ 슬랙 API 메소드를 사용하여 제작 : https://api.slack.com/methods

<br>

## 0. requirments.txt
필요한 패키지를 한 번에 install 할 수 있도록 requirments.txt에 담아두었습니다.

<br>

## 1. excelarchive.py
기능 : 원하는 공개 및 비공개 채널(public & private channel)의 모든 메세지 내용을 불러와 엑셀 파일로 저장할 수 있습니다.

설정 사항 : 자사 Slack APP 토큰을 환경 변수로 설정하고, 토큰의 권한 설정을 마치면(Slack API 메소드 참고) 누구나 이용하실 수 있습니다.

파일 실행시 input 값 : 
- 첫 번째 줄 : 원하는 채널명을 ' '(띄어쓰기)로 구분하여 나열
- 두 번째 줄 : 시작 기간 설정(그냥 enter 시 default는 '0')
- 세 번째 줄 : 종료 기간 설정(그냥 enter 시 default는 현재 시간)
- 네 번째 줄 : 저장하고자 하는 파일의 파일명

※ 메모어 서비스 특성상, 기간 중 소속된 채널의 멤버가 작성한 reply 횟수를 집계한 엑셀 파일도 함께 생성합니다.

<br>

## 2. memoir_app.py(미완 - 제작중)
기능 : Slack에서 특정 이벤트를 인지하고 받아와 원하는 답을 해주는 챗봇입니다.

<br>

-------
작성 및 제작자 : 이강희(2kangee1@gmail.com)
