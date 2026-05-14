# All4Land Auto Test

올포랜드 해양정보 활용시스템 유지관리 업무 중 반복적으로 수행되던 OpenAPI URL 검증 작업을 자동화하기 위해 개발한 Excel 기반 테스트 도구입니다.

기존에는 Excel에 호출 코드와 예시 URL 정도만 정리되어 있었고, 실제 테스트 대상 URL은 작업자가 직접 확인하고 호출해야 했습니다.  
이에 따라 테스트 대상 URL을 Excel 기준으로 정리하고, 해당 Excel 파일을 업로드하면 각 URL을 자동으로 호출한 뒤 결과를 다시 Excel 파일로 내려받을 수 있도록 구현했습니다.

## 주요 기능

- Excel 파일 업로드 및 테스트 대상 URL 파싱
- WMS, WMS 이미지, WFS, XML, JSON URL 자동 호출
- 응답 코드, Content-Type, XML/JSON 데이터 존재 여부 검증
- WMS 이미지의 단색 여부를 검사하여 실제 레이어 표시 여부 확인
- WFS 응답은 SAX Parser 기반으로 필요한 레이어 태그 탐색
- 테스트 결과를 Excel 파일로 생성 및 다운로드

## 사용 기술

- Java
- Spring MVC
- eGovFrame
- JSP
- Maven
- Apache POI
- HttpURLConnection
- SAX Parser
- DOM Parser
- Jackson
- Commons Imaging

## 프로젝트 구조

```text
all4land_auto_test
└── auto_test
    ├── pom.xml
    └── src
        └── main
            ├── java
            │   └── egovframework
            │       └── kss
            │           └── main
            │               ├── controller
            │               ├── model
            │               └── service
            └── webapp
```

## 처리 흐름

```text
1. 테스트 대상 URL을 Excel 기준으로 정리
2. Excel 파일 업로드
3. 서버에서 Excel 데이터 파싱
4. WMS / WMS 이미지 / WFS / XML / JSON URL 자동 호출
5. 응답 형식별 유효성 검증
6. 테스트 결과 Excel 파일 생성
7. 결과 파일 다운로드
```
