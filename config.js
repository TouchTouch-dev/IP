// config.js

module.exports = {
    // Google Sheets API 관련 설정
    SPREADSHEET_ID: '18mVHOnAQtqW6IL2TOXtMENB6WpNXiIVacBWi1_-2Ey0', // 실제 스프레드시트 ID로 변경하세요.
    RANGE: '시트1!A2:J', // Sheet1에서 데이터를 읽어올 범위 (A열부터 J열까지)

    // Google API 인증 관련 설정
    CREDENTIALS_FILE_NAME: 'credentials.json', // Google API OAuth2 클라이언트 인증 정보 파일
    TOKEN_FILE_NAME: 'token.json',             // Google API 액세스 토큰 파일
    SCOPES: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file'],

    // 스크린샷 저장 관련 설정
    TEMP_SCREENSHOT_DIR_NAME: 'screenshots', // 임시 스크린샷 파일 저장 디렉토리
    SCREENSHOT_FOLDER_ID: '1V9cWU6x1YFMhcVFSVPpCQ6_OdslVNVhT', // 스크린샷을 저장할 Google Drive 폴더 ID로 변경하세요.

    // Sheet1의 컬럼 인덱스 (0부터 시작)
    IP_ADDRESS_COLUMN_INDEX: 0,         // A열
    TITLE_COLUMN_INDEX: 1,              // B열
    COMPANY_COLUMN_INDEX: 2,            // C열
    POLICE_STATION_COLUMN_INDEX: 3,     // D열 (관할 경찰서)
    CAPTURE_DATE_COLUMN_INDEX: 4,       // E열 (채증일시)
    FINAL_POLICE_STATION_COLUMN_INDEX: 5, // F열 (관할경찰서 최종)
    LAWSUIT_LINK_COLUMN_INDEX: 6,       // G열 (고소장 링크 - 이제 비워집니다.)
    SCREENSHOT_FOLDER_LINK_COLUMN_INDEX: 7, // H열 (mylocation 스크린샷 ID가 저장됩니다.)
    MYLOCATION_ADDRESS_COLUMN_INDEX: 8, // I열 (mylocation 주소 정보)
    ERROR_MESSAGE_COLUMN_INDEX: 9,      // J열 (오류 메시지)

    // mylocation.co.kr 웹사이트 관련 셀렉터
    TARGET_URL_MYLOCATION: 'https://www.mylocation.co.kr/',
    IP_INPUT_SELECTOR_MYLOCATION: '#txtAddr', // IP 입력창 셀렉터 (사용자 피드백 반영)
    SUBMIT_BUTTON_SELECTOR_MYLOCATION: '#btnAddr2', // 제출 버튼 셀렉터 (사용자 피드백 반영)
    MYLOCATION_ADDRESS_SELECTOR: '#lbAddr', // mylocation.co.kr 위치 정보 결과 셀렉터

    // DB_Back 시트 관련 설정 (경찰서 정보 매핑용)
    DB_SHEET_NAME_BACK: 'DB_Back', // DB_Back 시트의 이름
    DB_RANGE_BACK: 'A:G', // DB_Back 시트에서 경찰서 정보를 가져올 범위
    DB_BACK_POLICE_STATION_COLUMN: 0, // DB_Back 시트의 A열 (경찰서명)
    DB_BACK_ADMIN_UNIT_C_COLUMN: 2, // DB_Back 시트의 C열 (시/도)
    DB_BACK_ADMIN_UNIT_E_COLUMN: 4, // DB_Back 시트의 E열 (구/군)
    DB_BACK_ADMIN_UNIT_F_COLUMN: 5, // DB_Back 시트의 F열 (읍/면/동/리)
    DB_BACK_JURISDICTION_COLUMN: 6, // DB_Back 시트의 G열 (관할 행정단위)
};
