// index.js

require('dotenv').config();
const puppeteer = require('puppeteer');
console.log('Loaded Puppeteer version:', puppeteer.version);

const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const moment = require('moment'); // 날짜 포맷팅을 위해 moment.js 사용

// automationHelper.js에서 필요한 함수들을 불러옵니다.
const {
    automateKisaWhois,
    automateMyLocation,
    getPoliceStation,
} = require('./automationHelper');


// --- Google API 설정 변수 ---
const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');
const TOKEN_PATH = path.join(__dirname, 'token.json');

// 필수 설정: 아래 값들을 실제 자신의 것으로 변경해야 합니다
const SPREADSHEET_ID = '18mVHOnAQtqW6IL2TOXtMENB6WpNXiIVacBWi1_-2Ey0'; // Google 스프레드시트 ID
const DB_SHEET_NAME = 'DB'; // DB 시트 이름 (예: 'DB' 또는 '경찰서DB')
const DB_RANGE = 'B:F'; // DB 시트에서 읽어올 범위 (B: 관할경찰서, C: 맨 왼쪽 값, E: 관할구역, F: 행정단위)

// 스프레드시트에서 읽어올 범위 (A열부터 K열까지, mylocation 정보 저장을 위해 범위 확장)
const RANGE = 'A2:K'; // A(IP), B(TITLE), C(COMPANY), D(POLICE), E(채증일시), F(관할경찰서 최종), G(고소장 링크), H(스크린샷 폴더 링크), I(mylocation 주소), J(오류 메시지)

// 구글 드라이브 스크린샷 저장 폴더 ID (생성 필요)
const SCREENSHOT_FOLDER_ID = '1V9cWU6x1YFMhcVFSVPpCQ6_OdslVNVhT'; // 실제 폴더 ID로 변경

// 스프레드시트 컬럼 인덱스 (0부터 시작)
const IP_ADDRESS_COLUMN_INDEX = 0; // A열
const TITLE_COLUMN_INDEX = 1; // B열
const COMPANY_COLUMN_INDEX = 2; // C열
const POLICE_STATION_COLUMN_INDEX = 3; // D열 (원래 입력된 경찰서명)
const CAPTURE_DATE_COLUMN_INDEX = 4; // E열
const FINAL_POLICE_STATION_COLUMN_INDEX = 5; // F열 (최종 결정된 경찰서명)
const DOC_LINK_COLUMN_INDEX = 6; // G열 (고소장 링크)
const SCREENSHOT_FOLDER_LINK_COLUMN_INDEX = 7; // H열 (스크린샷 폴더 링크)
const MYLOCATION_INFO_COLUMN_INDEX = 8; // I열 (mylocation 주소 정보 추가)
const ERROR_MESSAGE_COLUMN_INDEX = 9; // J열 (오류 메시지, 기존 I열에서 한 칸 밀림)

// DB 시트 컬럼 인덱스 (DB_RANGE 'B:F' 기준 0부터 시작)
const DB_POLICE_STATION_COLUMN = 0; // DB 시트 B열 (관할경찰서)
const DB_LEFTMOST_VALUE_COLUMN = 1; // DB 시트 C열 (맨 왼쪽 값)
const DB_JURISDICTION_COLUMN = 3; // DB 시트 E열 (관할구역)
const DB_ADMIN_UNIT_COLUMN = 4; // DB 시트 F열 (행정단위, automationHelper에 DB_ADMIN_UNIT_COLUMN이 정의되지 않아 추가)


// Google API 인증 및 권한 설정
const SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents'
];

/**
 * Google API 인증을 처리합니다.
 * @returns {Promise<google.auth.OAuth2>} 인증된 OAuth2 클라이언트
 */
async function authorize() {
    let credentials = {};
    if (fs.existsSync(CREDENTIALS_PATH)) {
        credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH));
    } else {
        throw new Error('credentials.json 파일을 찾을 수 없습니다.');
    }

    const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;
    const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

    if (fs.existsSync(TOKEN_PATH)) {
        oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
    } else {
        await getNewToken(oAuth2Client);
    }
    return oAuth2Client;
}

/**
 * 새로운 인증 토큰을 가져와서 저장합니다.
 * @param {google.auth.OAuth2} oAuth2Client OAuth2 클라이언트
 * @returns {Promise<void>}
 */
async function getNewToken(oAuth2Client) {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
    });
    console.log('Authorize this app by visiting this URL:', authUrl);

    // 사용자로부터 코드 입력 받기 (간단한 예시, 실제 환경에서는 웹 서버를 통해 리다이렉트 처리)
    const readline = require('readline').createInterface({
        input: process.stdin,
        output: process.stdout,
    });

    return new Promise((resolve, reject) => {
        readline.question('Enter the code from that page here: ', async (code) => {
            readline.close();
            try {
                const { tokens } = await oAuth2Client.getToken(code);
                oAuth2Client.setCredentials(tokens);
                fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
                console.log('Token stored to', TOKEN_PATH);
                resolve();
            } catch (err) {
                console.error('Error retrieving access token', err);
                reject(err);
            }
        });
    });
}

/**
 * 메인 자동화 함수
 */
async function main() {
    let browser;
    let pageKisa;
    let pageMyLocation;

    try {
        const auth = await authorize();
        const sheets = google.sheets({ version: 'v4', auth });
        const drive = google.drive({ version: 'v3', auth });

        // 경찰서 DB 정보 로드
        const policeStations = await getPoliceStation(sheets, SPREADSHEET_ID, DB_SHEET_NAME, DB_RANGE);

        // 스프레드시트에서 IP 주소 목록 읽기
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: RANGE,
        });
        let rows = response.data.values;

        if (!rows || rows.length === 0) {
            console.log('스프레드시트에서 데이터를 찾을 수 없습니다.');
            return;
        }

        browser = await puppeteer.launch({ headless: false }); // 디버깅을 위해 headless: false 설정
        pageKisa = await browser.newPage();
        await pageKisa.setViewport({ width: 1280, height: 720 });
        console.log("Puppeteer 브라우저 및 KISA WHOIS 페이지 초기화 완료.");

        // D열(POLICE_STATION_COLUMN_INDEX)이 비어있는 첫 번째 행을 찾습니다.
        let startIndex = 0;
        let foundEmptyDColumn = false;
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            // D열이 비어있는지 확인. trim()을 사용하여 공백만 있는 경우도 비어있는 것으로 간주.
            if (!row[POLICE_STATION_COLUMN_INDEX] || String(row[POLICE_STATION_COLUMN_INDEX]).trim() === '') {
                startIndex = i;
                foundEmptyDColumn = true;
                console.log(`D열이 비어있는 첫 번째 행 (A${startIndex + 2}부터 처리 시작)을 찾았습니다.`);
                break;
            }
        }

        if (!foundEmptyDColumn) {
            console.log('D열이 비어있는 행을 찾을 수 없습니다. 스크립트를 종료합니다.');
            return;
        }

        // 각 행 처리 (startIndex부터 시작)
        for (let i = startIndex; i < rows.length; i++) {
            const currentRow = i + 2; // A2부터 시작하므로 실제 스프레드시트 행 번호는 +2
            const row = rows[i];
            const ipAddress = row[IP_ADDRESS_COLUMN_INDEX];

            // IP 주소가 없으면 건너뛰기
            if (!ipAddress) {
                console.log(`${currentRow}행에 IP 주소가 없습니다. 건너뜁니다.`);
                continue;
            }

            // D열이 이미 채워져 있고, G열(고소장 링크)이나 J열(오류 메시지)이 비어있지 않은 경우 건너뛰기
            // D열이 비어있어야만 처리하도록 조건 추가
            if (row[POLICE_STATION_COLUMN_INDEX] && String(row[POLICE_STATION_COLUMN_INDEX]).trim() !== '') {
                console.log(`${currentRow}행 IP (${ipAddress}): D열이 이미 채워져 있습니다. 건너뜁니다.`);
                continue;
            }

            console.log(`\n--- ${currentRow}행 IP 주소 처리 시작: ${ipAddress} ---`);

            let kisaScreenshotFileId = null;
            let myLocationScreenshotFileId = null;
            let myLocationInfo = null;
            let kisaResultStatus = '정상'; // KISA WHOIS 결과 상태 초기화 (정상, 해외, 모바일, 오류)
            let finalPoliceStationForDColumn = ''; // D열에 넣을 최종 경찰서명
            // F열에 기록할 내용. 기본적으로는 기존 F열 값 유지.
            let fColumnContent = row[FINAL_POLICE_STATION_COLUMN_INDEX] || '';

            try {
                // 1. KISA WHOIS 자동화
                const kisaResult = await automateKisaWhois(pageKisa, ipAddress, drive, SCREENSHOT_FOLDER_ID);
                kisaScreenshotFileId = kisaResult.screenshotFileId;
                kisaResultStatus = kisaResult.kisaResultStatus;

                if (kisaScreenshotFileId) {
                    console.log(`KISA WHOIS 스크린샷 파일 ID: ${kisaScreenshotFileId}`);
                } else {
                    console.warn(`WARN: KISA WHOIS 스크린샷을 가져오지 못했습니다.`);
                }

                // 2. KISA WHOIS 결과에 따라 mylocation.co.kr 자동화 조건부 실행
                // '해외' 또는 '모바일'로 판정된 경우 mylocation.co.kr 절차 생략
                if (kisaResultStatus === '해외' || kisaResultStatus === '모바일') {
                    const overseasIpText = '해외IP.';
                    console.log(`KISA WHOIS 결과가 '${kisaResultStatus}'이므로 mylocation.co.kr 절차를 생략하고 '${overseasIpText}'로 D열과 I열에 표시합니다.`);
                    
                    myLocationInfo = overseasIpText; // I열에 기록할 내용
                    // fColumnContent는 변경하지 않고 기존 F열 값을 유지합니다.
                    finalPoliceStationForDColumn = overseasIpText; // D열에 기록할 내용

                    console.log('해외/모바일 처리 후 다음 행 실행 전 5초 대기...');
                    await pageKisa.waitForTimeout(5000); // 5초 대기 추가
                } else { // KISA WHOIS 결과가 '정상'이거나 다른 경우에만 mylocation 진행
                    pageMyLocation = await browser.newPage();
                    await pageMyLocation.setViewport({ width: 1280, height: 720 });
                    const myLocationResult = await automateMyLocation(pageMyLocation, ipAddress, drive, SCREENSHOT_FOLDER_ID);
                    myLocationScreenshotFileId = myLocationResult.screenshotFileId;
                    myLocationInfo = myLocationResult.locationInfo;

                    if (myLocationScreenshotFileId) {
                        console.log(`mylocation.co.kr 스크린샷 파일 ID: ${myLocationScreenshotFileId}`);
                    } else {
                        console.warn(`WARN: mylocation.co.kr 스크린샷을 가져오지 못했습니다.`);
                    }
                    if (myLocationInfo) {
                        console.log(`mylocation.co.kr 위치 정보: ${myLocationInfo}`);
                    } else {
                        console.warn(`WARN: mylocation.co.kr 위치 정보를 가져오지 못했습니다.`);
                    }
                    await pageMyLocation.close();
                    console.log("mylocation.co.kr 페이지 닫힘.");
                }


                // 3. 경찰서 정보 매칭 (I열 값 기반) - myLocationInfo가 정상일 경우에만 진행
                // '해외IP.' 문구가 있을 경우에는 경찰서 매칭을 시도하지 않습니다.
                if (myLocationInfo && myLocationInfo !== '해외IP.' && policeStations.length > 0) {
                    // 1. myLocationInfo 전처리: 괄호와 그 안의 내용 제거 및 공백 정규화
                    let processedLocationInfo = myLocationInfo
                        .replace(/\s*\(.*?\)\s*$/, '') // 괄호와 그 안의 내용 제거
                        .trim()                         // 앞뒤 공백 제거
                        .replace(/\s+/g, ' ');          // 여러 개의 공백을 하나로 줄임

                    console.log(`전처리된 위치 정보: ${processedLocationInfo}`);

                    let matchedStation = null;

                    // 2. '읍', '면', '동', '리'로 끝나는 가장 구체적인 행정 단어 추출 및 매칭 시도
                    const adminUnitRegex = /(\S+(읍|면|동|리))\s*(\d.*)?$/;
                    const match = processedLocationInfo.match(adminUnitRegex);

                    let specificAdminUnit = null;
                    if (match && match[1]) {
                        specificAdminUnit = match[1];
                        console.log(`추출된 구체적 행정단위: ${specificAdminUnit}`);
                    }

                    if (specificAdminUnit) {
                        matchedStation = policeStations.find(stationRow =>
                            stationRow[DB_ADMIN_UNIT_COLUMN] &&
                            (stationRow[DB_ADMIN_UNIT_COLUMN].includes('읍') ||
                             stationRow[DB_ADMIN_UNIT_COLUMN].includes('면') ||
                             stationRow[DB_ADMIN_UNIT_COLUMN].includes('동') ||
                             stationRow[DB_ADMIN_UNIT_COLUMN].includes('리')) &&
                            stationRow[DB_JURISDICTION_COLUMN] &&
                            stationRow[DB_JURISDICTION_COLUMN].includes(specificAdminUnit)
                        );

                        if (matchedStation) {
                            console.log(`'읍/면/동/리' 기준으로 매칭 성공: ${specificAdminUnit}`);
                        }
                    }

                    // 3. '시', '군', '구' 단위 매칭 시도 (위에서 매칭 안 된 경우에만 실행)
                    if (!matchedStation) {
                        const cityCountyGuUnits = ['구', '군', '시'];
                        for (const unit of cityCountyGuUnits) {
                            const regex = new RegExp(`(\\S+${unit})`);
                            const guMatch = processedLocationInfo.match(regex);
                            
                            if (guMatch && guMatch[1]) {
                                const foundUnit = guMatch[1];
                                matchedStation = policeStations.find(stationRow =>
                                    stationRow[DB_ADMIN_UNIT_COLUMN] &&
                                    (stationRow[DB_ADMIN_UNIT_COLUMN].includes('시') ||
                                     stationRow[DB_ADMIN_UNIT_COLUMN].includes('군') ||
                                     stationRow[DB_ADMIN_UNIT_COLUMN].includes('구')) &&
                                    stationRow[DB_JURISDICTION_COLUMN] &&
                                    stationRow[DB_JURISDICTION_COLUMN].includes(foundUnit)
                                );
                                if (matchedStation) {
                                    console.log(`'${unit}' 기준으로 매칭 성공: ${foundUnit}`);
                                    break;
                                }
                            }
                        }
                    }

                    if (matchedStation) {
                        finalPoliceStationForDColumn = matchedStation[DB_POLICE_STATION_COLUMN];
                        console.log(`최종 매칭된 경찰서 (D열): ${finalPoliceStationForDColumn}`);
                    } else {
                        console.log(`위치 정보 (${myLocationInfo})와 일치하는 경찰서를 찾을 수 없습니다.`);
                        // 매칭에 실패하면 D열은 비워둡니다.
                        finalPoliceStationForDColumn = '';
                    }
                } else if (myLocationInfo !== '해외IP.') { // myLocationInfo가 해외IP.가 아니고 매칭 시도도 안 된 경우 (경찰서DB 비어있거나 등)
                     console.log('mylocation 정보가 있으나 경찰서 DB와 매칭할 수 없습니다. D열은 비워둡니다.');
                     finalPoliceStationForDColumn = '';
                }


                // 4. 스프레드시트 업데이트
                const updateValues = [
                    row[IP_ADDRESS_COLUMN_INDEX], // A열 (IP)
                    row[TITLE_COLUMN_INDEX], // B열 (TITLE)
                    row[COMPANY_COLUMN_INDEX], // C열 (COMPANY)
                    finalPoliceStationForDColumn, // D열 (경찰서): '해외IP.' 또는 매칭된 경찰서 (매칭 실패 시 빈 값)
                    row[CAPTURE_DATE_COLUMN_INDEX] || '', // E열 (채증일시): 기존 값 유지 또는 빈 문자열
                    fColumnContent, // F열 (관할경찰서 최종): 변경 없음 (기존 F열 값 유지)
                    '', // G열 (고소장 링크): 비워둠
                    '', // H열 (스크린샷 폴더 링크): 비워둠
                    myLocationInfo || '', // I열 (mylocation 주소 정보): '해외IP.' 또는 실제 주소
                    '' // J열 (오류 메시지)
                ];

                const updateRange = `${String.fromCharCode(65 + IP_ADDRESS_COLUMN_INDEX)}${currentRow}:${String.fromCharCode(65 + ERROR_MESSAGE_COLUMN_INDEX)}${currentRow}`;
                await sheets.spreadsheets.values.update({
                    spreadsheetId: SPREADSHEET_ID,
                    range: updateRange,
                    valueInputOption: 'RAW',
                    resource: {
                        values: [updateValues],
                    },
                });
                console.log(`스프레드시트 ${updateRange} 업데이트 완료.`);

            } catch (err) {
                console.error(`${currentRow}행 IP (${ipAddress}) 처리 중 오류 발생:`, err.message);
                const errorRange = `${String.fromCharCode(65 + ERROR_MESSAGE_COLUMN_INDEX)}${currentRow}`;
                await sheets.spreadsheets.values.update({
                    spreadsheetId: SPREADSHEET_ID,
                    range: errorRange,
                    valueInputOption: 'RAW',
                    resource: {
                        values: [[`ERROR: ${err.message}`]],
                    },
                });
                console.log(`스프레드시트 ${errorRange} 오류 메시지 기록 완료.`);

                // Puppeteer 관련 치명적인 오류 발생 시 브라우저를 닫고 전체 스크립트를 종료
                if (err.message.includes('Protocol error') || err.message.includes('No target with given id found') || err.message.includes('Attempted to use detached Frame')) {
                    console.error("치명적인 Puppeteer 오류 발생. 남은 절차를 중단합니다.");
                    if (browser) {
                        try {
                            // 열려있는 모든 페이지를 강제로 닫기 시도
                            const pages = await browser.pages();
                            for (const p of pages) {
                                if (!p.isClosed()) {
                                    await p.close().catch(e => console.warn(`WARN: 페이지 닫기 실패 (오류: ${e.message})`));
                                }
                            }
                            await browser.close();
                            console.log("Puppeteer 브라우저 종료 (치명적 오류로 인한 강제 종료).");
                        } catch (closeErr) {
                            console.error(`ERROR: 브라우저 종료 중 오류 발생: ${closeErr.message}`);
                        }
                    }
                    return; // main 함수를 여기서 종료하여 전체 스크립트 실행을 멈춥니다.
                }
            }
        }
    } catch (error) {
        console.error("메인 자동화 프로세스 실패:", error);
    } finally {
        // Puppeteer 브라우저가 아직 열려있다면 종료합니다.
        // 치명적인 Puppeteer 오류로 인해 이미 닫혔을 수도 있으므로, 연결 상태를 확인합니다.
        if (browser && browser.isConnected()) {
            try {
                await browser.close();
                console.log("Puppeteer 브라우저 종료.");
            } catch (e) {
                console.warn(`WARN: 최종 브라우저 종료 실패 (오류: ${e.message})`);
            }
        }
        // 페이지 객체들도 명시적으로 닫아줍니다.
        if (pageKisa && !pageKisa.isClosed()) {
            try {
                await pageKisa.close();
                console.log("KISA WHOIS 페이지 닫힘.");
            } catch (e) { console.warn(`WARN: KISA WHOIS 페이지 닫기 실패 (오류: ${e.message})`); }
        }
        if (pageMyLocation && !pageMyLocation.isClosed()) {
            try {
                await pageMyLocation.close();
                console.log(`mylocation.co.kr 페이지 닫힘.`);
            } catch (e) { console.warn(`WARN: mylocation.co.kr 페이지 닫기 실패 (오류: ${e.message})`); }
        }
    }
}

main();
