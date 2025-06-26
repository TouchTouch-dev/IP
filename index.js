// index.js

require('dotenv').config();
const puppeteer = require('puppeteer');
console.log('Loaded Puppeteer version:', puppeteer.version);

const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const moment = require('moment'); // 날짜 포맷팅을 위해 moment.js 사용

// 설정 파일(config.js) 불러오기
const config = require('./config');

// 자동화 보조 함수 파일(automationHelper.js)에서 필요한 함수들을 불러오기
const {
    automateMyLocation,
} = require('./automationHelper');


// Google API 설정 파일 경로 (config.js의 파일 이름을 사용)
const CREDENTIALS_PATH = path.join(__dirname, config.CREDENTIALS_FILE_NAME);
const TOKEN_PATH = path.join(__dirname, config.TOKEN_FILE_NAME);


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
        scope: config.SCOPES, // config.js에서 SCOPES 사용
    });
    console.log('Authorize this app by visiting this URL:', authUrl);

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
                fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens)); // TOKEN_PATH 사용
                console.log('Token stored to', TOKEN_PATH); // TOKEN_PATH 사용
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
    let pageMyLocation; // mylocation 페이지 하나만 필요

    // --- 글로벌 종료 핸들러 설정 ---
    // 이 핸들러들은 Ctrl+C (SIGINT), 프로세스 종료 (SIGTERM), 처리되지 않은 Promise 거부 (unhandledRejection) 시 호출됩니다.
    // main 함수 내의 finally 블록과 함께, 브라우저가 확실히 닫히도록 하는 안전 장치입니다.
    const cleanupAndExit = async (signal) => {
        console.log(`\n${signal} 신호 수신. 브라우저 종료 및 스크립트 종료 시도...`);
        if (browser && browser.isConnected()) {
            try {
                await browser.close();
                console.log("Puppeteer 브라우저 정상 종료.");
            } catch (e) {
                console.error(`ERROR: 브라우저 종료 중 오류 발생: ${e.message}`);
            }
        }
        process.exit(1); // 오류 코드로 종료
    };

    process.on('SIGINT', cleanupAndExit.bind(null, 'SIGINT'));
    process.on('SIGTERM', cleanupAndExit.bind(null, 'SIGTERM'));
    process.on('unhandledRejection', async (reason, promise) => {
        console.error('Unhandled Rejection at:', promise, 'reason:', reason);
        await cleanupAndExit('UNHANDLED_REJECTION');
    });
    // ----------------------------

    try {
        const auth = await authorize();
        const sheets = google.sheets({ version: 'v4', auth });
        const drive = google.drive({ version: 'v3', auth });

        // 스프레드시트에서 IP 주소 목록 읽기 (config.js의 값 사용)
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: config.SPREADSHEET_ID,
            range: config.RANGE,
        });
        let rows = response.data.values;

        if (!rows || rows.length === 0) {
            console.log('스프레드시트에서 데이터를 찾을 수 없습니다.');
            return;
        }

        browser = await puppeteer.launch({ headless: false }); // 브라우저가 보이도록 headless: false 설정
        console.log("Puppeteer 브라우저 초기화 완료.");
        
        // I열(MYLOCATION_ADDRESS_COLUMN_INDEX)이 공백인 첫 번째 행을 찾는 로직은 제거하고,
        // 모든 행을 순회하면서 I열에 값이 없으면 처리하도록 변경합니다.
        for (let i = 0; i < rows.length; i++) {
            const currentRow = i + 2; // A2부터 시작하므로 실제 스프레드시트 행 번호는 +2
            const row = rows[i];
            const ipAddress = row[config.IP_ADDRESS_COLUMN_INDEX];
            const myLocationAddress = row[config.MYLOCATION_ADDRESS_COLUMN_INDEX]; // I열 값

            // IP 주소가 없거나, 이미 MYLOCATION 주소(I열)가 채워져 있으면 건너뛰기
            if (!ipAddress) {
                console.log(`${currentRow}행에 IP 주소가 없습니다. 건너킵니다.`);
                continue;
            }
            if (myLocationAddress && myLocationAddress.trim() !== '') {
                console.log(`${currentRow}행 (IP: ${ipAddress})은 이미 MYLOCATION 주소(I열)가 채워져 있어 건너킵니다.`);
                continue;
            }

            console.log(`\n--- ${currentRow}행 IP 주소 처리 시작: ${ipAddress} ---`);

            let myLocationScreenshotFileId = null;
            let myLocationInfo = null;
            
            try {
                // mylocation.co.kr 자동화
                pageMyLocation = await browser.newPage();
                await pageMyLocation.setViewport({ width: 1280, height: 720 });
                await pageMyLocation.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36');
                await pageMyLocation.setBypassCSP(true);
                await pageMyLocation.setCacheEnabled(false);

                const myLocationResult = await automateMyLocation(pageMyLocation, ipAddress, drive, config.SCREENSHOT_FOLDER_ID);
                myLocationScreenshotFileId = myLocationResult.screenshotFileId;
                myLocationInfo = myLocationResult.locationInfo;

                // myLocationInfo가 공백이거나 유효하지 않으면 '해외IP'로 기록
                if (!myLocationInfo || typeof myLocationInfo !== 'string' || myLocationInfo.trim() === '' || myLocationInfo.includes('찾을 수 없습니다')) {
                    console.log(`mylocation.co.kr에서 ${ipAddress}에 대한 주소 검색 결과가 없습니다. '해외IP'로 기록합니다.`);
                    myLocationInfo = '해외IP';
                }

                if (myLocationScreenshotFileId) {
                    console.log(`mylocation.co.kr 스크린샷 파일이 Google Drive에 업로드되었습니다. 파일 ID: ${myLocationScreenshotFileId}`); // 메시지 변경
                } else {
                    console.warn(`WARN: mylocation.co.kr 스크린샷을 가져오지 못했습니다.`);
                }
                if (myLocationInfo) {
                    console.log(`mylocation.co.kr 위치 정보: ${myLocationInfo}`);
                } else {
                    console.warn(`WARN: mylocation.co.kr 위치 정보를 가져오지 못했습니다.`);
                }
                
                // 스프레드시트 업데이트: I열(mylocation 주소), J열(오류 메시지)만 업데이트
                // H열(스크린샷 ID) 업데이트는 제거됨
                const updateRangeForIJ = `시트1!${String.fromCharCode(65 + config.MYLOCATION_ADDRESS_COLUMN_INDEX)}${currentRow}:${String.fromCharCode(65 + config.ERROR_MESSAGE_COLUMN_INDEX)}${currentRow}`;
                await sheets.spreadsheets.values.update({
                    spreadsheetId: config.SPREADSHEET_ID,
                    range: updateRangeForIJ,
                    valueInputOption: 'RAW',
                    resource: {
                        values: [[
                            myLocationInfo, // I열: 이제 '해외IP' 또는 실제 주소
                            '' // J열 (성공 시는 빈 값)
                        ]],
                    },
                });
                console.log(`스프레드시트 ${updateRangeForIJ} 업데이트 완료.`);

                // 페이지 닫기
                await pageMyLocation.close();
                console.log("mylocation.co.kr 페이지 닫힘.");

            } catch (err) {
                console.error(`${currentRow}행 IP (${ipAddress}) 처리 중 오류 발생:`, err.message);
                const errorRangeJ = `시트1!${String.fromCharCode(65 + config.ERROR_MESSAGE_COLUMN_INDEX)}${currentRow}`;
                await sheets.spreadsheets.values.update({
                    spreadsheetId: config.SPREADSHEET_ID,
                    range: errorRangeJ,
                    valueInputOption: 'RAW',
                    resource: {
                        values: [[`ERROR: ${err.message}`]],
                    },
                });
                console.log(`스프레드시트 ${errorRangeJ} 오류 메시지 기록 완료.`);

                // Puppeteer 관련 치명적인 오류 발생 시 브라우저를 닫고 전체 스크립트를 즉시 종료
                if (err.message.includes('Protocol error') || err.message.includes('No target with given id found') || err.message.includes('Attempted to use detached Frame') || err.message.includes('Execution context was destroyed') || err.message.includes('Navigating frame was detached')) {
                    console.error("치명적인 Puppeteer 오류 발생. 스크립트를 즉시 중단합니다."); // 메시지 변경
                    if (browser) {
                        try {
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

                // 일반 오류의 경우, 현재 페이지를 닫으려고 시도하고 다음 IP로 진행 (이전 로직 유지)
                if (pageMyLocation && !pageMyLocation.isClosed()) {
                    try {
                        await pageMyLocation.close();
                        console.log(`현재 mylocation.co.kr 페이지 닫힘 (오류 발생).`);
                    } catch (closePageErr) {
                        console.warn(`WARN: 현재 페이지 닫기 실패 (오류: ${closePageErr.message})`);
                    }
                }
            } finally {
                // 이 finally 블록은 현재 IP 처리 try/catch에 대한 것이며,
                // 페이지 닫기 로직은 위에 catch 블록에서 처리되었으므로 여기서는 추가 작업이 필요 없습니다.
            }
        }
    } catch (error) {
        console.error("메인 자동화 프로세스 실패:", error);
        // 최상위 예외는 글로벌 unhandledRejection 핸들러나 아래 finally 블록에서 브라우저 종료를 시도할 것입니다.
    } finally {
        // Puppeteer 브라우저가 아직 열려있고 연결되어 있다면 종료합니다.
        // 이는 최상위 오류나 모든 IP 처리가 완료된 후에도 브라우저를 확실히 닫는 역할을 합니다.
        if (browser && browser.isConnected()) {
            try {
                await browser.close();
                console.log("Puppeteer 브라우저 종료.");
            } catch (e) {
                console.warn(`WARN: 최종 브라우저 종료 실패 (오류: ${e.message})`);
            }
        }
    }
}

main();
