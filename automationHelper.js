// automationHelper.js

const path = require('path');
const fs = require('fs');
const moment = require('moment'); // 날짜 포맷팅을 위해 moment.js 사용

// --- 첫 번째 자동화 대상 웹사이트 설정 변수 (KISA WHOIS 서비스) ---
const TARGET_URL_KISA = 'https://xn--c79as89aj0e29b77z.xn--3e0b707e/kor/whois/whois.jsp'; // KISA WHOIS 주소
const IP_INPUT_SELECTOR_KISA = '#sWord'; // KISA WHOIS IP 입력 필드 셀렉터
const SUBMIT_BUTTON_SELECTOR_KISA = 'a[href="javascript:whois();"]'; // KISA WHOIS 제출 버튼 셀렉터


// --- 두 번째 자동화 대상 웹사이트 설정 변수 (mylocation.co.kr) ---
const TARGET_URL_MYLOCATION = 'https://www.mylocation.co.kr/'; // mylocation.co.kr URL
const IP_INPUT_SELECTOR_MYLOCATION = '#txtAddr'; // mylocation.co.kr IP 입력 필드 셀렉터
const SUBMIT_BUTTON_SELECTOR_MYLOCATION = '#btnAddr2'; // mylocation.co.kr 제출 버튼 셀렉터
const MYLOCATION_ADDRESS_SELECTOR = '#lbAddr'; // mylocation.co.kr에서 주소 텍스트를 포함하는 요소의 셀렉터.

// 스크린샷 임시 저장 폴더 경로 설정
const TEMP_SCREENSHOT_DIR = path.join(__dirname, 'tempscreenshot');


/**
 * 주어진 페이지에서 스크린샷을 찍고 Google Drive에 업로드합니다.
 * @param {Page} page Puppeteer 페이지 객체
 * @param {string} ipAddress 현재 처리 중인 IP 주소
 * @param {object} drive Google Drive API 클라이언트
 * @param {string} folderId 이미지를 저장할 Google Drive 폴더 ID
 * @param {string} screenshotNameType 스크린샷 파일명 유형 ('whois' 또는 'mylocation')
 * @returns {Promise<string|null>} 업로드된 파일의 ID, 또는 실패 시 null
 */
async function takeScreenshotAndUpload(page, ipAddress, drive, folderId, screenshotNameType) {
    // 임시 스크린샷 폴더가 없으면 생성
    if (!fs.existsSync(TEMP_SCREENSHOT_DIR)) {
        fs.mkdirSync(TEMP_SCREENSHOT_DIR, { recursive: true });
        console.log(`임시 스크린샷 폴더 생성: ${TEMP_SCREENSHOT_DIR}`);
    }

    const timestamp = moment().format('YYYYMMDD_HHmmss');
    const screenshotFileName = `${screenshotNameType}_${ipAddress.replace(/\./g, '_')}_${timestamp}.png`;
    const screenshotPath = path.join(TEMP_SCREENSHOT_DIR, screenshotFileName);

    try {
        await page.screenshot({ path: screenshotPath, fullPage: true });
        console.log(`스크린샷 임시 저장 완료: ${screenshotPath}`);

        const fileMetadata = {
            'name': screenshotFileName,
            'parents': [folderId],
        };
        const media = {
            mimeType: 'image/png',
            body: fs.createReadStream(screenshotPath),
        };

        const driveResponse = await drive.files.create({
            resource: fileMetadata,
            media: media,
            fields: 'id, webViewLink',
        });

        console.log(`Google Drive에 '${screenshotNameType}' 스크린샷 업로드 완료. 파일 ID: ${driveResponse.data.id}`);
        // 로컬에 저장된 스크린샷 파일 삭제 (선택 사항)
        fs.unlinkSync(screenshotPath);
        console.log(`로컬 임시 스크린샷 파일 삭제 완료: ${screenshotPath}`);
        return driveResponse.data.id;
    } catch (error) {
        console.error(`ERROR: 스크린샷 촬영 및 Google Drive 업로드 실패 (${screenshotNameType}):`, error);
        if (fs.existsSync(screenshotPath)) {
            fs.unlinkSync(screenshotPath);
            console.log(`오류 발생 후 로컬 임시 스크린샷 파일 정리 완료: ${screenshotPath}`);
        }
        return null;
    }
}


/**
 * KISA WHOIS 서비스에서 IP 주소를 조회하고 스크린샷을 찍습니다.
 * '국내에서 관리되는 IP가 아닙니다.' 또는 '잘 남았어' 문구 감지 로직 추가.
 * @param {Page} page Puppeteer 페이지 객체
 * @param {string} ipAddress 조회할 IP 주소
 * @param {object} drive Google Drive API 클라이언트
 * @param {string} folderId 이미지를 저장할 Google Drive 폴더 ID
 * @returns {Promise<{screenshotFileId: string|null, kisaResultStatus: string}>} KISA WHOIS 스크린샷 파일 ID 및 결과 상태 ('정상', '해외', '모바일', '오류')
 */
async function automateKisaWhois(page, ipAddress, drive, folderId) {
    console.log(`KISA WHOIS 자동화 시작: ${ipAddress}`);
    let kisaResultStatus = '정상';
    let screenshotFileId = null;

    try {
        await page.goto(TARGET_URL_KISA, { waitUntil: 'networkidle0', timeout: 60000 });
        await page.waitForSelector(IP_INPUT_SELECTOR_KISA, { timeout: 10000 });
        await page.type(IP_INPUT_SELECTOR_KISA, ipAddress);

        await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 60000 }),
            page.click(SUBMIT_BUTTON_SELECTOR_KISA),
        ]);

        console.log(`KISA WHOIS 결과 페이지 로딩 후 3초 대기...`);
        await new Promise(resolve => setTimeout(resolve, 3000));

        const pageContent = await page.content(); // 페이지의 전체 HTML을 가져옵니다.

        // '국내에서 관리되는 IP가 아닙니다.' 또는 '잘 남았어' 문구 감지
        if (pageContent.includes('국내에서 관리되는 IP가 아닙니다.') || pageContent.includes('잘 남았어')) {
            console.log(`KISA WHOIS에서 '국내에서 관리되는 IP가 아닙니다.' 또는 '잘 남았어' 문구 감지. 해외 IP로 처리.`);
            kisaResultStatus = '해외';
        }
        // 이 외에 KISA WHOIS 결과에 따라 '모바일' 등으로 판단하는 로직이 필요하다면 여기에 추가

        screenshotFileId = await takeScreenshotAndUpload(page, ipAddress, drive, folderId, 'whois_kisa');
        return { screenshotFileId, kisaResultStatus };
    } catch (error) {
        console.error(`ERROR: KISA WHOIS 자동화 중 오류 발생 (${ipAddress}):`, error);
        kisaResultStatus = '오류';
        throw error; // 오류를 다시 던져서 main 함수에서 처리할 수 있도록 합니다.
    }
}


/**
 * mylocation.co.kr 서비스에서 IP 주소를 조회하고 스크린샷을 찍습니다.
 * @param {Page} page Puppeteer 페이지 객체
 * @param {string} ipAddress 조회할 IP 주소
 * @param {object} drive Google Drive API 클라이언트
 * @param {string} folderId 이미지를 저장할 Google Drive 폴더 ID
 * @returns {Promise<{screenshotFileId: string|null, locationInfo: string|null}>} mylocation 스크린샷 파일 ID 및 위치 정보 객체
 */
async function automateMyLocation(page, ipAddress, drive, folderId) {
    console.log(`mylocation.co.kr 자동화 시작: ${ipAddress}`);
    let locationInfo = null;
    try {
        await page.goto(TARGET_URL_MYLOCATION, { waitUntil: 'networkidle0', timeout: 60000 });
        
        console.log(`mylocation.co.kr 페이지 로드 후 5초 대기...`);
        await new Promise(resolve => setTimeout(resolve, 5000));

        await page.waitForSelector(IP_INPUT_SELECTOR_MYLOCATION, { timeout: 10000 });
        await page.type(IP_INPUT_SELECTOR_MYLOCATION, ipAddress);

        console.log(`IP 주소 입력 후 5초 대기...`);
        await new Promise(resolve => setTimeout(resolve, 5000));

        await page.click(SUBMIT_BUTTON_SELECTOR_MYLOCATION);

        console.log(`주소 검색 버튼 클릭 후 5초 대기...`);
        await new Promise(resolve => setTimeout(resolve, 5000));

        // 결과 로딩 대기. 주소 정보가 나타날 때까지 기다립니다.
        // 대기 시간을 60초로 늘립니다.
        await page.waitForSelector(MYLOCATION_ADDRESS_SELECTOR, { timeout: 60000 }); // 대기 시간 60초로 변경
        console.log(`mylocation.co.kr 결과 페이지 로드 완료.`);

        locationInfo = await page.$eval(MYLOCATION_ADDRESS_SELECTOR, el => el.textContent.trim());
        console.log(`mylocation.co.kr 위치 정보: ${locationInfo}`);

        const screenshotFileId = await takeScreenshotAndUpload(page, ipAddress, drive, folderId, 'mylocation');
        return { screenshotFileId, locationInfo };
    } catch (error) {
        console.error(`ERROR: mylocation.co.kr 자동화 중 오류 발생 (${ipAddress}):`, error);
        throw error; // 오류를 다시 던져서 main 함수에서 처리할 수 있도록 합니다.
    }
}


/**
 * Google Sheets에서 경찰서 정보를 가져옵니다.
 * @param {object} sheets Google Sheets API 클라이언트
 * @param {string} spreadsheetId 스프레드시트 ID
 * @param {string} dbSheetName DB 시트 이름
 * @param {string} dbRange DB 시트 범위
 * @returns {Promise<Array<Array<string>>>} 경찰서 정보 배열
 */
async function getPoliceStation(sheets, spreadsheetId, dbSheetName, dbRange) {
    console.log(`경찰서 정보 조회 시작: 시트 '${dbSheetName}', 범위 '${dbRange}'`);
    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: spreadsheetId,
            range: `${dbSheetName}!${dbRange}`,
        });
        const rows = response.data.values;
        if (!rows || rows.length === 0) {
            console.log('경찰서 정보를 찾을 수 없습니다.');
            return [];
        }
        console.log(`경찰서 정보 ${rows.length}개 로드 완료.`);
        return rows;
    } catch (error) {
        console.error('ERROR: 경찰서 정보를 가져오는 중 오류 발생:', error);
        throw error;
    }
}


module.exports = {
    automateKisaWhois,
    automateMyLocation,
    getPoliceStation,
};
