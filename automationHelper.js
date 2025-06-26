// automationHelper.js

const path = require('path');
const fs = require('fs');
const moment = require('moment');

// 설정 파일(config.js) 불러오기
const config = require('./config');

// 스크린샷 임시 저장 폴더 경로 설정
const TEMP_SCREENSHOT_DIR = path.join(__dirname, config.TEMP_SCREENSHOT_DIR_NAME);

/**
 * 주어진 페이지에서 스크린샷을 찍고 Google Drive에 업로드합니다.
 * @param {Page} page Puppeteer 페이지 객체
 * @param {string} ipAddress 현재 처리 중인 IP 주소
 * @param {object} drive Google Drive API 클라이언트
 * @param {string} folderId 이미지를 저장할 Google Drive 폴더 ID
 * @param {string} screenshotNameType 스크린샷 파일명 유형 ('mylocation' 또는 'mylocation_error')
 * @returns {Promise<string|null>} 업로드된 파일의 ID, 또는 실패 시 null
 */
async function takeScreenshotAndUpload(page, ipAddress, drive, folderId, screenshotNameType) {
    if (!fs.existsSync(TEMP_SCREENSHOT_DIR)) {
        fs.mkdirSync(TEMP_SCREENSHOT_DIR, { recursive: true });
        console.log(`임시 스크린샷 폴더 생성: ${TEMP_SCREENSHOT_DIR}`);
    }

    const timestamp = moment().format('YYYYMMDD_HHmmss');
    const filename = `${ipAddress.replace(/\./g, '_')}_${screenshotNameType}_${timestamp}.png`;
    const filepath = path.join(TEMP_SCREENSHOT_DIR, filename);

    try {
        await page.screenshot({ path: filepath, fullPage: true });
        console.log(`스크린샷 저장: ${filepath}`);

        const fileMetadata = {
            name: filename,
            parents: [folderId],
            mimeType: 'image/png',
        };
        const media = {
            mimeType: 'image/png',
            body: fs.createReadStream(filepath),
        };

        const response = await drive.files.create({
            resource: fileMetadata,
            media: media,
            fields: 'id,webViewLink',
        });
        console.log(`Google Drive에 스크린샷 업로드 완료. 파일 ID: ${response.data.id}`);

        // 업로드 후 임시 파일 삭제
        fs.unlinkSync(filepath);
        console.log(`임시 스크린샷 파일 삭제: ${filepath}`);

        return response.data.id;
    } catch (uploadError) {
        console.error(`ERROR: 스크린샷 업로드 또는 임시 파일 삭제 중 오류 발생:`, uploadError);
        // 오류가 발생해도 임시 파일이 남아있을 수 있으므로 시도
        if (fs.existsSync(filepath)) {
            try {
                fs.unlinkSync(filepath);
                console.warn(`WARN: 오류 발생 후 임시 스크린샷 파일 정리: ${filepath}`);
            } catch (unlinkError) {
                console.warn(`WARN: 오류 발생 후 임시 파일 정리 실패: ${unlinkError.message}`);
            }
        }
        return null;
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
    let screenshotFileId = null;
    try {
        // 1. mylocation.co.kr 창 열기
        await page.goto(config.TARGET_URL_MYLOCATION, { waitUntil: 'networkidle0', timeout: 60000 });
        
        // 2. 초기 페이지 로딩 및 팝업 대기를 위해 충분히 대기 (1초 -> 5초 복원)
        console.log('mylocation.co.kr 사이트 초기화를 위해 5초 대기...');
        await new Promise(resolve => setTimeout(resolve, 5000));
        
        // 팝업창 처리 로직: '닫기' 버튼이 있는지 확인하고 클릭 시도
        try {
            await page.waitForSelector(config.POPUP_CLOSE_BUTTON_SELECTOR, { timeout: 5000, visible: true });
            console.log('팝업창 "닫기" 버튼 발견. 클릭하여 닫기 시도...');
            await page.click(config.POPUP_CLOSE_BUTTON_SELECTOR);
            console.log('팝업창 닫기 버튼 클릭 완료.');
            await new Promise(resolve => setTimeout(resolve, 1000)); // 팝업 닫힌 후 안정화 대기
        } catch (e) {
            console.log('팝업창 "닫기" 버튼을 찾을 수 없거나 팝업이 나타나지 않음. (정상 케이스)');
        }

        // IP 입력 필드 확인 및 입력 (재도입)
        await page.waitForSelector(config.IP_INPUT_SELECTOR_MYLOCATION, { visible: true, timeout: 10000 });
        console.log(`mylocation.co.kr IP 입력 필드 확인 완료.`);
        
        // 입력 필드 초기화 (기존 값 삭제)
        await page.evaluate(selector => {
            const input = document.querySelector(selector);
            if (input) {
                input.value = '';
            }
        }, config.IP_INPUT_SELECTOR_MYLOCATION);
        
        // 3. IP 넣기
        await page.type(config.IP_INPUT_SELECTOR_MYLOCATION, ipAddress);
        console.log(`IP 주소 입력 완료.`);

        // 주소 검색 버튼 확인 (재도입)
        await page.waitForSelector(config.SUBMIT_BUTTON_SELECTOR_MYLOCATION, { visible: true, timeout: 10000 });
        console.log(`mylocation.co.kr 제출 버튼 확인 완료.`);

        // 버튼 클릭 전에 1초 딜레이
        console.log(`버튼 클릭 전 1초 대기...`);
        await new Promise(resolve => setTimeout(resolve, 1000)); // 1000ms = 1초

        // 4. 주소검색 누르기 (클릭) - 페이지 탐색을 기다립니다.
        await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 60000 }), // 페이지 탐색을 기다립니다.
            page.evaluate(selector => { // JavaScript 클릭 사용 (더 안정적)
                document.querySelector(selector).click();
            }, config.SUBMIT_BUTTON_SELECTOR_MYLOCATION)
        ]);
        console.log(`mylocation.co.kr 주소 검색 버튼 클릭 완료 및 탐색 대기 완료.`);

        // 5. 추가 대기 (2초) (재도입)
        console.log(`mylocation.co.kr 검색 결과 요소 로드 대기를 위해 추가 2초 대기...`);
        await new Promise(resolve => setTimeout(resolve, 2000));

        // 주소 정보 요소 확인 및 결과값 받기 (재도입)
        await page.waitForSelector(config.MYLOCATION_ADDRESS_SELECTOR, { timeout: 60000 });
        console.log(`mylocation.co.kr 주소 정보 요소 확인 완료.`);
        
        // 추가 대기 (최종 결과값 받기 전)
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // 6. 결과값 받아서
        locationInfo = await page.$eval(config.MYLOCATION_ADDRESS_SELECTOR, el => el.textContent.trim());
        console.log(`mylocation.co.kr 위치 정보: ${locationInfo}`);

        // 스크린샷 찍고 업로드 (Google Drive 폴더 ID 사용)
        screenshotFileId = await takeScreenshotAndUpload(page, ipAddress, drive, folderId, 'mylocation');
        return { screenshotFileId, locationInfo };
    } catch (error) {
        console.error(`ERROR: mylocation.co.kr 자동화 중 오류 발생 (${ipAddress}):`, error);
        screenshotFileId = await takeScreenshotAndUpload(page, ipAddress, drive, folderId, 'mylocation_error').catch(e => {
            console.warn(`WARN: 오류 발생 시 mylocation 스크린샷 찍기 실패: ${e.message}`);
            return null;
        });
        return { screenshotFileId, locationInfo: `오류 발생: ${error.message}` };
    }
}

module.exports = {
    automateMyLocation,
};
