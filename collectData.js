// collectData.js

const puppeteer = require('puppeteer');
const moment = require('moment'); // 시간 관리를 위해 moment.js 사용
const fs = require('fs');
const path = require('path');

// 설정 변수
const TORRENT_WEB_UI_URL = 'https://utweb.rainberrytv.com/gui/index.html?v=1.5.0.6261&localauth=localapi5879981bd9d1bc66:#/dashboard';
const PUPPETEER_HEADLESS = true; // true: 브라우저 UI 숨김, false: 브라우저 UI 표시 (디버깅용)
const PUPPETEER_ARGS = [
    '--no-sandbox',
    '--disable-setuid-sandbox',
    '--disable-infobars',
    '--window-size=1280,800',
    '--ignore-certificate-errors',
    '--start-maximized', // 브라우저 창을 최대화하여 더 넓은 영역 스크린샷
];
const SCREENSHOT_DIR = path.join(__dirname, 'torrent_screenshots'); // 스크린샷 저장 폴더
const INTERVAL_SECONDS = 10; // 스크린샷을 찍을 간격 (초) - 30초에서 10초로 변경됨

/**
 * 주어진 URL의 페이지 스크린샷을 찍어 저장합니다.
 * @param {string} url 스크린샷을 찍을 URL
 * @returns {Promise<string|null>} 저장된 스크린샷 파일 경로, 또는 실패 시 null
 */
async function takeTorrentUIScreenshot(url) {
    let browser;
    let page;
    let screenshotPath = null;

    console.log(`URL 스크린샷 촬영 시작: ${url}`);
    try {
        browser = await puppeteer.launch({
            headless: PUPPETEER_HEADLESS,
            args: PUPPETEER_ARGS,
        });
        page = await browser.newPage();
        await page.setViewport({ width: 1280, height: 800 }); // 뷰포트 설정

        await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); // networkidle2 대신 domcontentloaded 사용 (초기 HTML 로드)

        // **핵심 변경 부분:**
        // 1. 특정 요소가 로드될 때까지 기다립니다. (가장 좋은 방법)
        //    토렌트 대시보드에서 주요 콘텐츠를 표시하는 특정 CSS 셀렉터가 있다면 그것을 사용합니다.
        //    예: 토렌트 목록을 담고 있는 컨테이너의 클래스나 ID
        const contentSelector = '.torrent-list-container, #torrent_table, .dashboard-content'; // 페이지 내용을 나타내는 적절한 셀렉터로 변경하세요!
        try {
            await page.waitForSelector(contentSelector, { timeout: 15000 }); // 15초 동안 해당 셀렉터가 나타날 때까지 대기
            console.log(`페이지의 주요 콘텐츠 (${contentSelector}) 로드 확인.`);
        } catch (selectorError) {
            console.warn(`WARN: 주요 콘텐츠 셀렉터(${contentSelector})를 ${15000 / 1000}초 내에 찾지 못했습니다. 페이지 로딩이 느리거나 셀렉터가 잘못되었을 수 있습니다. ${selectorError.message}`);
            // 셀렉터를 찾지 못하더라도 계속 진행하여 일단 스크린샷을 찍도록 합니다.
        }

        // 2. 또는, 모든 네트워크 요청이 완료될 때까지 기다립니다. (networkidle0)
        //    단, 백그라운드에서 계속 요청이 있는 경우 무한 대기할 수 있습니다.
        // await page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 30000 });

        // 3. 또는, 일정 시간 동안 기다립니다. (간단하지만 불안정)
        //    페이지 로딩 시간에 따라 달라질 수 있으므로 최후의 수단으로 사용하세요.
        await page.waitForTimeout(5000); // 5초 동안 추가 대기

        // 스크린샷 저장 폴더 생성
        if (!fs.existsSync(SCREENSHOT_DIR)) {
            fs.mkdirSync(SCREENSHOT_DIR, { recursive: true });
        }

        const timestamp = moment().format('YYYYMMDD_HHmmss');
        const filename = `torrent_ui_dashboard_${timestamp}.png`;
        screenshotPath = path.join(SCREENSHOT_DIR, filename);

        await page.screenshot({ path: screenshotPath, fullPage: true });
        console.log(`스크린샷 저장 완료: ${screenshotPath}`);

    } catch (error) {
        console.error(`ERROR: 스크린샷 촬영 실패: ${error.message}`);
        screenshotPath = null;
    } finally {
        if (page && !page.isClosed()) {
            await page.close();
        }
        if (browser && browser.isConnected()) {
            await browser.close();
            console.log("Puppeteer 브라우저 종료.");
        }
    }
    return screenshotPath;
}

// 스크린샷 촬영 작업을 반복 실행하는 함수
async function startScreenshotAutomation() {
    console.log(`\n스크린샷 자동화를 시작합니다. ${INTERVAL_SECONDS}초마다 실행됩니다.`);
    // 즉시 한 번 실행하고, 그 다음부터는 INTERVAL_SECONDS 간격으로 실행
    await takeTorrentUIScreenshot(TORRENT_WEB_UI_URL);

    setInterval(async () => {
        await takeTorrentUIScreenshot(TORRENT_WEB_UI_URL);
    }, INTERVAL_SECONDS * 1000); // 밀리초 단위로 변환
}

// 자동화 시작
startScreenshotAutomation().catch(err => {
    console.error("자동화 시작 중 예외 발생:", err);
    process.exit(1);
});
