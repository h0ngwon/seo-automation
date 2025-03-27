const fs = require("fs");
const xlsx = require("xlsx");
const { google } = require("googleapis");

// ✅ 1. 엑셀에서 URL 읽기
function getUrlListFromExcel(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const range = xlsx.utils.decode_range(sheet["!ref"]);
  const urlList = [];

  for (let row = 1; row <= range.e.r; row++) {
    const cell = sheet[xlsx.utils.encode_cell({ c: 0, r: row })];
    if (cell && cell.v) urlList.push(cell.v.trim());
  }
  return urlList;
}

// ✅ 2. 브라우저 자동 열기 (ESM 방식)
const openUrl = async (url) => {
  const open = (await import("open")).default;
  await open(url);
};

// ✅ 3. OAuth2 인증 처리
async function getOAuth2Client() {
  const credentials = require("./oauth_client.json");
  const { client_secret, client_id, redirect_uris } = credentials.web;

  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
  );

  const tokenPath = "./token.json";

  if (fs.existsSync(tokenPath)) {
    const token = JSON.parse(fs.readFileSync(tokenPath));
    oAuth2Client.setCredentials(token);
    return oAuth2Client;
  }

  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    scope: ["https://www.googleapis.com/auth/indexing"],
  });

  console.log("👉 브라우저에서 로그인 후 인증 코드를 입력하세요:");
  console.log(authUrl);
  await openUrl(authUrl);

  const readline = require("readline").createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  const code = await new Promise((resolve) => {
    readline.question("코드 입력: ", (code) => {
      readline.close();
      resolve(code);
    });
  });

  const { tokens } = await oAuth2Client.getToken(code);
  oAuth2Client.setCredentials(tokens);
  fs.writeFileSync(tokenPath, JSON.stringify(tokens));
  console.log("✅ 토큰이 저장되었습니다.");

  return oAuth2Client;
}

// ✅ 4. 색인 요청 & 캐시 저장 + 연속 실패 제한 + 실패 원인 출력 + 실행 통계
async function requestIndexing(urlList) {
  const auth = await getOAuth2Client();
  const indexing = google.indexing({ version: "v3", auth });

  const cachePath = "./indexed_cache.json";
  let indexedCache = [];

  if (fs.existsSync(cachePath)) {
    indexedCache = JSON.parse(fs.readFileSync(cachePath, "utf8"));
  }

  const failedUrls = [];
  const newIndexedUrls = [];
  let consecutiveFails = 0;
  let skipped = 0;

  for (let i = 0; i < urlList.length; i++) {
    const url = urlList[i];

    if (indexedCache.includes(url)) {
      console.log(`⏭️ 이미 색인된 URL (캐시): ${url}`);
      skipped++;
      continue;
    }

    try {
      await indexing.urlNotifications.publish({
        requestBody: {
          url: url,
          type: "URL_UPDATED",
        },
      });
      console.log(`✅ 색인 요청 성공 (${i + 1}/${urlList.length}): ${url}`);
      newIndexedUrls.push(url);
      consecutiveFails = 0;
    } catch (err) {
      const status = err.response?.status || "Unknown";
      const message = err.response?.data?.error?.message || "Unknown Error";

      console.log(`❌ 색인 요청 실패 (${i + 1}/${urlList.length}): ${url}`);
      console.log(`   ↳ HTTP ${status} - ${message}`);

      failedUrls.push({ url, status, message });
      consecutiveFails++;

      if (consecutiveFails >= 20) {
        console.error(`🚨 연속 실패 20회 초과! 실행을 중단합니다.`);
        break;
      }
    }

    await new Promise((r) => setTimeout(r, 100));
  }

  const updatedCache = [...new Set([...indexedCache, ...newIndexedUrls])];
  fs.writeFileSync(cachePath, JSON.stringify(updatedCache, null, 2));

  // 실패 목록 출력
  console.log("\n🎯 [실패한 URL 목록]");
  if (failedUrls.length > 0) {
    failedUrls.forEach(({ url, status, message }) => {
      console.log(`❌ ${url} → HTTP ${status} - ${message}`);
    });
  } else {
    console.log("✅ 모두 성공!");
  }

  // ✅ 실행 통계 요약
  console.log("\n📊 [실행 요약]");
  console.log(`- 총 URL 수       : ${urlList.length}`);
  console.log(`- 색인 성공       : ${newIndexedUrls.length}`);
  console.log(`- 색인 실패       : ${failedUrls.length}`);
  console.log(`- 캐시로 건너뜀   : ${skipped}`);
}

// ✅ 5. 실행
(async () => {
  const urlList = getUrlListFromExcel("./urls.xlsx");
  console.log(`📢 총 ${urlList.length}개의 URL을 색인 확인 및 요청합니다.`);
  await requestIndexing(urlList);
})();
