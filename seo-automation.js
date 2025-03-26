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
    if (cell && cell.v) urlList.push(cell.v);
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
  const credentials = require("./oauth_client.json"); // oauth_client.json 파일 경로
  const { client_secret, client_id, redirect_uris } = credentials.web;

  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0] // 예: http://localhost:3000
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

  console.log("👉 브라우저에서 로그인하여 인증 코드를 입력하세요 (url의 code 부분부터 &scope전까지 복붙!!!):");
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

// ✅ 4. Google Indexing API에 URL 제출
async function requestIndexing(urlList) {
  const auth = await getOAuth2Client();
  const indexing = google.indexing({ version: "v3", auth });

  let failedUrls = [];

  for (let i = 0; i < urlList.length; i++) {
    const url = urlList[i];
    try {
      await indexing.urlNotifications.publish({
        requestBody: {
          url: url,
          type: "URL_UPDATED",
        },
      });
      console.log(`✅ 성공 (${i + 1}/${urlList.length}): ${url}`);
    } catch (err) {
      console.log(`❌ 실패 (${i + 1}/${urlList.length}): ${url}`);
      failedUrls.push(url);
    }
    await new Promise((r) => setTimeout(r, 100));
  }

  console.log("\n🎯 [실패한 URL 목록]");
  failedUrls.length > 0
    ? failedUrls.forEach((url) => console.log(url))
    : console.log("✅ 모두 성공!");
}

// ✅ 5. 실행
(async () => {
  const urlList = getUrlListFromExcel("./urls.xlsx");
  console.log(`📢 총 ${urlList.length}개의 URL을 색인 요청합니다.`);
  await requestIndexing(urlList);
})();
