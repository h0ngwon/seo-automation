# 🔍 SEO Automation (Google Indexing API)

이 프로젝트는 Google Indexing API를 통해 엑셀 파일(urls.xlsx)에 저장된 URL 목록을 자동으로 색인 요청해주는 Node.js 스크립트입니다.

# 🚀 주요 기능

✅ 엑셀 파일에서 URL을 자동 추출하여 Google에 색인 요청 <br/>
✅ OAuth 2.0 사용자 인증 방식 지원 <br/>
✅ 수천 개의 URL도 빠르게 처리 가능 <br/>
✅ 색인 실패한 URL 목록 자동 출력

# 📂 파일 구조
```
📦 seo-automation
 ┣ 📄 seo-automation.js
 ┣ 📄 .gitignore
 ┣ 📄 oauth_client.json (Google Cloud에서 직접 다운로드 필요)
 ┗ 📄 urls.xlsx (URL 목록을 직접 생성)
 ```

# 🛠️ 사용 방법

## 1. 사전 준비
Google Cloud Console의 라이브러리에서 Web Search Indexing API가 설치되어있어야 합니다.

Google Cloud Console에서 OAuth 2.0 클라이언트 ID를 생성하고, oauth_client.json을 다운로드합니다.

Google Search Console에서 OAuth 로그인에 사용할 계정 이메일을 해당 웹사이트의 전체 권한 사용자로 추가합니다.

색인할 URL 목록을 포함한 엑셀 파일(urls.xlsx)을 프로젝트 루트에 배치합니다. (A열 2행부터 URL 작성)

## 2. 설치 및 실행

### 저장소 클론
```bash
git clone <repository_url>
cd seo-automation
```
### 의존성 설치
```bash
npm install
```
### 스크립트 실행
```bash
node seo-automation.js
```

브라우저가 열리면 OAuth 로그인 후, 주소창에 뜨는 code 값을 터미널에 붙여넣으면 자동으로 색인이 시작됩니다.

# 🚨 중요사항
민감한 인증 정보(oauth_client.json, token.json)는 저장소에 절대 업로드하지 마세요.
.gitignore 설정으로 민감한 파일 및 불필요한 파일은 제외됩니다.

# 📌 의존성
```bash
npm install xlsx googleapis open
```
# 📄 참고 문서

<a href="https://developers.google.com/search/apis/indexing-api/v3/quickstart">Google Indexing API Quickstart</a>
<br/>
<a href="https://support.google.com/webmasters/answer/2453966?hl=ko">Google Search Console 사용자 추가 방법</a>