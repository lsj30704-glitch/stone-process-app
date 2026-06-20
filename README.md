# 석공사 일일공정관리 (stone-process-app)

석공사_일일공정관리_lsj.xlsx 파일과 OneDrive로 동기화되는 모바일 친화형 공정관리 앱입니다.
site-diary와는 완전히 별도의 프로젝트입니다.

## 기능
- 일일실적입력(외부/내부)만 모바일에서 입력 가능, 나머지(기준정보/달성률현황/만회계획)는 엑셀 수식을 그대로 옮긴 자동계산 결과를 �으로만 봄
- OneDrive(개인 계정) 연동: 로그인 후 저장하면 석공사_일일공정관리_lsj.xlsx 파일에 새 행이 자동 추가됨

## 실행 방법
1. `npm install`
2. `.env.example`을 `.env`로 복사하고 `VITE_MSAL_CLIENT_ID` 값을 채움 (Azure 앱 등록 가이드 별도 안내)
3. `npm run dev`

## 배포
GitHub 리포지토리 생성 후 Vercel에 연결하면 site-diary와 동일한 방식으로 배포됩니다.
