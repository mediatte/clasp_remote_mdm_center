# 고운고등학교 원격 요청

## 원격 요청 센터

고운고등학교 학생용 원격 요청 웹앱입니다.

### 비공개 설정

민감한 값은 `.env`에 두고, 로컬 생성 파일로만 반영합니다.

1. `.env.example`을 복사해서 `.env`를 만듭니다.
2. `.env`에 Apps Script ID, Spreadsheet ID, 발신 메일 주소를 입력합니다.
3. 아래 명령으로 설정 파일을 생성합니다.

```bash
node scripts/sync-env.mjs
```

생성되는 파일:

- `Config.generated.js`
- `.clasp.json`

이 두 파일과 `.env`는 `.gitignore`에 포함되어 있습니다.
# clasp_remote_mdm_center
# clasp_remote_mdm_center
