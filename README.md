# SharkNinja Daily Report 대시보드

Streamlit 앱. 로컬 폴더 또는 **구글 드라이브**(서비스 계정)에서 `Madup_Sharkninja_Daily Report_YYMMDD.xlsx` 를 읽습니다.

## Streamlit Community Cloud 배포

1. 이 저장소를 GitHub에 푸시합니다.
2. [share.streamlit.io](https://share.streamlit.io) 에서 GitHub로 로그인 → **New app** → 저장소 / 브랜치 / Main file: `app.py` → **Deploy**.
3. 앱 설정 **Secrets** 에 다음을 넣습니다 (이름 그대로):

   ```toml
   GOOGLE_DRIVE_FOLDER_ID = "여기에_드라이브_폴더_ID"
   GOOGLE_SERVICE_ACCOUNT_JSON = """{ ... 서비스 계정 JSON 전체 ... }"""
   ```

4. **구글 클라우드 콘솔**에서 서비스 계정을 만들고 JSON 키를 내려받습니다. **Drive API** 를 사용 설정합니다.
5. 드라이브에서 리포트가 있는 **폴더**를 해당 서비스 계정 이메일( `xxx@....iam.gserviceaccount.com` )과 **공유**(뷰어 이상).

폴더 ID는 URL `https://drive.google.com/drive/folders/폴더ID` 의 마지막 부분입니다.

Secrets 가 없으면 앱은 로컬과 같이 사이드바의 **리포트 폴더** 경로를 사용합니다 (클라우드에서는 사용 불가).

## 로컬 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

로컬에서 드라이브를 쓰려면 프로젝트에 `.streamlit/secrets.toml` 을 만들고 위와 동일한 키를 넣습니다.
