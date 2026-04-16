# SharkNinja Daily Report 대시보드

Streamlit 앱. **Madup API**(Dropbox HTTP 다운로드), **구글 드라이브**(서비스 계정), 또는 로컬 폴더에서 `Madup_Sharkninja_Daily Report_YYMMDD.xlsx` 를 읽습니다.  
우선순위: **Madup API → 구글 드라이브 → 로컬 폴더** (Secrets에 Madup 키가 있으면 Madup을 사용).

## 내 PC에서만 돌리고 팀에게 보여주기 (가장 단순)

1. 이 폴더에서 `pip install -r requirements.txt` (최초 1회)
2. **`run_dashboard.bat`** 더블클릭 (또는 터미널에서 `streamlit run app.py`)
3. 리포트는 **항상 사이드바에 적어 둔 로컬 폴더**에 두면 됨 (Dropbox 동기 경로 그대로 가능)
4. **같은 Wi-Fi / 사내망**에 있는 팀원에게 브라우저 주소 알려주기:
   - `http://192.168.x.x:8501` 처럼 **대시보드 돌리는 PC의 IPv4**
   - IP 확인: Windows에서 `ipconfig` → 무선/LAN 어댑터의 **IPv4 주소**

**주의:** 대시보드를 켠 **PC가 켜져 있고 Streamlit이 실행 중**이어야 다른 사람도 볼 수 있음.  
처음 한 번 **Windows 방화벽**에서 Python/포트 8501 허용 창이 뜨면 **허용** 선택.

외부(집 등)에서까지 보려면 Tailscale·VPN 등은 별도로 필요함.

---

## Streamlit Community Cloud 배포

1. 이 저장소를 GitHub에 푸시합니다.
2. [share.streamlit.io](https://share.streamlit.io) 에서 GitHub로 로그인 → **New app** → 저장소 / 브랜치 / Main file: `app.py` → **Deploy**.
3. 앱 설정 **Secrets** 에 다음 중 필요한 것을 넣습니다 (이름 그대로).

   **(선택 A) Madup API — Dropbox에서 파일 받기**

   ```toml
   MADUP_API_KEY = "발급받은_API_키"
   # 선택: 기본값 https://api-auth.madup-dct.site
   # MADUP_API_BASE = "https://api-auth.madup-dct.site"

   # 아래 둘 중 하나: 폴더 경로(끝에 슬래시 없이) 또는 파일 전체 경로
   MADUP_DROPBOX_FOLDER = "/reports/샤크닌자"
   # MADUP_DROPBOX_PATH = "/reports/.../Madup_Sharkninja_Daily Report_260401.xlsx"
   # (호환) MADUP_DROPBOX_FILE_PATH 로 동일 의미
   ```

   - `MADUP_DROPBOX_FOLDER` 가 있으면 사이드바에서 **최신 자동** 또는 **날짜 선택**으로 `Madup_Sharkninja_Daily Report_YYMMDD.xlsx` 를 받습니다.
   - 폴더 대신 **단일 파일**만 쓰려면 `MADUP_DROPBOX_PATH`(또는 `MADUP_DROPBOX_FILE_PATH`)에 Dropbox 상의 전체 경로를 넣습니다.

   **(선택 B) 구글 드라이브**

   ```toml
   GOOGLE_DRIVE_FOLDER_ID = "여기에_드라이브_폴더_ID"
   GOOGLE_SERVICE_ACCOUNT_JSON = """{ ... 서비스 계정 JSON 전체 ... }"""
   ```

4. Madup을 쓰지 않고 드라이브만 쓸 때: **구글 클라우드 콘솔**에서 서비스 계정을 만들고 JSON 키를 내려받습니다. **Drive API** 를 사용 설정합니다.
5. 드라이브에서 리포트가 있는 **폴더**를 해당 서비스 계정 이메일( `xxx@....iam.gserviceaccount.com` )과 **공유**(뷰어 이상).

폴더 ID는 URL `https://drive.google.com/drive/folders/폴더ID` 의 마지막 부분입니다.

Madup·드라이브 Secrets 가 모두 없으면 앱은 로컬과 같이 사이드바의 **리포트 폴더** 경로를 사용합니다 (클라우드에서는 사용 불가).

## 로컬 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

로컬에서 드라이브를 쓰려면 프로젝트에 `.streamlit/secrets.toml` 을 만들고 위와 동일한 키를 넣습니다.
