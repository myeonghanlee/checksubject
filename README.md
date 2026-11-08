# 고등학교 학점 배당표 검증기

교육과정 기준 과목과 교육감 승인 과목을 대조해 학점 배당표의 **D열 과목명**, **M열 비고**를 확인하는 Streamlit 앱입니다.

## 기능

- 여러 `.xlsx`/`.xlsm` 파일을 한 번에 업로드하고, 파일마다 검토할 시트와 헤더 행을 선택할 수 있습니다.
- 표 1~3의 과목과 교육감 승인 완료 과목을 독립적으로 조회합니다.
- 표 2는 `특목고 선택 과목`, 표 3은 `전문 교과` 표기가 M열에 있는지 확인합니다.
- 서울특별시교육청 승인 과목은 `교육감 승인 과목`, 다른 시·도교육청 승인 과목은 `타 시·도 교육감 승인 과목`으로 확인합니다. 전문으로 판정되는 승인 과목은 `(전문)`을 붙입니다.
- 표 1에 등재된 교과군명도 인식합니다. 예를 들어 `기술⋅가정/정보`의 교과군 `기술·가정`, `정보`, 그리고 `한문`을 교과군 행으로 판정합니다. 기술·가정 교과군명에 한해 `⋅`, `∙`, `·`, `ㆍ` 기호 차이도 같은 교과군 표기로 봅니다.
- 통합 결과에 파일명·시트명·엑셀 행을 담아 CSV로 내려받을 수 있습니다.

## 과목명 일치 규칙

과목명 비교 전에 끝의 빈칸과 말미의 `*`, `**`를 무시합니다. 앞쪽 공백, 이름 중간의 띄어쓰기, 기호, 숫자, 로마자 차이는 자동 보정하지 않습니다. 단, 교과군명은 교과군 별도 규칙에 따라 대조합니다. 유사 후보는 참고 정보일 뿐 자동 일치로 판정하지 않습니다.

## 로컬 실행

Python 3.10 이상을 권장합니다.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python scripts/build_reference_catalog.py
streamlit run app.py
```

## GitHub에서 `reference_catalog.json` 생성하기

프로젝트에는 `scripts/build_reference_catalog.py` 스크립트와 `.github/workflows/build-reference-catalog.yml` GitHub Actions 워크플로가 포함되어 있습니다. 원본 기준 문서로부터 `data/reference_catalog.json`을 만들고, 결과가 바뀌면 같은 브랜치에 자동 커밋합니다.

1. 프로젝트의 `data/고등학교교과및과목검증기준.md`, `scripts/build_reference_catalog.py`, `.github/workflows/build-reference-catalog.yml` 파일을 GitHub 저장소에 올립니다.
2. 저장소의 **Actions** 탭에서 **Build reference catalog** 워크플로를 엽니다.
3. **Run workflow**를 눌러 실행할 브랜치를 선택하고 다시 **Run workflow**를 누릅니다.
4. 실행이 성공하면 `data/reference_catalog.json` 변경 커밋이 해당 브랜치에 생성됐는지 확인합니다.

이 워크플로는 기준 Markdown 또는 생성 스크립트가 저장소에 push될 때도 자동 실행됩니다. 저장소의 Actions 권한 정책이 쓰기를 제한해 커밋이 실패하면 **Settings → Actions → General → Workflow permissions**에서 쓰기 권한을 허용한 뒤 다시 실행하세요. 워크플로 자체에도 `contents: write` 권한이 선언되어 있습니다.

로컬에서 생성하려면 저장소 루트에서 `python scripts/build_reference_catalog.py`를 실행하면 됩니다.

## Streamlit Community Cloud 배포

1. 프로젝트 전체와 생성된 `data/reference_catalog.json`을 GitHub 저장소에 둡니다.
2. Streamlit Community Cloud에서 해당 저장소와 `app.py`를 선택합니다.
3. `requirements.txt`를 기준으로 의존성을 설치해 배포합니다.
4. 기준 Markdown, 승인 과목 CSV, 생성된 JSON이 모두 `data/` 폴더에 있는지 확인합니다.

기준 목록을 갱신할 때는 원본 기준 파일을 교체해 GitHub에 push하고 Actions 성공 및 JSON 커밋을 확인하세요.

## 자료 및 개인정보

기준 자료의 출처·유의사항은 `data/고등학교교과및과목검증기준.md`에 포함되어 있습니다. 앱은 업로드 파일을 읽기 전용으로 처리하며 앱 코드에서 업로드 파일을 디스크에 저장하거나 외부로 전송하지 않습니다. 공개 배포 시에는 Streamlit 호스팅 환경의 접근·보관 정책을 별도로 확인하세요.
