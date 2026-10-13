# EZAccount PR Incremental Excel Automation

## 개요
SAP UI5 기반 ezaccount PR 데이터를 OData/REST로 수집해 누적 Excel 파일로 관리하는 자동화입니다. 실행할 때마다 마지막 실행 이후 추가된 PR 아이템만 증분 반영합니다.

## 프로젝트 구조
- `config/`
  - `config.example.json`: OData/REST 엔드포인트, 필터 필드 등 설정 예시
  - `mapping.example.yaml`: 소스 필드와 Excel 컬럼 매핑 규칙
- `src/`
  - `main.py`: 전체 실행 흐름
  - `client.py`: OData/REST 호출과 Playwright 로그인 지원
  - `sync.py`: `state/state.json` 관리 및 중복 키 처리
  - `transform.py`: PR 아이템 데이터 정규화, 매핑, USE 규칙 처리
  - `excel_template.py`: 템플릿 생성
  - `excel_writer.py`: 누적 엑셀 기록
- `templates/`
- `output/`
- `state/`
- `logs/`

## DevTools Network에서 확인할 항목
- 검색 키워드: `/sap/opu/odata`, `$format=json`, `sap/` 등
- PR 목록 요청은 `list_endpoint`에, PR 아이템 상세는 `detail_endpoint`에 설정
- `date_filter_field`는 리스트 요청에서 날짜 필터에 사용할 필드 이름
- 필요한 경우 `list_params`에 추가 쿼리 인자를 넣어주세요

## config 복사 가이드
1. `config/config.example.json`을 `config/config.json`으로 복사
2. `base_url`, `list_endpoint`, `detail_endpoint`, `date_filter_field` 등을 실제 OData 요청 URL에 맞춰 수정
3. `use_playwright_login`을 `true`로 설정하면 환경 변수 `EZACCOUNT_USERNAME`/`EZACCOUNT_PASSWORD`로 로그인 후 쿠키를 사용합니다

## 매핑 파일 작성
- `config/mapping.example.yaml`을 `config/mapping.yaml`로 복사
- `rules`는 `USE` 컬럼을 자동으로 채우기 위한 규칙입니다
- `source_to_excel`에 `source field name: excel column name` 매핑을 추가하세요

## 환경 변수 설정
1. `.env.example`을 복사해 `.env`를 만들고 실제 값으로 채웁니다.
2. 아래 값이 필요합니다.

```env
EZACCOUNT_USERNAME=your_user_id
EZACCOUNT_PASSWORD=your_password
EZACCOUNT_BASE_URL=https://ezaccount.wamc.co.kr
MASTER_REFRESH_URL=
```

> `MASTER_REFRESH_URL`은 회사 마스터 JSON 파일 URL을 넣는 옵션입니다. 비워두면 로컬의 `config/company_masters.json`을 사용합니다.

## EZAccount 자동입력 흐름
- `/api/pr/<request_id>/ezaccount` 엔드포인트가 Playwright 브라우저를 열어 로그인 후 PR 페이지에 값을 입력합니다.
- 실제 SAP UI element selector가 달라질 수 있어, 브라우저 콘솔/로그에서 selector 힌트를 확인해 맞춰야 합니다.
- 필수 환경 변수는 `EZACCOUNT_USERNAME`, `EZACCOUNT_PASSWORD`, `EZACCOUNT_BASE_URL`입니다.

## 실행 방법 (Windows PowerShell)
```powershell
cd C:\ezaccount
python app.py
```

## 공유 승인 엑셀 최신화
- `config/config.json`의 `shared_output_dir`에 OneDrive/SharePoint 동기화 폴더 경로를 설정합니다.
- `shared_output_filename`은 공유 링크가 계속 유지되도록 고정 파일명을 사용합니다. 기본값은 `MT_PRS_PR_LIST_SHARED.xlsx`입니다.
- 승인 완료 엑셀을 갱신하고 공유본을 다시 게시하려면 다음 명령을 실행합니다.

```powershell
cd C:\ezaccount
python -m src.main
```

- 동기화가 끝나면 `output_dir`에서 가장 최근 승인 완료 워크북을 임시 파일로 완성한 뒤, 공유 폴더의 고정 파일에 내용을 덮어씁니다. 기존 OneDrive 파일 항목과 공유 링크를 유지하기 위한 방식입니다. 해당 파일을 Excel에서 열어 잠근 경우 게시가 실패할 수 있으므로 닫고 다시 실행하세요.

## 결과
- 웹 화면에서 PR draft 작성
- 마스터 자동 새로고침
- EZAccount 자동입력 시도 및 로그 기록
- `state/`와 `logs/`에 실행 상태 저장
