# M365 Outlook 견적서·회신 초안 자동화

Windows PC에서 M365 받은편지함을 주기적으로 확인하고 **첨부된 회사 견적서 양식**으로 견적서를 작성해 회신 초안을 임시보관함에 저장합니다. 자동 발송 기능과 `Mail.Send` 권한이 없습니다.

현재 저장소의 `outlook-quote-automation/`에 저장된 실행 가능한 Python 소스입니다. 실제 Outlook/API 계정 연결은 사용자의 최초 설정이 필요합니다. 이 개발 환경에서 실제 메일함 연결 및 실제 가격 검색을 수행한 것은 아닙니다.

## 동작

1. 받은편지함의 신규 메일 중 ‘견적’, ‘quotation’, ‘rfq’, ‘가격 문의’를 포함하는 메일을 감지합니다. 읽음 여부와 무관합니다. 다른 폴더로 자동 분류된 메일·공유 사서함은 현재 대상이 아닙니다.
2. 본문과 지원되는 첨부파일에서 모델·설명·수량을 추출합니다. 자유형은 선택 사항인 OpenAI API, 정형 입력은 규칙 추출을 사용합니다.
3. SerpApi의 Google Shopping 검색으로 국내 원화 판매가를 조회합니다. 같은 모델과 명시된 용량이 일치하는 결과 중 최저 표시가를 고릅니다. 최대 2페이지(설정 가능) 검색 범위 내 가격이지 인터넷 전체 최저가 보장이 아닙니다.
4. 지정 마진을 적용해 VAT 별도 단가를 계산합니다. 모델·수량·가격·마진·VAT 조건이 확인되지 않으면 ‘확인필요’로 남깁니다.
5. 원본 양식을 복사해 여러 품목을 추가하고, 공급가·부가세·총합계 수식을 유지합니다. Python Decimal 계산과 저장 파일 재열람으로 검산합니다. 미확인 품목이 있으면 전체 합계도 ‘확인필요’입니다.
6. 원본 메일의 회신 초안을 만들고 `Reply-To`(없으면 발신자)를 수신자로 지정합니다. 제목·본문·견적서 첨부를 채운 후 실제 초안의 수신자·제목·첨부 여부를 조회해 검증합니다. CC/BCC는 비웁니다.

## Windows 최초 설정

### 1. 설치

- Windows 10/11, **Python 3.11 이상**과 인터넷 연결이 필요합니다. Python 설치 시 Python Launcher(`py`)를 포함하세요.
- `setup.bat`을 실행합니다. 가상환경·패키지를 설치하고 `.env`, `config.json`을 생성해 메모장으로 엽니다.
- 테스트만 하려면 설치 후 `demo.bat`을 실행합니다. 가상 단가를 사용하는 데모 견적서와 메일을 생성하며, Outlook이나 검색 API에 접속하지 않습니다.

### 2. M365 앱 등록 (최초 1회)

Microsoft Entra 관리 센터 → **앱 등록 → 새 등록**:

1. 앱 이름 예: `Outlook Quote Draft Automation`, 계정 유형은 자신의 조직 디렉터리만 지원하는 단일 테넌트.
2. 개요에서 **애플리케이션(클라이언트) ID**와 **디렉터리(테넌트) ID**를 복사합니다.
3. 인증 → 고급 설정 → **공용 클라이언트 흐름 허용: 예**.
4. API 권한 → Microsoft Graph → **위임된 권한 `Mail.ReadWrite`** 추가. `Mail.Send`와 애플리케이션 권한은 추가하지 않습니다. 조직 정책상 필요한 경우 관리자 동의를 받습니다.
5. `.env`에 입력합니다. 비밀 클라이언트 키는 만들거나 입력할 필요가 없습니다.

```dotenv
M365_TENANT_ID=본인-조직-테넌트-ID
M365_CLIENT_ID=앱-클라이언트-ID
SERPAPI_API_KEY=본인-SerpApi-키
OPENAI_API_KEY=선택-OpenAI-키
```

6. `login.bat` 실행 → 화면에 표시되는 Microsoft 로그인 URL과 코드를 사용해 자신의 Outlook 계정으로 로그인합니다. 조직이 Device Code 로그인을 차단하면 IT 관리자의 정책 확인이 필요합니다.

Outlook 플러그인을 ChatGPT에 설치하는 것과 별개인 독립 실행 프로그램입니다. Outlook 데스크톱 앱을 켜둘 필요는 없습니다. 이 코드는 기본 Microsoft 상용 클라우드용입니다.

### 3. 가격 검색·마진 설정

[SerpApi](https://serpapi.com/) 계정의 API 키가 필요합니다. Google Shopping 결과를 가져오는 **제3자 유료/쿼터 기반 서비스**이며 Google 계정 비밀번호를 받지 않습니다. 검색에는 제품명·모델·제품 설명만 전달하고 문의자 이메일은 전달하지 않습니다.

`config.json` 예시(요청하신 매입가 10% 가산을 기본값으로 반영):

```json
{
  "mailbox_login": "your-account@your-company.com",
  "poll_seconds": 120,
  "initial_lookback_hours": 0,
  "subject_keywords": ["견적", "quotation", "rfq", "가격 문의"],
  "extractor": "auto",
  "openai_model": "gpt-4.1-mini",
  "search_pages": 2,
  "margin_mode": "markup",
  "margin_percent": 10,
  "search_price_vat": "included",
  "signature": "회사명 / 담당자 / 연락처"
}
```

- `margin_mode: "markup"`: VAT 제외 매입단가 × (1 + 가산율). 예: 검색가 11,000원(VAT 포함) → 매입 공급가 10,000원 → 10% 가산 공급단가 11,000원 → 판매 VAT 포함 12,100원.
- `margin_mode: "gross_margin"`: VAT 제외 매입단가 ÷ (1 − 판매가 기준 이익률). 20% 이익률이면 10,000 ÷ 0.8 = 공급단가 12,500원.
- **기본값은 `margin_mode: "markup"`, `margin_percent: 10`입니다.** 사용자가 지정한 매입가 10% 가산입니다. `null`로 변경하면 마진 미설정으로 처리해 단가는 확인필요로 남깁니다.
- `search_price_vat`: `included`(검색 가격 VAT 포함), `excluded`(VAT 별도), `unknown`(불명확). 기본 `unknown`이며 이 경우 검색 근거는 남기고 단가는 확인필요로 둡니다. 검색 API가 개별 판매처의 세금 조건을 확정하지 않으므로, `included/excluded`는 사용자가 적용 가능한 기준을 확인한 뒤 설정해야 합니다.
- 배송비는 판매처·수량·지역에 따라 달라 자동 합산하지 않습니다. **상품 표시가 기준**이며 견적서·본문에 배송비 등 최종 확인필요를 남깁니다.
- 쿠폰·회원·카드 조건부 가격, 중고·리퍼·할부·대여·품절, 모델 불일치, 요청에 없는 묶음/호환품/액세서리는 가능한 범위에서 제외합니다. 상세 페이지의 실재고·단위/옵션은 별도 확인해야 합니다.
- 0% 마진을 원하면 `0`을 명시합니다. 단가와 VAT는 Excel `ROUND`와 같이 원 단위 반올림합니다.

### 4. 실행

`start.bat` 실행 → 기본 120초 간격으로 감시합니다. 창을 닫으면 중지합니다.

처음 감시를 시작한 **이후** 들어오는 메일부터 처리합니다. 이전 메일도 처리하려면 첫 실행 **전** `initial_lookback_hours`를 `24` 등으로 변경하세요. 최초 커서가 생성된 뒤에는 이 옵션을 변경해도 과거로 되돌아가지 않습니다. PC가 꺼져 있던 동안의 메일은 다음 실행 때 저장된 커서 이후부터 처리합니다.

명령 프롬프트에서도 실행할 수 있습니다:

```bat
.venv\Scripts\python.exe -m quote_automation doctor
.venv\Scripts\python.exe -m quote_automation auth
.venv\Scripts\python.exe -m quote_automation once
.venv\Scripts\python.exe -m quote_automation watch
.venv\Scripts\python.exe -m quote_automation status
```

Windows 로그인 시 자동 실행하려면 PowerShell에서 `register-startup.ps1`을 실행합니다(조직의 스크립트 실행 정책 범위 내에서 사용). 작업 스케줄러의 `OutlookQuoteDraftAutomation` 작업으로 등록됩니다. 사용자 로그인 상태에서 실행되며 로그아웃·절전·종료 중에는 감시하지 않습니다. 자동 작업 등록 스크립트는 설치 시 자동 실행하지 않습니다.

삭제:

```powershell
Unregister-ScheduledTask -TaskName OutlookQuoteDraftAutomation -Confirm:$false
```

## 품목 추출

- `OPENAI_API_KEY`가 있으면 메일 본문과 지원 첨부자료의 텍스트를 OpenAI API에 보내 구조화 추출합니다. **고객 정보가 포함될 수 있으므로 회사의 외부 AI 사용 정책에 따라 설정하세요.** 메일 수신 주소와 가격 결정은 AI 추출값을 사용하지 않습니다. 모델명은 정확히 추출됐는지 초안을 검토하세요.
- 키가 없거나 `extractor: "rules"`이면 다음처럼 명확한 줄만 자동 추출합니다. 읽지 못한 자연어 요청은 ‘품목·모델·수량 확인필요’ 초안으로 남깁니다.

```text
ABC-100 / 품목 설명 / 3
XYZ-200 2개
```

- 첨부 지원: `.xlsx`, `.csv`, `.txt`, 텍스트 추출 가능한 `.pdf`. 1파일 3MB 이하, XLSX 200행/20열 이하, PDF 20페이지 이하. 스캔 PDF OCR, 이미지, DOCX, 압축파일, 암호화 문서는 자동 해석하지 않고 확인필요로 남깁니다.
- 표 데이터는 기본적으로 모델 / 설명 / 수량 순서로 작성하면 규칙 추출이 가능합니다. 자유형 표는 AI 추출이 필요할 수 있습니다.
- 분할 납품·납기·결제·재고를 임의로 약속하지 않습니다. 양식의 공급자 정보는 원본 그대로 유지하고 고객 정보를 채웁니다. 고객사가 식별되지 않으면 ‘확인필요’입니다.

## 기록·중복 방지·장애 대응

- `runtime/quotes/<작업ID>/견적서.xlsx`: 실제 첨부된 견적서
- 같은 폴더 `draft.json`: 수신자·제목·본문, `audit.json`: 검색 출처·판매가·시각·품목·검산(내부 검토용), `saved-draft.json`: 초안 ID/링크
- `runtime/state.sqlite3`: 커서·작업 상태. 처리 완료 메일은 재실행해도 중복 작성하지 않습니다.
- `runtime/automation.log`: 작업 ID와 처리 결과. 인증키나 본문 전체를 로그에 출력하지 않습니다.
- Windows 로그인 캐시는 Windows DPAPI로 현재 사용자에게 암호화해 저장합니다. Linux 개발 시에는 권한 600의 캐시 파일을 사용합니다.
- 같은 폴더에서 한 프로세스만 실행됩니다.
- `preparing`/`drafting` 작업은 재시도합니다. 초안 첨부 중 실패하면 같은 파일명 첨부 존재 여부를 확인해 중복을 방지합니다.
- **`creating`/`ambiguous`는 자동 재생성하지 않습니다.** 초안 생성 POST 직후 연결이 끊기면 서버에는 이미 생성됐을 수 있어, Outlook 임시보관함과 `status`를 확인해야 합니다. 복구 명령은 아래 참고.
- API 사용 한도·로그인 만료·조직 동의 오류는 로그와 `doctor`로 확인하세요. 인증 만료 시 `login.bat`을 다시 실행합니다.
- `.env`, `config.json`, `runtime/`은 Git에서 제외했습니다. 실제 메일/견적 기록이 들어 있으므로 저장소에 강제로 추가하지 마세요.

## 테스트

```bat
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

모의 응답으로 마진/VAT/원 단위 반올림, 정확한 모델·규격 필터, 검색 오류와 미확인 가격, 수신자, 여러 품목의 양식 수식 및 저장 검산, 중복 방지와 불확실한 생성 응답 처리를 검증합니다. `demo.bat`의 가격은 가상 테스트값이며 실제 검색가가 아닙니다.

**검증 범위:** 로컬 기능 테스트와 XLSX 저장·재열람 검산. Windows 실제 실행, M365 실계정 초안 생성, 실시간 Google 검색은 API 설정 후 사용자 PC에서 연결 검증이 필요합니다.

## 공식 참고 자료

- [Microsoft Graph createReply](https://learn.microsoft.com/en-us/graph/api/message-createreply?view=graph-rest-1.0)
- [MSAL Python 로그인 흐름](https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens)
- [Google Shopping API / SerpApi](https://serpapi.com/google-shopping-api)
- [검색 결과의 가격·판매처 필드](https://serpapi.com/shopping-results)

## 불명확한 초안 복구

먼저 감시 프로그램을 중지하고 `status`로 작업 ID를 확인하세요. Outlook에서 초안 내용을 직접 확인한 뒤 기존 초안 ID를 지정하면 수신자·초안 여부를 검증해 저장을 이어갑니다.

```bat
.venv\Scripts\python.exe -m quote_automation recover --job 작업ID --draft 기존초안ID
```

초안 ID는 `runtime/quotes/<작업ID>/saved-draft.json`에 저장되어 있으면 거기서 확인할 수 있습니다. 생성 직후 응답이 유실돼 ID를 알 수 없다면 Microsoft Graph Explorer에서 임시보관함의 초안 ID를 확인하거나 IT 담당자의 지원을 받으세요.

기존 초안이 **없음을 직접 확인**한 경우에만 재생성을 대기시킵니다. 기존 초안이 남아 있으면 중복이 생길 수 있습니다.

```bat
.venv\Scripts\python.exe -m quote_automation reset --job 작업ID --confirm-no-draft
.venv\Scripts\python.exe -m quote_automation once
```
