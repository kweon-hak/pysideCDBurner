# Translation Guide (INI)

## 한글 안내
- 위치: `locales/<언어코드>.ini` (UTF-8). 예: `locales/ko.ini`.
- 구성:
  - `[meta]` 섹션(선택): `name`(표시 이름), `code`(언어 코드, 기본은 파일명).
  - `[translations]` 섹션: `영문 키 = 번역 값`.
- 규칙:
  - 왼쪽 영문 키는 UI 문자열과 완전히 동일해야 합니다. 공백·콜론 등도 그대로 유지합니다. (`" Stop"`처럼 앞에 공백이 있는 키도 그대로 두세요.)
  - `{size}`, `{duration}` 등 플레이스홀더는 값에서도 그대로 남겨둡니다.
  - 줄바꿈은 `\n`, 탭은 `\t`로 표기합니다.
  - 찾기/교체 시 `=` 오른쪽 값만 수정합니다.
- 새 언어 추가:
  1) `locales/ko.ini`를 복사해 `locales/<code>.ini`로 저장  
  2) `[meta]`의 `name`, `code`를 새 언어에 맞게 수정  
  3) `[translations]`의 값만 번역  
  4) 앱 재시작 → 언어 메뉴에 자동 추가 (INI 로드 실패나 키 누락 시 영어로 폴백)

## English guide
- Location: `locales/<lang>.ini` (UTF-8). Example: `locales/ko.ini`.
- Structure:
  - `[meta]` (optional): `name` (display name), `code` (lang code, defaults to filename).
  - `[translations]`: `English key = translated value`.
- Rules:
  - Left side keys must match the UI English text exactly (keep spaces/colons; even `" Stop"` with a leading space).
  - Keep placeholders like `{size}` or `{duration}` in the value.
  - Use `\n` for newlines and `\t` for tabs.
  - Only edit the value to the right of `=`.
- Add a new language:
  1) Copy `locales/ko.ini` to `locales/<code>.ini`  
  2) Edit `name`/`code` in `[meta]`  
  3) Translate only the values in `[translations]`  
  4) Restart the app; the language appears automatically (English is used as fallback on errors/missing keys)
