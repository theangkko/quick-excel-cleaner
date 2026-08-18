# Quick Excel Cleaner

Windows 11용 Excel 정리 도구입니다.

## 기술 스택

- C# / WPF
- .NET 10 (`net10.0-windows`)
- DocumentFormat.OpenXml 3.3.0
- Excel 설치 불필요

## 지원 파일

- `.xlsx`
- `.xlsm`

원본 Excel을 직접 열어 조작하지 않고 Open XML 패키지를 분석하고 수정합니다. `.xls`는 V1에서 지원하지 않습니다.

## 현재 V1 기능

### 검사

- 사용 중 Cell Style 검사
- 미사용 `cellXfs` 검사
- 중복 Cell Style 검사
- Cell뿐 아니라 Row/Column Style 참조 확인
- Drawing의 `oneCellAnchor` / `twoCellAnchor` 검사
- 기본 2px 이하 작은 객체 후보 탐지

### 정리

- 미사용 Style 제거
- 동일한 Style 병합
- Cell Style index 재매핑
- Row/Column Style index 재매핑
- 1px 수준의 작은 Drawing 객체 제거
- 원본 자동 백업
- `*_clean.xlsx` 또는 `*_clean.xlsm` 생성
- 결과 파일 Open XML 검증

원본 파일은 정리 과정에서 덮어쓰지 않습니다.

## Windows 11 빌드

PowerShell:

```powershell
dotnet restore .\QuickExcelCleaner.sln
dotnet build .\QuickExcelCleaner.sln -c Release
dotnet run --project .\src\QuickExcelCleaner\QuickExcelCleaner.csproj -c Release
```

## 테스트

통합 테스트는 실제 `.xlsx`를 메모리/임시 파일로 생성하여 검사합니다.

```powershell
dotnet build .\QuickExcelCleaner.sln -c Release
dotnet run --project .\tests\QuickExcelCleaner.Tests\QuickExcelCleaner.Tests.csproj -c Release --no-build
```

테스트에는 다음 항목이 포함됩니다.

- 미사용 Style 탐지
- 중복 Style 탐지
- Style index 재매핑
- `cellXfs` 압축
- 백업 생성
- 결과 workbook 검증
- 1px `oneCellAnchor` 객체 제거

## GitHub Actions

`.github/workflows/build.yml`에서 Windows runner에 `.NET SDK 10.0.400`을 설치합니다.

CI는 다음 순서로 실행합니다.

```text
.NET SDK 10.0.400
      ↓
restore
      ↓
Release build
      ↓
Excel cleanup integration tests
      ↓
win-x64 self-contained publish
      ↓
QuickExcelCleaner-win-x64 artifact
```

## 안전성 원칙

Style은 단순 번호 삭제가 아니라 실제 Cell/Row/Column 참조를 먼저 수집한 뒤 대표 Style을 선택하고 모든 참조를 재매핑합니다.

정리된 파일은 다시 Open XML로 열어 Worksheet와 Style index 관계를 검증합니다. 검증에 실패하면 생성된 결과 파일을 폐기하고 원본에는 아무 변경도 하지 않습니다.

## 개발 브랜치

현재 개발 기준 브랜치:

```text
feature/net10-windows
```

PR:

https://github.com/theangkko/quick-excel-cleaner/pull/1
