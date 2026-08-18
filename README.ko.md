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

`.xls`는 V1에서 지원하지 않습니다.

## V1 기능

### 검사

- 미사용 `cellXfs` Style 탐지
- 중복 Style 탐지
- Cell/Row/Column의 Style 참조 확인
- Drawing의 `oneCellAnchor` / `twoCellAnchor` 검사
- 기본 2px 이하 작은 객체 탐지

### 정리

- 미사용 Style 제거
- 동일 Style 병합
- Cell/Row/Column Style index 재매핑
- 작은 Drawing 객체 제거
- 원본 자동 백업
- `_clean.xlsx` / `_clean.xlsm` 생성
- 결과 파일 Open XML 검증

원본 파일은 직접 덮어쓰지 않습니다.

## 사용법

1. `Excel 파일 열기`로 `.xlsx` 또는 `.xlsm` 파일을 선택합니다.
2. `검사 시작`으로 정리 후보를 확인합니다.
3. 필요한 정리 옵션을 선택합니다.
4. `정리 실행`을 누르고 결과 파일 위치를 선택합니다.
5. 원본과 같은 폴더의 `ExcelCleaner_Backup`에 백업이 생성됩니다.

## Windows 11 빌드

```powershell
dotnet restore .\QuickExcelCleaner.sln
dotnet build .\QuickExcelCleaner.sln -c Release
dotnet run --project .\src\QuickExcelCleaner\QuickExcelCleaner.csproj -c Release
```

## 통합 테스트

```powershell
dotnet build .\QuickExcelCleaner.sln -c Release
dotnet run --project .\tests\QuickExcelCleaner.Tests\QuickExcelCleaner.Tests.csproj -c Release --no-build
```

실제 `.xlsx` 테스트 파일을 생성하여 다음을 검증합니다.

- 미사용 Style 탐지
- 중복 Style 탐지
- Style index 재매핑
- `cellXfs` 압축
- 백업 생성
- 결과 workbook 검증
- 1px `oneCellAnchor` 객체 제거

## GitHub Actions

Windows runner에서 `.NET SDK 10.0.400`을 설치하고 다음 순서로 실행합니다.

```text
.NET 10.0.400
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

현재 개발 브랜치:

`feature/net10-windows`

PR #1:
https://github.com/theangkko/quick-excel-cleaner/pull/1
