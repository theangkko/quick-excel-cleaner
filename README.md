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

## V1 기능

### 검사

- 사용 중 Cell Style 검사
- 미사용 `cellXfs` 검사
- 중복 Cell Style 검사
- Cell / Row / Column Style 참조 확인
- Drawing의 `oneCellAnchor` / `twoCellAnchor` 검사
- 기본 2px 이하 작은 객체 후보 탐지

### 정리

- 미사용 Style 제거
- 동일한 Style 병합
- Cell / Row / Column Style index 재매핑
- 1px 수준의 작은 Drawing 객체 제거
- 원본 자동 백업
- `*_clean.xlsx` 또는 `*_clean.xlsm` 생성
- 결과 파일 Open XML 검증
- 잘못된 Style index 검출

원본 파일은 정리 과정에서 덮어쓰지 않습니다.

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

워크플로에는 `workflow_dispatch`도 포함되어 있어 GitHub Actions 화면에서 수동 실행할 수 있습니다.

## 안전성 원칙

Style은 단순 번호 삭제가 아니라 실제 Cell / Row / Column 참조를 먼저 수집한 뒤 대표 Style을 선택하고 모든 참조를 재매핑합니다.

정리된 파일은 다시 Open XML로 열어 Worksheet와 Style index 관계를 검증합니다. 검증에 실패하면 생성된 결과 파일을 폐기하고 원본에는 아무 변경도 하지 않습니다.

## 개발 브랜치

```text
feature/net10-windows
```

PR:

https://github.com/theangkko/quick-excel-cleaner/pull/1
