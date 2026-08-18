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

## V1 구현 완료 범위

### 분석

- Cell Style 사용/미사용 분석
- 중복 `cellXfs` 분석
- Cell / Row / Column Style 참조 분석
- `oneCellAnchor` / `twoCellAnchor` Drawing 분석
- 기본 2px 이하 작은 객체 탐지

### 정리

- 미사용 Style 제거
- 중복 Style 병합
- Cell / Row / Column Style index 재매핑
- 1px 수준 작은 Drawing 객체 제거
- 원본 백업 생성
- `_clean.xlsx` / `_clean.xlsm` 생성
- 결과 파일 Open XML 재검증
- Style index 무결성 검사
- 실패 시 결과 파일 삭제

### Windows 11 UI

- Excel 파일 열기
- 정리 옵션 선택
- 검사 결과 DataGrid
- 검사 실행
- 결과 파일 저장 위치 선택
- 정리 결과/백업 위치/통계 표시

## 실행

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

실제 `.xlsx`를 생성하여 Style 탐지, Style 재매핑, `cellXfs` 압축, 백업, workbook 검증, 1px 객체 제거를 검증합니다.

## GitHub Actions

Windows runner에서 `.NET SDK 10.0.400`을 사용합니다.

```text
.NET 10.0.400
  ↓
restore
  ↓
Release build
  ↓
Excel cleanup integration tests
  ↓
win-x64 self-contained single-file publish
  ↓
QuickExcelCleaner-win-x64 artifact
```

`workflow_dispatch`로 수동 실행할 수 있습니다.

현재 개발 브랜치: `feature/net10-windows`

PR #1: https://github.com/theangkko/quick-excel-cleaner/pull/1
