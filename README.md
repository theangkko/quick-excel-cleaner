# Quick Excel Cleaner

Windows 11용 Excel 정리 도구입니다.

## 기술 스택

- C# / WPF
- .NET 10 (`net10.0-windows`)
- DocumentFormat.OpenXml
- Excel 설치 불필요

## 현재 개발 기준

`feature/net10-windows` 브랜치에서 .NET 10 Windows 데스크톱 앱으로 개발합니다.

현재 V1은 다음 기능의 기반을 제공합니다.

- `.xlsx`, `.xlsm` 파일 선택
- 미사용 Cell Style 검사
- 중복 Cell Style 검사
- Drawing 객체 검사
- 작은 객체 후보 표시(기본 2 px 기준)
- 검사 단계에서는 원본 파일을 수정하지 않음

## 빌드

```powershell
dotnet restore .\QuickExcelCleaner.sln
dotnet build .\QuickExcelCleaner.sln -c Release
dotnet run --project .\src\QuickExcelCleaner\QuickExcelCleaner.csproj
```

## 다음 단계

검사 결과를 바탕으로 백업을 만든 뒤 정리된 복사본을 생성합니다. Style 정리는 실제 참조 관계를 보존하면서 `cellXfs`를 재매핑하고, 작은 객체 삭제도 사용자가 확인할 수 있도록 후보 목록을 먼저 제공합니다.
