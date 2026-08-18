# Quick Excel Cleaner engineering entry point

When substantial repository development is performed from chat, read `.agents/skills/luna-chat-coder/SKILL.md` first. The repository inherits Luna Chat Coder's sandbox-first and exact-source policies, then adds the project-specific rules below.

## Runtime and platform

- Target framework: `net10.0-windows`
- SDK: `.NET 10.0.400`
- UI: WPF
- CI OS: Windows
- Primary runtime/package target: `win-x64`

## Excel contract

- Supported input formats: `.xlsx` and `.xlsm`
- Excel manipulation: Open XML SDK
- Never modify the source workbook in place.
- Create a backup before cleanup.
- Do not report success until the generated workbook passes validation.
- Delete an invalid generated output instead of returning it as a successful result.

## Style cleanup contract

Style cleanup must account for all workbook references, not only Cell styles:

- Cell `StyleIndex`
- Row `StyleIndex`
- Column `Style`
- duplicate `cellXfs`
- unused `cellXfs`

Duplicate Style canonicalization must be performed before unused-style removal can discard a member of a duplicate group. Style indexes are implementation references and are not stable semantic identifiers.

## Object cleanup contract

Small drawing cleanup must consider both `oneCellAnchor` and `twoCellAnchor`. The configured pixel threshold is a policy value, not a license to remove normal-size objects.

## Workbook preservation contract

Cleanup must preserve workbook semantics, including where present:

- worksheet count and worksheet names
- cell values
- merged cells
- hidden rows and columns
- freeze panes
- conditional formatting
- named styles
- ordinary drawing objects
- VBA payloads in `.xlsm`

## Verification contract

The applicable verification path is:

```text
Build
  -> Scanner tests
  -> Cleanup tests
  -> Complex workbook tests
  -> XLSM/VBA preservation tests
  -> Feature preservation tests
  -> Publish validation
```

Do not weaken tests by changing expected values without first determining whether the product logic or test expectation is semantically wrong.

## Release contract

Application version is repository-defined metadata in `QuickExcelCleaner.csproj`. Releases use a `v<Version>` tag, never overwrite an existing tag, and publish only artifacts produced by a commit that has passed build and integration tests.
