# Workbook safety

Cleanup must be transactional at the file level:

```text
source
  -> backup
  -> working output copy
  -> pre-mutation validation
  -> mutation
  -> post-mutation validation
  -> success
```

If post-mutation validation fails, delete the generated output and keep the source and backup intact.
