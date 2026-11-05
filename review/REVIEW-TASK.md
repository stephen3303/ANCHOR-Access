# Review task (review VMs)

Review only: make no code changes. Post the verdict to the build thread.

For a PR, check out its exact head SHA, then in the `access` folder run:

```
Remove-Item .\build\ANCHOR_BE.accdb -ErrorAction SilentlyContinue
.\backend\new-backend.ps1
if ($?) { .\build.ps1 -BackEnd .\build\ANCHOR_BE.accdb }
```

The data file is made new each time. Open `access\build\ANCHOR.accdb`, sign in (the first Personnel admin becomes ANCHOR's System Admin), and:

1. Code review with `review/code-review-prompt.md`, on the files the PR changes.
2. Visual and functional QA of the pages the PR names with `review/qa-visual-review-prompt.md`: in light, dark and one more palette, at 1280 px wide and wider.
3. Read the error log before and after (Admin, or the data file's `error_log` table).

Report PASS or CHANGES NEEDED, the head SHA tested, and each finding tagged must-fix or optional with file and line.
