# ISSUE-0008｜GitHub Release 資產驗證紀錄 v0.2

2026-11-17，Codex 依序執行 PM、SA、軟體及測試，自查非獨立審查。使用者已核准 v0.2，要求把既有 03 Setup 附上測試版 Release，並於未來成功打包時自動建立測試版 Release。

- 通過：PowerShell 隔離模擬測試 `pwsh -NoProfile -NonInteractive -File Installer/Tests/Test-PublishGitHubRelease.ps1`：draft、上傳、下載雜湊核對及 prerelease；已存在 Release、tag/commit 不符、上傳失敗、資產被改動、來源雜湊錯誤均安全停止。模擬不等於 GitHub 實際執行。
- 通過：`python Installer/Tests/Test-PackageVersion.py` 與裸 Git 併發保留測試；workflow YAML 語法檢查通過。
- 待執行：真實 GitHub 既有 03 的補發、後續新打包的 prerelease、下載確認檔案名稱及 SHA-256。測試版不等於已完成實際安裝、升級、COM、GUI 或 LabVIEW 驗收。
