# ISSUE-0008｜GitHub Actions 打包驗證紀錄

2026-11-17，Codex 依序執行開發與測試，自查非獨立審查。需求、架構 v0.1 已核准。狀態：進行中，雲端端到端結果待確認。

- 通過：查詢 GitHub ruleset 27431700，target tag、Active、include `refs/tags/excelkit-package/**/*`、限制 update 與 deletion、無 bypass；使用者確認無其他日期序號歷史。
- 通過：遠端新增 `refs/tags/excelkit-package/20261117/01`，保留本機歷史序號；此參照指向導入時基底 1ad27cb，不代表歷史產物的來源提交。唯讀建議下一個版本 `1.2.4.3-20261117-02`。
- 通過：`python Installer/Tests/Test-PackageVersion.py`。模擬 GitHub 保護及命令驗證歷史、序號、重複、離線、規則更動、髒來源與失敗；另以本機裸 Git 儲存庫的雙程序同時 push 驗證同一 tag 僅一個成功。模擬不等同真實 GitHub 的併行壓力測試。
- 通過：`Test-Package-Preflight.ps1` 於獨立測試 fixture 使用虛構 FileVersion 1.2.4.4 驗證有效預檢、無副作用、缺編譯器、版本不一致、缺素材及部分產物衝突。虛構版本僅用於測試，以避開主機已註冊 1.2.4.3；不代表已驗證本機同版本安裝。
- 此前主機原版本預檢 fixture 失敗，原因為 FileVersion 檢查讀到已安裝同版 DLL，fixture 沒有相符 manifest；未更改主機 COM。
- 待執行：實際 GitHub Actions 的依賴取得、正式來源雜湊、Inno Setup 編譯、artifact 與執行中保留序號；完成後需記錄 run URL、提交、輸出與結果。互動安裝、升級、回復、COM、GUI 和 LabVIEW 不由雲端打包成功推定通過。
