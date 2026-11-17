# ISSUE-0008｜GitHub Actions 打包驗證紀錄

2026-11-17，Codex 依序執行開發與測試，自查非獨立審查。需求、架構 v0.1 已核准。狀態：進行中，雲端端到端結果待確認。

- 通過：查詢 GitHub ruleset 27431700，target tag、Active、include `refs/tags/excelkit-package/**/*`、限制 update 與 deletion、無 bypass；使用者確認無其他日期序號歷史。
- 通過：遠端新增 `refs/tags/excelkit-package/20261117/01`，保留本機歷史序號；此參照指向導入時基底 1ad27cb，不代表歷史產物的來源提交。唯讀建議下一個版本 `1.2.4.3-20261117-02`。
- 通過：`python Installer/Tests/Test-PackageVersion.py`。模擬 GitHub 保護及命令驗證歷史、序號、重複、離線、規則更動、髒來源與失敗；另以本機裸 Git 儲存庫的雙程序同時 push 驗證同一 tag 僅一個成功。模擬不等同真實 GitHub 的併行壓力測試。
- 通過：`Test-Package-Preflight.ps1` 於獨立測試 fixture 使用虛構 FileVersion 1.2.4.4 驗證有效預檢、無副作用、缺編譯器、版本不一致、缺素材及部分產物衝突。虛構版本僅用於測試，以避開主機已註冊 1.2.4.3；不代表已驗證本機同版本安裝。
- 此前主機原版本預檢 fixture 失敗，原因為 FileVersion 檢查讀到已安裝同版 DLL，fixture 沒有相符 manifest；未更改主機 COM。
- 待執行：實際 GitHub Actions 的依賴取得、正式來源雜湊、Inno Setup 編譯、artifact 與執行中保留序號；完成後需記錄 run URL、提交、輸出與結果。互動安裝、升級、回復、COM、GUI 和 LabVIEW 不由雲端打包成功推定通過。

## GitHub Actions 實際執行
2026-11-17，run 44018954715、attempt 1、commit `0e3080f707adbbd7436a754f8514a43f8507104d`：Windows runner 所有步驟成功，包含 Inno Setup 7.1.1 雜湊與 Authenticode、.NET SDK 10.0.403、版本測試、預檢、實際建置與 artifact 上傳。安裝器 `ExcelKit-Setup-1.2.4.3-20261117-02.exe` SHA-256 `64849CB0B29B1E4EFAA14995D3E81B108F7A342C7ECC482FC4FD26D1965A2F3E`；manifest SHA-256 `61CFC8D78D572C230355A8A849FFE9E0E4318EFAC0BBC62F217D2604BBDBC8F3`。artifact `ExcelKit-installer-44018954715-1`，GitHub 回報 ZIP SHA-256 `64a1d85e21765434102aa99b4352bb1d397c777417f6b4f0890a38a46e3ff932`；GitHub 上的 `20261117/02` 指向該 commit。執行證據：https://github.com/Ryan-tw-Hans/ExcelKit/actions/runs/44018954715 。雲端測試包未進行安裝、COM、GUI、LabVIEW 或升級回復測試。
