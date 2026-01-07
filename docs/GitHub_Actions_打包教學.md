# GitHub Actions 自動打包教學

## 🎯 目標
在 Mac 上推送程式碼到 GitHub，自動在雲端 Windows 環境中打包成 .exe

---

## 📋 前置準備

### 必要條件：
- ✅ GitHub 帳號（免費即可）
- ✅ Git 已安裝（Mac 通常已內建）
- ✅ 本專案的所有檔案

### 檢查 Git 是否已安裝：
```bash
git --version
```
如果沒有，執行：
```bash
xcode-select --install
```

---

## 🚀 完整步驟（10 分鐘完成）

### 步驟 1：在 GitHub 建立新的 Repository

1. **登入 GitHub**  
   前往：https://github.com

2. **建立新 repository**
   - 點擊右上角 `+` → `New repository`
   - Repository name: `text-to-ppt-converter`（或您喜歡的名稱）
   - Description: `文字轉 PowerPoint 工具`
   - ⚠️ **選擇 Public**（免費使用 GitHub Actions）
   - ❌ **不要**勾選 "Add a README file"（我們已經有了）
   - 點擊 `Create repository`

3. **記下您的 repository URL**
   ```
   https://github.com/您的帳號/text-to-ppt-converter.git
   ```

---

### 步驟 2：在 Mac 上初始化 Git（在專案目錄執行）

```bash
# 1. 進入專案目錄（假設在桌面）
cd ~/Desktop/text-to-ppt-converter

# 2. 初始化 Git repository
git init

# 3. 設定您的 Git 身份（如果還沒設定過）
git config --global user.name "您的名字"
git config --global user.email "您的Email"

# 4. 建立 .gitignore 檔案（避免上傳不必要的檔案）
cat > .gitignore << 'EOF'
# Python
__pycache__/
*.pyc
*.pyo
*.egg-info/
dist/
build/
*.spec

# macOS
.DS_Store

# PyInstaller
*.spec

# 編輯器
.vscode/
.idea/
*.swp

# 測試檔案
tmp_*
test_*.pptx
EOF

# 5. 將所有檔案加入 Git
git add .

# 6. 建立第一個 commit
git commit -m "Initial commit: Text to PowerPoint Converter"

# 7. 連結到 GitHub（替換成您的 URL）
git remote add origin https://github.com/您的帳號/text-to-ppt-converter.git

# 8. 推送到 GitHub
git branch -M main
git push -u origin main
```

### 如果推送時要求輸入帳號密碼：

GitHub 現在需要使用 Personal Access Token（不能用密碼）

**產生 Token**：
1. GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)
2. Generate new token (classic)
3. 勾選 `repo` 權限
4. 產生後複製 token（只會顯示一次！）
5. 推送時，帳號用您的 GitHub 帳號，密碼用這個 token

---

### 步驟 3：檢查 GitHub Actions 是否開始執行

1. **前往您的 GitHub repository**
   ```
   https://github.com/您的帳號/text-to-ppt-converter
   ```

2. **點擊 "Actions" 標籤**
   - 應該會看到一個正在執行的工作流程
   - 名稱：`Build Windows Executable`
   - 狀態：🟡 黃色圓圈（執行中）

3. **點擊該工作流程查看詳細進度**
   - 可以即時看到執行日誌
   - 整個過程約 3-5 分鐘

4. **等待完成**
   - 成功：✅ 綠色勾勾
   - 失敗：❌ 紅色叉叉（如果失敗，請往下看疑難排解）

---

### 步驟 4：下載打包好的 .exe 檔案

打包完成後，有兩種下載方式：

#### 方法 A：從 Artifacts 下載（每次推送都可用）

1. 在 Actions 頁面，點擊完成的工作流程
2. 往下捲動到 "Artifacts" 區域
3. 會看到：
   - `Windows-GUI-Executable` - GUI 版本（推薦）
   - `Windows-CLI-Executable` - 命令列版本
4. 點擊下載（會下載成 .zip 檔案）
5. 解壓縮後就能看到 `文字轉PPT工具.exe`

#### 方法 B：建立 Release（推薦給正式發布）

```bash
# 在 Mac 上建立版本標籤
git tag v1.0.0
git push origin v1.0.0

# GitHub Actions 會自動建立 Release
# 前往 repository → Releases 查看
```

Release 的優點：
- ✅ 更正式的發布方式
- ✅ 可以附上版本說明
- ✅ 方便其他人下載
- ✅ 檔案永久保存（Artifacts 會過期）

---

## 🧪 測試 .exe 檔案

### 在 Mac 上無法測試（需要 Windows）

選項 1：使用 Windows 虛擬機
```bash
# 在 Parallels/VMware 的 Windows 中
# 1. 複製 .exe 到 Windows
# 2. 雙擊執行
# 3. 測試所有功能
```

選項 2：請 Windows 使用者幫忙測試

選項 3：使用雲端 Windows（如果有 AWS/Azure 帳號）

---

## 🔄 日常工作流程

### 修改程式碼後，重新打包：

```bash
# 1. 修改程式碼（例如 text_to_ppt_gui.py）
# 用您喜歡的編輯器修改

# 2. 測試（在 Mac 上）
python3 text_to_ppt_gui.py

# 3. 提交變更
git add .
git commit -m "修改功能：新增 XXX"

# 4. 推送到 GitHub（自動觸發打包）
git push

# 5. 前往 Actions 查看打包進度
# 6. 下載新的 .exe
```

**每次推送都會自動打包！** 🎉

---

## 📊 GitHub Actions 執行流程解析

當您推送程式碼時，GitHub Actions 會：

```yaml
1. 啟動 Windows Server 虛擬機（雲端）
   ↓
2. 安裝 Python 3.11
   ↓
3. 下載您的程式碼
   ↓
4. 安裝相依套件（python-pptx, pyinstaller）
   ↓
5. 執行 PyInstaller 打包
   - GUI 版本 → 文字轉PPT工具.exe
   - CLI 版本 → text_to_ppt.exe
   ↓
6. 上傳 Artifacts（可下載）
   ↓
7. （如果是 tag）建立 Release
   ↓
8. 關閉虛擬機
```

**總時間**：3-5 分鐘  
**費用**：免費（公開 repository 每月 2000 分鐘額度）

---

## ❗ 疑難排解

### 問題 1：推送時出現 "Permission denied"

**原因**：沒有權限或使用密碼（GitHub 已禁用密碼）

**解決**：
```bash
# 使用 SSH（推薦）
ssh-keygen -t ed25519 -C "您的email"
cat ~/.ssh/id_ed25519.pub  # 複製輸出

# 前往 GitHub → Settings → SSH and GPG keys → New SSH key
# 貼上公鑰

# 修改 remote URL
git remote set-url origin git@github.com:您的帳號/text-to-ppt-converter.git
git push
```

或使用 Personal Access Token（見步驟 2）

---

### 問題 2：GitHub Actions 失敗（紅色叉叉）

**檢查步驟**：
1. 點擊失敗的工作流程
2. 查看錯誤訊息（通常是紅色文字）
3. 常見原因：
   - 檔案路徑錯誤
   - Python 套件版本不相容
   - 語法錯誤

**常見解決方法**：
```bash
# 檢查 .github/workflows/build.yml 是否正確
cat .github/workflows/build.yml

# 確認所有必要檔案都已上傳
git ls-files

# 如果是套件問題，修改 requirements.txt
```

---

### 問題 3：找不到 Artifacts

**可能原因**：
- 工作流程還在執行中（等待完成）
- 工作流程失敗了
- Artifacts 過期（90 天後自動刪除）

**解決**：
- 確認工作流程已完成（綠色勾勾）
- 如果失敗，查看日誌並修復
- 使用 Release 代替 Artifacts（不會過期）

---

### 問題 4：.exe 檔案在 Windows 上無法執行

**常見原因**：
1. **Windows Defender 封鎖**
   - 第一次執行會警告
   - 點選「更多資訊」→「仍要執行」

2. **缺少 Visual C++ 運行庫**（罕見）
   - 下載安裝：https://aka.ms/vs/17/release/vc_redist.x64.exe

3. **權限問題**
   - 右鍵 → 以系統管理員身分執行

---

### 問題 5：.exe 檔案太大

**正常大小**：20-30 MB

**如果想縮小**（需修改 build.yml）：
```yaml
# 使用 UPX 壓縮（可能不穩定）
- name: Compress with UPX
  run: |
    choco install upx
    upx --best dist/文字轉PPT工具.exe
```

**注意**：壓縮可能導致某些防毒軟體誤報

---

## 🎨 自訂 GitHub Actions

### 修改觸發條件

編輯 `.github/workflows/build.yml`：

```yaml
# 只在推送到 main 分支時打包
on:
  push:
    branches: [ main ]

# 或：只在建立 tag 時打包
on:
  push:
    tags:
      - 'v*'

# 或：手動觸發
on:
  workflow_dispatch:
```

### 增加更多平台

```yaml
jobs:
  build-windows:
    runs-on: windows-latest
    # ... Windows 打包
  
  build-mac:
    runs-on: macos-latest
    # ... Mac 打包
  
  build-linux:
    runs-on: ubuntu-latest
    # ... Linux 打包
```

---

## 📦 完整指令速查表

```bash
# 初次設定
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/您的帳號/repo名稱.git
git push -u origin main

# 日常更新
git add .
git commit -m "更新說明"
git push

# 建立 Release
git tag v1.0.0
git push origin v1.0.0

# 查看狀態
git status
git log --oneline

# 撤銷變更
git checkout -- 檔案名稱  # 撤銷單一檔案
git reset --hard HEAD     # 撤銷所有未提交的變更
```

---

## 🎯 檢查清單

打包前確認：

- [ ] 所有 Python 檔案可在 Mac 上正常執行
- [ ] `.github/workflows/build.yml` 存在且格式正確
- [ ] `requirements.txt` 包含所有相依套件
- [ ] `.gitignore` 已建立（避免上傳臨時檔案）
- [ ] GitHub repository 已建立
- [ ] 本地已連結到 GitHub remote
- [ ] 已成功推送到 GitHub

打包後確認：

- [ ] GitHub Actions 執行成功（綠色勾勾）
- [ ] Artifacts 可下載
- [ ] .exe 檔案可在 Windows 上執行（需 Windows 測試）
- [ ] 所有功能正常運作

---

## 💡 最佳實踐

1. **使用有意義的 commit message**
   ```bash
   # 好的範例
   git commit -m "新增：支援圖片插入功能"
   git commit -m "修復：Mac 輸入框多行問題"
   git commit -m "優化：減少執行檔大小"
   
   # 不好的範例
   git commit -m "update"
   git commit -m "fix"
   ```

2. **使用版本標籤**
   ```bash
   git tag v1.0.0  # 主要版本
   git tag v1.1.0  # 新功能
   git tag v1.1.1  # 修復 bug
   ```

3. **定期測試**
   - 每次重要修改後都重新打包測試
   - 在真實 Windows 環境測試

4. **備份重要版本**
   - 使用 Release 功能
   - 保存穩定版本的 .exe

---

## 🚀 下一步

完成打包後，您可以：

1. **測試 .exe 檔案**（需要 Windows）
2. **分發給使用者**
   - 提供 `文字轉PPT工具.exe`
   - 提供 `Windows用戶使用說明.txt`
3. **建立正式 Release**
   - 附上版本說明
   - 列出新功能和修復的問題
4. **收集使用者回饋**
   - 持續改進功能
   - 修復發現的問題

---

## 📞 需要協助？

如果在步驟中遇到問題：

1. **檢查 GitHub Actions 日誌**
   - 通常會有詳細的錯誤訊息

2. **常見錯誤關鍵字**
   - `Permission denied` → SSH/Token 問題
   - `Module not found` → requirements.txt 缺少套件
   - `Syntax error` → Python 程式碼有錯誤

3. **測試本地是否正常**
   ```bash
   # 在 Mac 上測試 Python 程式
   python3 text_to_ppt_gui.py
   ```

準備好開始了嗎？執行步驟 1 開始設定！🚀
