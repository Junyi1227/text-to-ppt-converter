# 如何在 Mac 上打包 Windows .exe 執行檔

## 🎯 目標

您在 Mac 上開發，需要為 Windows 用戶製作一個**無需安裝 Python** 的獨立執行檔（.exe）。

---

## ⚠️ 重要提醒：跨平台打包的限制

**Mac 無法直接打包 Windows .exe！**

PyInstaller 只能在目標平台上打包：
- Mac 上只能打包 Mac 執行檔
- Windows 上才能打包 Windows .exe
- Linux 上只能打包 Linux 執行檔

---

## 🛠️ 三種解決方案

### 方案 1：使用 Windows 虛擬機（推薦給個人使用）

#### 使用 Parallels Desktop / VMware Fusion

1. **在 Mac 上安裝 Windows 虛擬機**
   - Parallels Desktop（付費，但效能最好）
   - VMware Fusion（有免費版本）
   - VirtualBox（免費，但效能較差）

2. **在 Windows 虛擬機中執行打包**
   ```cmd
   # 在 Windows 虛擬機中
   pip install pyinstaller python-pptx
   python build_windows_exe.py
   ```

3. **優點**：
   - ✅ 完全控制打包過程
   - ✅ 可以立即測試執行檔
   - ✅ 不需要其他服務

4. **缺點**：
   - ❌ 需要購買虛擬機軟體（或使用免費版）
   - ❌ 需要 Windows 授權
   - ❌ 佔用硬碟空間

---

### 方案 2：使用 GitHub Actions（推薦給開源專案）✨

這是**最推薦**的自動化方案！完全免費，無需本地 Windows 環境。

#### 步驟：

1. **建立 GitHub repository**
   ```bash
   cd 您的專案資料夾
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/您的帳號/專案名稱.git
   git push -u origin main
   ```

2. **建立 GitHub Actions 工作流程**
   
   我已經為您準備好了 `.github/workflows/build.yml` 檔案（見下方）

3. **觸發自動打包**
   - 推送程式碼到 GitHub
   - GitHub Actions 會自動在 Windows 環境中打包
   - 完成後從 Releases 或 Artifacts 下載 .exe

4. **優點**：
   - ✅ 完全免費
   - ✅ 自動化，推送即打包
   - ✅ 支援多平台（可同時打包 Mac/Windows/Linux）
   - ✅ 無需本地 Windows 環境

5. **缺點**：
   - ❌ 需要 GitHub 帳號
   - ❌ 需要公開 repository（或付費使用私有 repo）
   - ❌ 每次打包需要幾分鐘

---

### 方案 3：借用 Windows 電腦（最簡單）

如果您有 Windows 電腦可用：

1. **複製檔案到 Windows 電腦**
   - 將所有 `.py` 檔案複製過去

2. **在 Windows 上執行**
   ```cmd
   # 安裝 Python（如果沒有）
   # 從 https://www.python.org/downloads/ 下載安裝
   
   # 安裝必要套件
   pip install pyinstaller python-pptx
   
   # 執行打包
   python build_windows_exe.py
   ```

3. **取得執行檔**
   - 打包完成後會在 `dist/` 資料夾
   - 複製 `文字轉PPT工具.exe` 回 Mac

---

## 📦 GitHub Actions 自動打包設定

建立檔案：`.github/workflows/build.yml`

```yaml
name: Build Windows Executable

on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]
  workflow_dispatch:  # 允許手動觸發

jobs:
  build:
    runs-on: windows-latest
    
    steps:
    - name: Checkout code
      uses: actions/checkout@v3
    
    - name: Set up Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.11'
    
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install pyinstaller python-pptx
    
    - name: Build executable
      run: |
        pyinstaller --onefile --windowed --name=文字轉PPT工具 text_to_ppt_gui.py
    
    - name: Upload artifact
      uses: actions/upload-artifact@v3
      with:
        name: Windows-Executable
        path: dist/文字轉PPT工具.exe
    
    - name: Create Release (on tag)
      if: startsWith(github.ref, 'refs/tags/')
      uses: softprops/action-gh-release@v1
      with:
        files: dist/文字轉PPT工具.exe
      env:
        GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
```

### 使用 GitHub Actions 的步驟：

```bash
# 1. 在專案根目錄建立目錄結構
mkdir -p .github/workflows

# 2. 將上面的內容儲存為 .github/workflows/build.yml

# 3. 提交到 GitHub
git add .github/workflows/build.yml
git commit -m "Add GitHub Actions build workflow"
git push

# 4. 前往 GitHub repository → Actions 頁面
# 5. 等待打包完成（約 3-5 分鐘）
# 6. 下載 Artifacts 中的執行檔
```

---

## 🚀 快速開始指南（推薦流程）

### 給 Mac 開發者的最佳實踐：

1. **開發階段**
   ```bash
   # 在 Mac 上開發和測試（使用 Python 版本）
   python3 text_to_ppt_gui.py
   ```

2. **打包階段**
   - 選項 A：推送到 GitHub，使用 GitHub Actions 自動打包
   - 選項 B：使用 Windows 虛擬機打包

3. **發布階段**
   ```bash
   # 建立 release
   git tag v1.0.0
   git push origin v1.0.0
   
   # GitHub Actions 會自動建立 Release 並附上 .exe
   ```

4. **交付給 Windows 用戶**
   - 提供 `文字轉PPT工具.exe`
   - 提供 `Windows使用說明.txt`
   - Windows 用戶無需安裝任何東西，雙擊即可使用！

---

## 📋 檔案清單（打包所需）

確保以下檔案存在：

```
專案根目錄/
├── text_to_ppt_gui.py          # GUI 版本主程式
├── text_to_ppt.py              # 命令列版本（選用）
├── build_windows_exe.py        # 打包腳本
├── 範例輸入文字.txt             # 範例檔案
├── .github/
│   └── workflows/
│       └── build.yml           # GitHub Actions 設定
└── requirements.txt            # Python 相依套件
```

建立 `requirements.txt`：
```
python-pptx==0.6.21
pyinstaller==6.3.0
```

---

## 🧪 測試清單

打包完成後，請在 Windows 上測試：

- [ ] 執行檔可以雙擊開啟
- [ ] GUI 介面正常顯示
- [ ] 可以輸入文字
- [ ] 可以載入 .txt 檔案
- [ ] 可以轉換並儲存 .pptx
- [ ] 產生的 PPT 可以用 PowerPoint 開啟
- [ ] 投影片格式正確（顏色、字型、排版）

---

## 💡 建議

**對於個人專案或小型團隊：**
- 使用 **GitHub Actions**（免費、自動化）

**對於商業專案或大型團隊：**
- 投資 **Parallels Desktop**（約 $99/年）
- 或設定專用的 Windows 打包機器

**對於臨時需求：**
- 借用 Windows 電腦打包
- 或使用雲端 Windows 環境（如 AWS Windows EC2）

---

## ⚙️ 打包選項說明

在 `build_windows_exe.py` 中的 PyInstaller 參數：

```python
pyinstaller \
  --onefile \                    # 打包成單一 .exe（不是多個檔案）
  --windowed \                   # GUI 程式，不顯示命令列視窗
  --name=文字轉PPT工具 \          # 執行檔名稱
  --add-data=範例檔案.txt;. \    # 包含額外檔案
  --icon=icon.ico \              # 自訂圖示（選用）
  text_to_ppt_gui.py
```

如果想要**更小的執行檔**（但會分散成多個檔案）：
```python
# 移除 --onefile，改用 --onedir
pyinstaller --onedir --windowed text_to_ppt_gui.py
```

---

## 📞 下一步

現在您可以：

1. ✅ 選擇一個打包方案（建議：GitHub Actions）
2. ✅ 執行打包流程
3. ✅ 測試產生的 .exe
4. ✅ 交付給 Windows 用戶

**推薦流程：**
```bash
# 在 Mac 上
git init
git add .
git commit -m "Initial commit"

# 推送到 GitHub（會觸發自動打包）
git remote add origin https://github.com/你的帳號/text-to-ppt.git
git push -u origin main

# 等待 3-5 分鐘，前往 GitHub Actions 查看結果
# 下載打包好的 .exe 檔案
```

需要協助設定 GitHub Actions 嗎？或想了解其他打包選項？
