#!/bin/bash
# 文字轉 PowerPoint 工具 - 快速開始腳本
# 自動設定 Git 和推送到 GitHub

set -e  # 遇到錯誤就停止

echo "========================================"
echo "文字轉 PowerPoint - GitHub Actions 設定"
echo "========================================"
echo ""

# 檢查是否在正確的目錄
if [ ! -f "text_to_ppt_gui.py" ]; then
    echo "❌ 錯誤：找不到 text_to_ppt_gui.py"
    echo "請在專案根目錄執行此腳本"
    exit 1
fi

# 檢查 Git 是否已安裝
if ! command -v git &> /dev/null; then
    echo "❌ 錯誤：Git 未安裝"
    echo "請執行：xcode-select --install"
    exit 1
fi

echo "✅ Git 已安裝"
echo ""

# 詢問 GitHub 資訊
echo "請輸入您的 GitHub 資訊："
echo ""
read -p "GitHub 帳號名稱: " GITHUB_USERNAME
read -p "Repository 名稱 [text-to-ppt-converter]: " REPO_NAME
REPO_NAME=${REPO_NAME:-text-to-ppt-converter}

read -p "您的名字（用於 Git commit）: " GIT_NAME
read -p "您的 Email（用於 Git commit）: " GIT_EMAIL

echo ""
echo "設定資訊："
echo "  GitHub 帳號: $GITHUB_USERNAME"
echo "  Repository: $REPO_NAME"
echo "  Git 名字: $GIT_NAME"
echo "  Git Email: $GIT_EMAIL"
echo ""
read -p "確認以上資訊正確嗎？(y/n) " -n 1 -r
echo ""

if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    echo "已取消"
    exit 1
fi

echo ""
echo "步驟 1: 設定 Git..."

# 檢查是否已經是 Git repository
if [ -d ".git" ]; then
    echo "⚠️  已經是 Git repository，跳過初始化"
else
    git init
    echo "✅ Git repository 已初始化"
fi

# 設定 Git 使用者資訊
git config user.name "$GIT_NAME"
git config user.email "$GIT_EMAIL"
echo "✅ Git 使用者資訊已設定"

echo ""
echo "步驟 2: 建立 .gitignore..."

# 建立 .gitignore
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
tmp_rovodev_*
test_*.pptx
EOF

echo "✅ .gitignore 已建立"

echo ""
echo "步驟 3: 加入所有檔案到 Git..."

git add .
echo "✅ 檔案已加入"

echo ""
echo "步驟 4: 建立第一個 commit..."

git commit -m "Initial commit: Text to PowerPoint Converter"
echo "✅ Commit 已建立"

echo ""
echo "步驟 5: 設定 GitHub remote..."

# 檢查是否已有 remote
if git remote | grep -q "origin"; then
    echo "⚠️  Remote 'origin' 已存在"
    git remote set-url origin "https://github.com/$GITHUB_USERNAME/$REPO_NAME.git"
    echo "✅ Remote URL 已更新"
else
    git remote add origin "https://github.com/$GITHUB_USERNAME/$REPO_NAME.git"
    echo "✅ Remote 已加入"
fi

echo ""
echo "步驟 6: 準備推送到 GitHub..."
echo ""
echo "⚠️  重要提醒："
echo "1. 請先在 GitHub 建立 repository: $REPO_NAME"
echo "2. 前往：https://github.com/new"
echo "3. Repository name 填入: $REPO_NAME"
echo "4. 選擇 Public（才能免費使用 GitHub Actions）"
echo "5. 不要勾選 'Add a README file'（我們已經有了）"
echo "6. 建立完成後，按 Enter 繼續..."
echo ""
read -p "按 Enter 繼續推送，或按 Ctrl+C 取消..."

echo ""
echo "正在推送到 GitHub..."

# 重命名分支為 main（如果需要）
CURRENT_BRANCH=$(git branch --show-current)
if [ "$CURRENT_BRANCH" != "main" ]; then
    git branch -M main
fi

# 推送
if git push -u origin main; then
    echo ""
    echo "========================================"
    echo "🎉 成功！"
    echo "========================================"
    echo ""
    echo "下一步："
    echo "1. 前往查看 GitHub Actions："
    echo "   https://github.com/$GITHUB_USERNAME/$REPO_NAME/actions"
    echo ""
    echo "2. 等待 3-5 分鐘打包完成"
    echo ""
    echo "3. 下載 .exe 檔案："
    echo "   點擊完成的工作流程 → Artifacts → 下載"
    echo ""
    echo "4. 或建立 Release："
    echo "   git tag v1.0.0"
    echo "   git push origin v1.0.0"
    echo ""
else
    echo ""
    echo "❌ 推送失敗"
    echo ""
    echo "可能原因："
    echo "1. Repository 尚未在 GitHub 建立"
    echo "2. 需要設定 Personal Access Token"
    echo ""
    echo "解決方法："
    echo "1. 確認 repository 已建立："
    echo "   https://github.com/$GITHUB_USERNAME/$REPO_NAME"
    echo ""
    echo "2. 產生 Personal Access Token："
    echo "   https://github.com/settings/tokens"
    echo "   - Generate new token (classic)"
    echo "   - 勾選 'repo' 權限"
    echo "   - 產生並複製 token"
    echo ""
    echo "3. 重新推送："
    echo "   git push -u origin main"
    echo "   帳號：$GITHUB_USERNAME"
    echo "   密碼：使用剛才複製的 token（不是您的 GitHub 密碼）"
    echo ""
fi
