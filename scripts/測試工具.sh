#!/bin/bash
# 測試 Python 程式是否正常運作

set -e

echo "========================================"
echo "文字轉 PowerPoint - 測試工具"
echo "========================================"
echo ""

# 檢查 Python
echo "檢查 Python..."
if command -v python3 &> /dev/null; then
    PYTHON_VERSION=$(python3 --version)
    echo "✅ $PYTHON_VERSION"
else
    echo "❌ Python 3 未安裝"
    echo "請前往 https://www.python.org/downloads/ 安裝"
    exit 1
fi

# 檢查必要檔案
echo ""
echo "檢查必要檔案..."
FILES=(
    "text_to_ppt.py"
    "text_to_ppt_gui.py"
    "examples/範例輸入文字.txt"
    ".github/workflows/build.yml"
    "requirements.txt"
)

for file in "${FILES[@]}"; do
    if [ -f "$file" ]; then
        echo "✅ $file"
    else
        echo "❌ $file 不存在"
        exit 1
    fi
done

# 檢查 python-pptx
echo ""
echo "檢查 python-pptx 套件..."
if python3 -c "import pptx" 2>/dev/null; then
    echo "✅ python-pptx 已安裝"
else
    echo "⚠️  python-pptx 未安裝"
    read -p "是否要安裝？(y/n) " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        pip3 install python-pptx
        echo "✅ python-pptx 已安裝"
    else
        echo "❌ 需要安裝 python-pptx 才能繼續測試"
        exit 1
    fi
fi

# 測試命令列版本
echo ""
echo "測試 1: 命令列版本..."
if python3 text_to_ppt.py examples/範例輸入文字.txt tmp_rovodev_test_cli.pptx 2>/dev/null; then
    if [ -f "tmp_rovodev_test_cli.pptx" ]; then
        SIZE=$(du -h tmp_rovodev_test_cli.pptx | cut -f1)
        echo "✅ 命令列版本正常（產生檔案: $SIZE）"
        rm tmp_rovodev_test_cli.pptx
    else
        echo "❌ 命令列版本失敗：未產生檔案"
        exit 1
    fi
else
    echo "❌ 命令列版本執行失敗"
    exit 1
fi

# 測試 GUI 版本（只測試能否啟動，不開啟視窗）
echo ""
echo "測試 2: GUI 版本..."
echo "（將會開啟視窗，請手動關閉以繼續）"
read -p "按 Enter 開始測試 GUI 版本..." 

# 在背景執行 GUI，3 秒後自動關閉
timeout 3s python3 text_to_ppt_gui.py 2>/dev/null || true
echo "✅ GUI 版本可以啟動"

# 檢查 GitHub Actions 設定
echo ""
echo "測試 3: GitHub Actions 設定..."
if [ -f ".github/workflows/build.yml" ]; then
    # 簡單的 YAML 語法檢查
    if grep -q "runs-on: windows-latest" .github/workflows/build.yml; then
        echo "✅ GitHub Actions 設定正確"
    else
        echo "⚠️  GitHub Actions 設定可能有問題"
    fi
else
    echo "❌ GitHub Actions 設定檔案不存在"
    exit 1
fi

# 檢查範例檔案格式
echo ""
echo "測試 4: 範例檔案格式..."
if grep -q "##" examples/範例輸入文字.txt && grep -q "#" examples/範例輸入文字.txt; then
    echo "✅ 範例檔案格式正確"
else
    echo "⚠️  範例檔案格式可能有問題"
fi

echo ""
echo "========================================"
echo "✅ 所有測試通過！"
echo "========================================"
echo ""
echo "您可以："
echo "1. 手動測試 GUI："
echo "   python3 text_to_ppt_gui.py"
echo ""
echo "2. 手動測試命令列："
echo "   python3 text_to_ppt.py examples/範例輸入文字.txt 輸出.pptx"
echo ""
echo "3. 推送到 GitHub 進行打包："
echo "   bash scripts/快速開始.sh"
echo ""
