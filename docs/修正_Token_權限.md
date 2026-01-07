# 修正 Token 權限

## ❌ 問題
Token 缺少 `workflow` 權限，無法推送 `.github/workflows/build.yml`

## ✅ 解決方案

### 方法 1：重新產生 Token（增加 workflow 權限）⭐ 推薦

1. **前往 GitHub Token 頁面**：
   ```
   https://github.com/settings/tokens
   ```

2. **刪除剛才的 Token**（或保留，產生新的）

3. **點擊 "Generate new token (classic)"**

4. **填寫資訊**：
   - Note: `text-to-ppt-converter`
   - Expiration: `90 days` 或 `No expiration`
   - **勾選權限**：
     - ✅ `repo`（整個區塊）
     - ✅ `workflow` ⭐ **重要！這次要勾選這個**

5. **Generate token → 複製新的 token**

6. **重新推送**：
   ```bash
   git push -u origin main
   ```
   - Username: `Junyi1227`
   - Password: 貼上新的 token

---

### 方法 2：暫時移除 GitHub Actions 檔案（不推薦）

如果您想先推送程式碼，之後再加入 GitHub Actions：

```bash
# 移除 GitHub Actions 設定
git rm -r .github/
git commit -m "暫時移除 GitHub Actions"
git push -u origin main

# 之後再加回來（使用有 workflow 權限的 token）
git checkout HEAD~1 -- .github/
git add .github/
git commit -m "加入 GitHub Actions"
git push
```

**不推薦這個方法**，因為我們的目標就是要用 GitHub Actions 打包！

---

### 方法 3：使用 SSH（一勞永逸）⭐ 長期最佳方案

設定 SSH 後就不需要 token 了：

```bash
# 1. 產生 SSH 金鑰
ssh-keygen -t ed25519 -C "trance1227@gmail.com"
# 全部按 Enter（使用預設值）

# 2. 啟動 ssh-agent 並加入金鑰
eval "$(ssh-agent -s)"
ssh-add ~/.ssh/id_ed25519

# 3. 複製公鑰
cat ~/.ssh/id_ed25519.pub
# 複製全部輸出

# 4. 前往 GitHub 加入 SSH 金鑰
# https://github.com/settings/ssh/new
# Title: Mac
# Key: 貼上剛才複製的內容
# 點擊 Add SSH key

# 5. 修改 remote URL
git remote set-url origin git@github.com:Junyi1227/text-to-ppt-converter.git

# 6. 推送
git push -u origin main
```

---

## 🎯 推薦做法

**選擇以下任一方式**：

### 快速方案（5 分鐘）
→ **方法 1**：重新產生 Token，這次勾選 `workflow` 權限

### 長期方案（10 分鐘）
→ **方法 3**：設定 SSH，以後都不用輸入密碼

---

## 📝 立即行動

我建議：
1. 先用**方法 1**（重新產生 Token + 勾選 workflow）快速完成推送
2. 之後有時間再設定 SSH（方法 3）

**現在請前往產生新的 Token（記得勾選 workflow）！** 🚀
