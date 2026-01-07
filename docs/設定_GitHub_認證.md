# 設定 GitHub 認證 - 快速指南

## 🔑 方法：使用 Personal Access Token（最簡單）

### 步驟 1：產生 Token

1. **開啟瀏覽器，前往**：
   ```
   https://github.com/settings/tokens
   ```

2. **點擊右上角的 "Generate new token"**
   - 選擇 "Generate new token (classic)"

3. **填寫 Token 資訊**：
   - **Note**: `text-to-ppt-converter`
   - **Expiration**: 選擇 `90 days` 或 `No expiration`
   - **Select scopes**: 
     - ✅ 勾選 `repo` （勾選整個 repo 區塊）

4. **捲動到最下方，點擊綠色按鈕 "Generate token"**

5. **立即複製 token！**
   - ⚠️ Token 只會顯示一次
   - 看起來像：`ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx`
   - 複製並暫時貼到記事本

---

### 步驟 2：使用 Token 推送

**複製好 Token 後，回到終端機執行**：

```bash
git push -u origin main
```

**當要求輸入時**：
- **Username**: `Junyi1227`
- **Password**: 貼上剛才複製的 token（不是您的 GitHub 密碼！）

---

### 步驟 3：確認推送成功

推送成功後會看到類似訊息：
```
Enumerating objects: 25, done.
Counting objects: 100% (25/25), done.
...
To https://github.com/Junyi1227/text-to-ppt-converter.git
 * [new branch]      main -> main
```

---

## 🔄 之後如何避免每次都輸入？

### 方法 A：儲存認證（推薦）

```bash
# macOS 使用 Keychain 儲存
git config --global credential.helper osxkeychain
```

之後 Git 會自動記住您的 token。

### 方法 B：設定 SSH（一次設定，永久使用）

```bash
# 1. 產生 SSH 金鑰
ssh-keygen -t ed25519 -C "trance1227@gmail.com"
# 按 Enter 使用預設路徑
# 按 Enter 跳過密碼（或設定密碼）

# 2. 啟動 ssh-agent
eval "$(ssh-agent -s)"

# 3. 加入 SSH 金鑰
ssh-add ~/.ssh/id_ed25519

# 4. 複製公鑰
cat ~/.ssh/id_ed25519.pub
# 複製顯示的內容

# 5. 加入到 GitHub
# 前往：https://github.com/settings/ssh/new
# Title: Mac
# Key: 貼上剛才複製的公鑰
# 點擊 Add SSH key

# 6. 測試連線
ssh -T git@github.com
# 應該會看到：Hi Junyi1227! You've successfully authenticated...

# 7. 修改 remote URL 為 SSH
git remote set-url origin git@github.com:Junyi1227/text-to-ppt-converter.git
```

---

## 📝 現在請執行

1. ✅ 前往 GitHub 產生 Token
2. ✅ 複製 Token
3. ✅ 回到終端機執行：`git push -u origin main`
4. ✅ 輸入帳號和 Token

**完成後告訴我，我會幫您檢查 GitHub Actions 的狀態！**
