# 他プロジェクトで使うときの手順

この `scripts` フォルダを**他の GAS（clasp）プロジェクト**にもコピーして使えます。

---

## 1. コピー

- 使いたいプロジェクトの**ルート**に、この `scripts` フォルダをそのままコピーする。
- プロジェクトルート直下に `scripts` がある形にする。

```
プロジェクトルート/
  ├── .clasp.json
  ├── appsscript.json
  ├── （その他 .gs / .js）
  └── scripts/
        ├── up.ps1
        ├── down.ps1
        ├── up.bat
        ├── down.bat
        └── README.md
```

---

## 2. そのプロジェクトで clasp を用意

- そのプロジェクトで `clasp` を使う設定になっていること。
- 未設定なら、プロジェクトルートで:
  - `clasp login`
  - `clasp clone <スクリプトID>` または既存の `appsscript.json` がある状態で `clasp pull` などで `.clasp.json` を用意。

---

## 3. GAS に送る方（up）でデプロイを有効にする

- **GAS に送るときにデプロイまで行いたい**場合は、`scripts/up.ps1` を開き、先頭付近の次の行を書き換える。

```powershell
# 変更前（デプロイなし）
$EnableDeploy = $false

# 変更後（デプロイあり）
$EnableDeploy = $true
```

- `$true` にすると、`up.ps1` 実行時に **clasp push のあと clasp deploy** まで実行されます。

---

## 4. 実行方法

- 必ず**プロジェクトルート**で実行する（`scripts` の親フォルダで）。

**GAS に送る（ローカル → GitHub → GAS、デプロイありの場合はここでデプロイまで）:**

```powershell
.\scripts\up.ps1
```

**GAS から取り込む（GAS → ローカル → GitHub）:**

```powershell
.\scripts\down.ps1
```

- `.bat` をダブルクリックしても同じ動作です。

---

## 5. 前提

- **PowerShell** が使えること（Windows は標準で利用可）。
- **Git** がインストールされ、そのプロジェクトで `git init` 済みでリモート（GitHub 等）を設定済みのこと。
- **clasp** がインストールされ、`clasp login` 済みであること。
- プロジェクトルートに `.clasp.json` があり、正しい `scriptId` が書かれていること。

---

## 6. まとめ（他プロジェクト用）

| やりたいこと           | やること |
|------------------------|----------|
| スクリプトをコピー     | この `scripts` フォルダをプロジェクトルートにコピー |
| GAS に送る＋デプロイ   | `scripts/up.ps1` の `$EnableDeploy = $true` に変更 |
| GAS に送る             | プロジェクトルートで `.\scripts\up.ps1` |
| GAS から取り込む       | プロジェクトルートで `.\scripts\down.ps1` |
