# Management Dashboard GAS

経営管理用のLINE通知システム（Google Apps Script）

## 機能

- 📊 日次PLレポートの自動送信
- 📅 期間指定でのPLレポート送信
- 💬 LINEグループへの自動通知
- 📈 売上・粗利・営業利益の可視化

## ブランチ戦略

このプロジェクトは以下のGit-flowブランチ戦略を採用しています：

### ブランチ構成

```
main (本番環境・安定版)
  ├── develop (開発用メインブランチ)
  │   ├── feature/機能名 (新機能開発用)
  │   ├── bugfix/バグ名 (バグ修正用)
  │   └── refactor/対象名 (リファクタリング用)
  └── hotfix/緊急修正名 (本番環境の緊急修正用)
```

### ブランチ運用ルール

#### 1. **main ブランチ**
- 常に動作保証された安定版を維持
- 直接の編集は**禁止**
- developブランチから定期的にマージ
- タグでバージョン管理（v1.0.0形式）

#### 2. **develop ブランチ**
- 開発のメインブランチ
- 各機能ブランチをマージして統合テスト
- mainへのマージ前の最終確認場所

#### 3. **feature ブランチ**
- 新機能開発時に必ず作成
- 命名規則: `feature/機能名-日付`
- 例: `feature/自動レポート追加-20240826`
- developから分岐し、developへマージ

#### 4. **bugfix ブランチ**
- バグ修正時に作成
- 命名規則: `bugfix/不具合内容-日付`
- 例: `bugfix/データ取得エラー-20240826`
- developから分岐し、developへマージ

#### 5. **hotfix ブランチ**
- 本番環境の緊急修正時に使用
- 命名規則: `hotfix/修正内容-日付`
- mainから分岐し、mainとdevelopへマージ

### 作業フロー

#### 新機能開発
```bash
git checkout develop
git checkout -b feature/機能名-日付
# 開発作業
git add -A
git commit -m "feat: 機能説明"
git checkout develop
git merge feature/機能名-日付
git push origin develop
```

#### バグ修正
```bash
git checkout develop
git checkout -b bugfix/バグ名-日付
# 修正作業
git add -A
git commit -m "fix: 修正内容"
git checkout develop
git merge bugfix/バグ名-日付
git push origin develop
```

#### 緊急修正
```bash
git checkout main
git checkout -b hotfix/修正名-日付
# 修正作業
git add -A
git commit -m "fix: 緊急修正内容"
git checkout main
git merge hotfix/修正名-日付
git push origin main
git checkout develop
git merge hotfix/修正名-日付
git push origin develop
```

#### 本番リリース
```bash
git checkout main
git merge develop
git tag v1.x.x
git push origin main --tags
clasp push  # GASへデプロイ
```

### コミットメッセージ規則

```
<type>: <subject>

<body>

🤖 Generated with Claude Code
Co-Authored-By: Claude <noreply@anthropic.com>
```

**タイプ一覧:**
- `feat`: 新機能
- `fix`: バグ修正
- `docs`: ドキュメント変更
- `style`: コード整形（動作に影響なし）
- `refactor`: リファクタリング
- `test`: テスト追加・修正
- `chore`: ビルド処理やツール変更

### バージョン管理

セマンティックバージョニング（v1.2.3形式）を採用：
- **メジャー（1.x.x）**: 大規模な機能追加や仕様変更
- **マイナー（x.2.x）**: 機能追加や改善
- **パッチ（x.x.3）**: バグ修正

### 重要な注意事項

1. **必ず動作確認してからマージ** - 特にmainへのマージは慎重に
2. **こまめなコミット** - 作業単位で細かくコミット
3. **ブランチの削除** - マージ済みのfeature/bugfixブランチは削除
4. **定期的なバックアップ** - リモートリポジトリへのプッシュ
5. **mainブランチの保護** - 直接pushは禁止、必ずPR経由でマージ

## セットアップ

### 1. GASプロジェクトのクローン
```bash
clasp clone [Script ID]
```

### 2. LINE Messaging APIの設定
1. LINE Developersコンソールでチャンネルを作成
2. チャンネルアクセストークンを取得
3. GASの初期設定メニューから設定を保存

### 3. スプレッドシートの準備
- 「全社PLシート」という名前のシートを作成
- A列: 日付
- C列: 単日売上
- D列: 当月累計売上
- G列: 単日粗利
- H列: 当月累計粗利
- K列: 単日営業利益
- L列: 当月累計営業利益

### 4. 自動実行の設定
GASのトリガー設定で`sendDailyPLReportAuto`関数を毎日定時実行に設定

## 使用方法

### 手動送信
スプレッドシートのメニューから「📱 LINE通知」を選択：
- 昨日のPL情報を送信
- 指定期間のPL情報を送信

### 設定確認
メニューから「🔍 設定確認」で現在の設定を確認可能

## 技術スタック

- Google Apps Script
- LINE Messaging API
- Google Sheets API
- clasp (Command Line Apps Script Projects)

## ライセンス

Private Project - All Rights Reserved

## サポート

問題が発生した場合は、GitHubのIssuesでお知らせください。