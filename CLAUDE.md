# juku-scheduler

塾の季節講習（春期・夏期・冬期）の時間割を自動生成するローカルHTMLアプリ。git管理。

## 主要ファイル
- `index.html` — アプリ本体。ブラウザで開いて使う。
- `時間割データ*.json` — 保存データ。`app.seasons[0].data.history[]` に確定履歴（sessions）が入る。**label="確定" の最新版が正**。
- `export_personal.py` / `export_classroom.py` / `export_to_template.py` — 配布用Excel書き出し。

## 他プロジェクトとの連携
Comiru請求（~/Projects/Comiru請求）が講習費の算定にこのJSONを直接読む。生徒別コマ数は sessions × units.members で数える。Excel書き出しは古くなりうるのでJSONが正データ。
