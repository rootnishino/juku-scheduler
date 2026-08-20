#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
バックアップJSONの整理。古いものを _アーカイブ/YYYY-MM/ に移すだけで、削除はしない。

  python3 tidy_backups.py [対象フォルダ]           # 下見（何も動かさない）
  python3 tidy_backups.py [対象フォルダ] --apply   # 実行

アプリの復元ダイアログはサブフォルダを列挙しないため（index.html: h.kind==="file"）、
移動するだけで一覧から消える。Comiru請求 の glob も非再帰なので同様。

保護するもの:
  - 時間割データ_自動バックアップ.json   … アプリが上書きし続ける現役ファイル
  - 自動生成パターン以外の名前のファイル … 凍結データ・2学期データなど手で付けたもの
    （移すのは 時間割データ_YYYY-MM-DD_HHMM.json 形式だけ）
  - 直近 KEEP_DAYS 日以内のファイル
  - それ以前は「1日につき最新の1個」だけ残す
"""
import sys, os, re, json, shutil, datetime, unicodedata

AUTO      = "時間割データ_自動バックアップ.json"
# アプリが自動で付ける名前だけを移動対象にする（exportBackup / downloadBackup 由来）。
#   時間割データ_2026-07-21_1336.json
#   時間割データ_バックアップ_2026-07-20_1300.json
GENERATED = re.compile(r"^時間割データ_(バックアップ_)?\d{4}-\d{2}-\d{2}_\d{4}\.json$")
ARCHIVE   = "_アーカイブ"
KEEP_DAYS = 7          # この日数以内は全部残す
KEEP_LAST = 3          # 加えて、最新のこれだけは無条件で残す

def norm(s):
    """macOSのファイル名はNFDで並ぶ。比較は必ず正規化してから行う。"""
    return unicodedata.normalize("NFC", s)

def scan(root):
    out = []
    for name in os.listdir(root):          # glob は NFC リテラルと一致しないので使わない
        p = os.path.join(root, name)
        if not os.path.isfile(p) or not name.lower().endswith(".json"):
            continue
        out.append((name, p, os.path.getmtime(p), os.path.getsize(p)))
    return sorted(out, key=lambda r: r[2], reverse=True)

def sessions_of(path):
    """確定履歴のセッション数。読めなければ None。"""
    try:
        a = json.load(open(path, encoding="utf-8"))
        a = a.get("app", a)
        ses = a.get("seasons") or []
        hist = (ses[0].get("data", {}).get("history") or []) if ses else []
        kaku = [h for h in hist if h.get("label") == "確定"]
        return len(kaku[-1].get("sessions", [])) if kaku else 0
    except Exception:
        return None

def classify(files):
    """(残す, 移す) に振り分け、理由を添えて返す。"""
    now = datetime.datetime.now().timestamp()
    cutoff = now - KEEP_DAYS * 86400
    keep, move, seen_day = [], [], set()
    for i, (name, path, mtime, size) in enumerate(files):
        n = norm(name)
        if n == AUTO:
            keep.append((name, path, mtime, size, "現役の自動バックアップ")); continue
        if not GENERATED.match(n):
            keep.append((name, path, mtime, size, "手で名前を付けたファイル")); continue
        if i < KEEP_LAST:
            keep.append((name, path, mtime, size, f"最新{KEEP_LAST}件")); continue
        if mtime >= cutoff:
            keep.append((name, path, mtime, size, f"直近{KEEP_DAYS}日以内")); continue
        day = datetime.date.fromtimestamp(mtime).isoformat()
        if day not in seen_day:
            seen_day.add(day)
            keep.append((name, path, mtime, size, f"{day} の最新")); continue
        move.append((name, path, mtime, size, f"{day} の重複分"))
    return keep, move

def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    apply_ = "--apply" in sys.argv
    root = os.path.abspath(args[0]) if args else os.getcwd()
    if not os.path.isdir(root):
        sys.exit(f"フォルダがありません: {root}")

    files = scan(root)
    keep, move = classify(files)

    print(f"対象: {root}")
    print(f"JSON {len(files)}個 → 残す {len(keep)} / 移す {len(move)}\n")

    print("--- 残す ---")
    for name, path, mtime, size, why in keep:
        n = sessions_of(path)
        tag = f"確定{n}コマ" if n else "確定なし"
        print(f"  {name}  [{tag}]  ({why})")

    if not move:
        print("\n移動対象はありません。")
        return

    print("\n--- 移す ---")
    total = 0
    for name, path, mtime, size, why in move:
        n = sessions_of(path)
        tag = f"確定{n}" if n else "確定なし"
        sub = datetime.date.fromtimestamp(mtime).strftime("%Y-%m")
        print(f"  {name}  [{tag}]  → {ARCHIVE}/{sub}/  ({why})")
        total += size

    print(f"\n移動サイズ合計: {total//1024} KB")
    if not apply_:
        print("\n下見のみ。実行するには --apply を付けてください。")
        return

    for name, path, mtime, size, why in move:
        sub = datetime.date.fromtimestamp(mtime).strftime("%Y-%m")
        dest_dir = os.path.join(root, ARCHIVE, sub)
        os.makedirs(dest_dir, exist_ok=True)
        dest = os.path.join(dest_dir, name)
        if os.path.exists(dest):
            print(f"  スキップ {name}: 移動先に同名あり")
            continue
        shutil.move(path, dest)
    print(f"\n{len(move)}個を {ARCHIVE}/ へ移しました。削除はしていません。")

if __name__ == "__main__":
    main()
