#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
個人時間割（配布用）を講師・生徒の人数分生成
使い方: python3 export_personal.py <テンプレ.xlsx> <バックアップ.json> <出力.xlsx> [履歴ラベル]
テンプレ内の「個人時間割フォーマット のコピー」（5週スロット版）をマスターに使用。
"""
import sys, json, re, io, datetime
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, Alignment

MASTER="夏期講習個人時間割フォーマット"
FONT_NAME="Zen Maru Gothic"
# セル幅いっぱいに揃えるための文字数→フォントサイズ（列幅13・3行結合セル基準）
def fit_size(text):
    n=len(text)
    return {1:26,2:24,3:18,4:14,5:11}.get(n,9)
# スロット: (日付行, 曜日開始列)。左列=第1,2,3週 / 右列=第4,5週（去年の春の並びと同じ列優先）
SLOTS=[(3,3),(32,3),(61,3),(3,12),(32,12)]
WEEK_LABEL_CELLS=[(3,1),(32,1),(61,1),(3,10),(32,10)]  # A3,A32,A61,J3,J32

def set_val(ws,r,c,v):
    cell=ws.cell(r,c)
    if not isinstance(cell,MergedCell): cell.value=v

def subject_display(code):
    s=re.sub(r"^\d+","",code)                      # 通し番号を除去
    s=re.sub(r"^(中[1-3]|高[1-3]|高校|小[1-6]?)","",s)  # 学年接頭辞を除去
    return s or code

# 生徒名簿CSV（塾システムからの書き出し）。アプリ上の略称→本名フルネーム を引くのに使う
import os as _os
ROSTER_CSV=next((_p for _p in [_os.path.join(_os.path.dirname(_os.path.abspath(__file__)),"生徒情報.csv"),"/Users/rootnishino/Downloads/生徒情報.csv"] if _os.path.exists(_p)), "/Users/rootnishino/Downloads/生徒情報.csv")
# CSVに載っていない生徒の本名（名のみ）。実名なのでリポジトリ（公開）には置かず
# 生徒名補完.json（.gitignoreの*.json対象）に分離。無い場合はCSVだけで動く。
MANUAL_NAMES_PATH=_os.path.join(_os.path.dirname(_os.path.abspath(__file__)),"生徒名補完.json")
try: MANUAL_NAMES=json.load(open(MANUAL_NAMES_PATH,encoding="utf-8"))
except Exception: MANUAL_NAMES={}

def load_fullnames(app_names):
    """アプリの表示名（苗字 or 苗字+名の頭文字）→ 本名フルネーム の対応表を作る"""
    import csv, os
    rows=[]
    if os.path.exists(ROSTER_CSV):
        raw=open(ROSTER_CSV,"rb").read().decode("utf-8-sig",errors="replace")
        rows=[(r["氏"].strip(), r["名"].strip()) for r in csv.DictReader(io.StringIO(raw)) if r.get("氏")]
    out={}; missing=[]
    for nm in app_names:
        hit=[c for c in rows if c[0]==nm]                       # 苗字がそのまま一致
        if len(hit)!=1:
            hit=[c for c in rows if nm.startswith(c[0]) and c[1].startswith(nm[len(c[0]):])]  # 苗字+名の頭文字
        if len(hit)==1: out[nm]=hit[0][0]+hit[0][1]
        elif nm in MANUAL_NAMES:                                 # CSV未収録は手入力ぶんで補う
            sur=nm[:-1] if nm[-1]==MANUAL_NAMES[nm][0] else nm   # 「名字+名の頭文字」→「名字」
            out[nm]=sur+MANUAL_NAMES[nm]
        else: missing.append(nm)
    return out, missing

def main():
    tpl,bkp,out=sys.argv[1],sys.argv[2],sys.argv[3]
    want=sys.argv[4] if len(sys.argv)>4 else None
    data=json.load(open(bkp)); app=data.get("app",data)
    season=next(s for s in app["seasons"] if s["id"]==app["currentId"])
    st=season["data"]
    entry=next((h for h in st["history"] if h["label"]==want), st["history"][0])
    units={u["id"]:u for u in st["units"]}
    students={s["id"]:s for s in st["students"]}
    komas=st["komas"]; kidx={k["id"]:i for i,k in enumerate(komas)}
    start=datetime.datetime.strptime(st["startDate"],"%Y-%m-%d")
    start-=datetime.timedelta(days=(start.weekday()+1)%7)  # 日曜起点

    # 人ごとのセッション: (w,dow,komaIdx) -> 表示文字列
    def person_cells(kind,pid):
        cells={}
        for s in entry["sessions"]:
            u=units.get(s["unitId"]);
            if not u or s["k"] not in kidx: continue
            mem=u["members"]+ (s.get("extra") or [])
            if kind=="t" and u["teacherId"]!=pid: continue
            if kind=="s" and pid not in mem: continue
            key=(s["w"],s["d"],kidx[s["k"]])
            if kind=="t":
                def sabbr(x):
                    o=students.get(x,{}); a=(o.get("abbr") or "").strip()
                    return a[:1] if a else (o.get("name") or "?")[0]
                cells[key]="".join(sabbr(m) for m in mem)
            else:
                cells[key]=subject_display(u["subject"])
        return cells

    wb=load_workbook(tpl)
    # 雛形の探し方: ①先頭シートが雛形ならそれ ②名前で検索 ③旧名
    if wb.worksheets and wb.worksheets[0].title==MASTER: master=wb.worksheets[0]
    elif MASTER in wb.sheetnames: master=wb[MASTER]
    else: master=wb["個人時間割フォーマット のコピー"]
    # 氏名欄がフルネーム運用。CSVが読めればフォールバックに使うが、基本は氏名をそのまま印刷
    try: fullname,missing=load_fullnames([s["name"] for s in st["students"]])
    except Exception: fullname,missing={},[]
    def full_of(s):
        return (s.get("fullname") or "").strip() or fullname.get(s["name"]) or s["name"]
    persons=[("t",t["id"],t["name"]+"先生") for t in st["teachers"]] \
           +[("s",s["id"],full_of(s)) for s in st["students"]]
    made=[]
    for kind,pid,label in persons:
        sh=wb.copy_worksheet(master)
        sh.title=re.sub(r"[\[\]:*?/\\]","",label)[:31]
        # タイトル・氏名
        t=sh.cell(1,1).value
        if isinstance(t,str): set_val(sh,1,1,re.sub(r"\d{4}",str(start.year),t,count=1))
        set_val(sh,1,7,label)
        # 週ラベル
        for i,(r,c) in enumerate(WEEK_LABEL_CELLS): set_val(sh,r,c,i+1)
        # 各スロット: 日付更新＋中身クリア＋書き込み
        cells=person_cells(kind,pid)
        for w,(r0,c0) in enumerate(SLOTS):
            for d in range(7):
                set_val(sh,r0,c0+d, start+datetime.timedelta(days=w*7+d))
                for ki in range(8):
                    r=r0+5+3*ki; v=cells.get((w,d,ki))
                    set_val(sh,r,c0+d, v)
                    cell=sh.cell(r,c0+d)
                    if v and not isinstance(cell,MergedCell):
                        cell.font=Font(name=FONT_NAME,size=fit_size(v))
                        cell.alignment=Alignment(horizontal="center",vertical="center")
        made.append((label,len(cells)))
    # 個人票＋雛形だけの配布用ブックにする（雛形は先頭に置き、次回はこれを雛形として使える）
    keep={re.sub(r"[\[\]:*?/\\]","",l)[:31] for l,_ in made} | {MASTER}
    for name in list(wb.sheetnames):
        if name not in keep: del wb[name]
    if master.title!=MASTER: master.title=MASTER
    wb.move_sheet(master, offset=-wb.index(master))  # 先頭へ
    wb.save(out)
    print(json.dumps({"nameNotFound":missing,"people":len(made),"cells":{l:n for l,n in made},"label":entry["label"]},ensure_ascii=False))

if __name__=="__main__": main()
