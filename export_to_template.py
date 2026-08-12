#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
時間割スケジューラのバックアップJSON → 全体時間割テンプレ(xlsx)へ書き出し
使い方: python3 export_to_template.py <テンプレ.xlsx> <バックアップ.json> <出力.xlsx> [履歴ラベル]
（履歴ラベル省略時は最新の履歴を使用）
"""
import sys, json, io, datetime
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

BANDS=[4,41,78,115,152]          # 各週バンドの先頭行（日付行）
DAY_W=20                          # 1日ブロックの幅（デスクA-D×4列＋自習室4列）
GRID_C0=3                         # グリッド開始列(C)
GRID_LAST=142                     # グリッド最終列(EL)
PAL_GROUPS_L=[144,151,158]        # パレット左半分(EN,EU,FB)
PAL_GROUPS_R=[165,172,179]        # パレット右半分(FI,FP,FW)
PAL_AREAS=[9,46,83]               # パレットのブロック開始行（各8ブロック）

def set_val(ws,r,c,v):
    cell=ws.cell(r,c)
    if isinstance(cell,MergedCell): return
    cell.value=v

# 生徒名簿CSV（塾システム書き出し）。略字が被る生徒の「下の名前の頭文字」を引くのに使う
import os as _os
ROSTER_CSV=next((_p for _p in [_os.path.join(_os.path.dirname(_os.path.abspath(__file__)),"生徒情報.csv"),"/Users/rootnishino/Downloads/生徒情報.csv"] if _os.path.exists(_p)), "/Users/rootnishino/Downloads/生徒情報.csv")
# CSV未収録ぶんの名。実名なのでリポジトリ（公開）には置かず
# 生徒名補完.json（.gitignoreの*.json対象）に分離。無い場合はCSVだけで動く。
MANUAL_GIVEN_PATH=_os.path.join(_os.path.dirname(_os.path.abspath(__file__)),"生徒名補完.json")
try: MANUAL_GIVEN=json.load(open(MANUAL_GIVEN_PATH,encoding="utf-8"))
except Exception: MANUAL_GIVEN={}

def load_given(app_names):
    """アプリ表示名 → 下の名前（名）の対応表。CSVの氏名から引く。"""
    import csv, os
    rows=[]
    try:
        if os.path.exists(ROSTER_CSV):
            raw=open(ROSTER_CSV,"rb").read().decode("utf-8-sig",errors="replace")
            rows=[(r["氏"].strip(), r["名"].strip()) for r in csv.DictReader(io.StringIO(raw)) if r.get("氏")]
    except Exception:
        rows=[]  # CSVが読めなくても略字は abbr / 氏名頭文字で成立する
    out={}
    for nm in app_names:
        hit=[c for c in rows if c[0]==nm]                                   # 苗字一致
        if len(hit)!=1:
            hit=[c for c in rows if nm.startswith(c[0]) and c[1].startswith(nm[len(c[0]):])]  # 苗字+名頭
        if len(hit)==1: out[nm]=hit[0][1]
        elif nm in MANUAL_GIVEN: out[nm]=MANUAL_GIVEN[nm]
    return out

def abbrev_cells(student_objs, given):
    """メンバーの略字リストを作る。
    ①アプリで登録した表示略字(abbr)があれば最優先。
    ②未登録で、苗字頭文字が集合内で重複する生徒は下の名前の頭文字を使う。"""
    base=[ (o.get("abbr") or "").strip()[:1] or (o.get("name") or "?")[0] for o in student_objs ]
    cnt={}
    for c in base: cnt[c]=cnt.get(c,0)+1
    chars=[]
    for o,c in zip(student_objs,base):
        if not (o.get("abbr") or "").strip() and cnt[c]>1:  # abbr未登録かつ苗字頭が被る → 名の頭文字へ
            g=given.get(o.get("name",""),"")
            chars.append(g[0] if g else c)
        else:
            chars.append(c)
    return chars

def main():
    tpl,bkp,out=sys.argv[1],sys.argv[2],sys.argv[3]
    want_label=sys.argv[4] if len(sys.argv)>4 else None
    data=json.load(open(bkp))
    app=data["app"] if "app" in data else data
    season=next(s for s in app["seasons"] if s["id"]==app["currentId"])
    st=season["data"]
    hist=st.get("history",[])
    if not hist: sys.exit("履歴がありません。ツールで生成→★履歴に保存 してからバックアップしてください。")
    entry=next((h for h in hist if h["label"]==want_label), hist[0])
    teachers=st["teachers"]; units={u["id"]:u for u in st["units"]}
    students={s["id"]:s for s in st["students"]}
    given=load_given([s["name"] for s in st["students"]])
    komas=st["komas"]; kidx={k["id"]:i for i,k in enumerate(komas)}
    tidx={t["id"]:i for i,t in enumerate(teachers)}
    start=datetime.datetime.strptime(st["startDate"],"%Y-%m-%d")
    start-=datetime.timedelta(days=start.weekday()+1 if start.weekday()!=6 else 0)  # 直前の日曜へ
    open_dows=set(st["openDows"])

    wb=load_workbook(tpl)
    ws=wb["全体時間割"]

    # タイトル: 年だけ更新
    import re
    t=ws["AE2"].value
    if isinstance(t,str): ws["AE2"]=re.sub(r"\d{4}",str(start.year),t,count=1)

    # ---- グリッドのクリア（コマ行のみ。ヘッダ行・B列は残す） ----
    for bt in BANDS:
        for r in range(bt+5, bt+5+4*8):
            for c in range(GRID_C0, GRID_LAST+1):
                set_val(ws,r,c,None)

    # ---- 日付・講師名・コマ時刻 ----
    for w,bt in enumerate(BANDS):
        for d in range(7):
            c0=GRID_C0+d*DAY_W
            set_val(ws,bt,c0, start+datetime.timedelta(days=w*7+d))
            for j,tc in enumerate(teachers[:5]):
                set_val(ws,bt+3,c0+j*4, tc["name"] if d in open_dows else None)
            # 5列目は旧「自習室」。5人目の講師がいればデスクEとして扱う
            if len(teachers)>=5: set_val(ws,bt+2,c0+16,"E")
        for i,k in enumerate(komas):
            tm=(k.get("time") or "").replace(":","：").replace("-","\n～\n")
            if tm: set_val(ws,bt+5+4*i,2, tm)

    # ---- 配置の書き込み（コード＋メンバー略字。色は条件付き書式が自動適用） ----
    placed={}
    for s in entry["sessions"]:
        u=units.get(s["unitId"]);
        if not u: continue
        w,d,k=s["w"],s["d"],s["k"]
        if w>=len(BANDS) or k not in kidx: continue
        bt=BANDS[w]; r=bt+5+4*kidx[k]
        c0=GRID_C0+d*DAY_W+tidx.get(u["teacherId"],0)*4
        set_val(ws,r,c0,u["subject"])
        mem=[m for m in u["members"]]+[m for m in s.get("extra",[])]
        chars=abbrev_cells([students.get(mid,{"name":"?"}) for mid in mem], given)
        cells=chars[:3]+(["".join(chars[3:])] if len(chars)>3 else [])  # 5人以上は4席目に連記
        for i,ch in enumerate(cells):
            set_val(ws,r+3,c0+i,ch)
        placed[u["id"]]=placed.get(u["id"],0)+1

    # ---- EN2: 延べ人数（生徒の重複・かけもち受講も含む）----
    # グリッドの生徒名の行だけを対象に文字数を数える。1セルに2文字あれば2人（5人目以降の連記に対応）。
    # 行の判定: 週バンド(37行周期・先頭4行目)内のオフセットが8以上かつ4の倍数＝生徒名の行。
    ws["EN2"]=("=SUMPRODUCT((MOD(ROW($C$9:$EL$188)-4,37)>=8)"
               "*(MOD(MOD(ROW($C$9:$EL$188)-4,37),4)=0)*LEN($C$9:$EL$188))")

    # ---- パレット再構築 ----
    # クリア（EN2/EO2 の総合計式がある2行目は残す）
    for r in range(3,189):
        for c in range(144,187):
            set_val(ws,r,c,None)
    from openpyxl.utils import get_column_letter
    # 講師の割当先: [ (ヘッダ行, 列グループ群) ] 左上=講師1, 左下=講師2, 右上=講師3, 右下=講師4
    slots=[(BANDS[0],PAL_GROUPS_L),(BANDS[1],PAL_GROUPS_L),(BANDS[0],PAL_GROUPS_R),
           (BANDS[1],PAL_GROUPS_R),(BANDS[2],PAL_GROUPS_L),(BANDS[2],PAL_GROUPS_R)]
    area_of={BANDS[0]:PAL_AREAS[0],BANDS[1]:PAL_AREAS[1],BANDS[2]:PAL_AREAS[2]}
    for ti,tc in enumerate(teachers[:6]):
        hdr_row,groups=slots[ti]
        set_val(ws,hdr_row,groups[0],tc["name"])
        my_units=[u for u in st["units"] if u["teacherId"]==tc["id"]]
        for n,u in enumerate(my_units):
            g=groups[(n//8)%3]; r=area_of[hdr_row]+4*(n%8)
            set_val(ws,r,g,u["subject"])
            code_ref=f"${get_column_letter(g)}${r}"
            names=[students.get(mid,{}).get("name","?") for mid in u["members"]]
            rows=names[:3]+(["・".join(names[3:])] if len(names)>3 else [])  # 5人以上は4行目に連記
            for i,nm in enumerate(rows):
                set_val(ws,r+i,g+4,nm)
                set_val(ws,r+i,g+5,u.get("komaMin",u.get("komaCount",8)))  # 回数は従来どおり生徒ごとの行に
            # 検算(COUNTIF)はユニットにつき1つ。各行に書くとEO2の合計が人数分ふくらむため
            set_val(ws,r,g+6,f"=COUNTIF($C$9:$EL$188,{code_ref})")
            chars=abbrev_cells([students.get(mid,{"name":"?"}) for mid in u["members"]], given)
            cells=chars[:3]+(["".join(chars[3:])] if len(chars)>3 else [])
            for i,ch in enumerate(cells):
                set_val(ws,r+3,g+i,ch)
    wb.save(out)
    restore_styles_from_template(tpl, out)  # openpyxlが壊す罫線等を雛形の書式で丸ごと復元（値は出力のまま）
    print(json.dumps({"placedUnits":{units[k]["subject"]:v for k,v in placed.items()},
                      "sessions":len(entry["sessions"]),"label":entry["label"]},ensure_ascii=False))

# 日と日の境界罫線（medium）: openpyxlは結合セルの従属セル書式を保存時に落とすため、
# 保存後にxlsxのXMLを直接書き換えて両側をmediumに揃える
def restore_styles_from_template(tpl_path, out_path):
    """openpyxlは結合セルの従属セル書式（罫線など）を保存時に落とす。
    全体時間割シートには値しか書き込まないので、出力の styles.xml を雛形のものに差し替え、
    各セルの s(スタイル)属性を雛形から復元すれば、値は出力・見た目は雛形と完全一致になる。"""
    import zipfile, re, shutil, os, tempfile
    SHEET="xl/worksheets/sheet1.xml"   # 全体時間割＝先頭シート
    with zipfile.ZipFile(tpl_path) as z:
        tpl_styles=z.read("xl/styles.xml"); tpl_sheet=z.read(SHEET).decode("utf-8")
    with zipfile.ZipFile(out_path) as z:
        names=z.namelist(); data={n:z.read(n) for n in names}
    out_sheet=data[SHEET].decode("utf-8")
    cell_s={}
    for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)/?>', tpl_sheet):
        sm=re.search(r'\bs="(\d+)"', m.group(2)); cell_s[m.group(1)]=sm.group(1) if sm else "0"
    def fix(mo):
        ref, attrs, tail = mo.group(1), mo.group(2), mo.group(3)
        if ref not in cell_s: return mo.group(0)
        attrs=re.sub(r'\s*\bs="\d+"','',attrs)
        return f'<c r="{ref}" s="{cell_s[ref]}"{attrs}{tail}'
    out_sheet=re.sub(r'<c r="([A-Z]+\d+)"([^>]*?)(/?>)', fix, out_sheet)
    data["xl/styles.xml"]=tpl_styles; data[SHEET]=out_sheet.encode("utf-8")
    fd,tmp=tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
    with zipfile.ZipFile(tmp,"w",zipfile.ZIP_DEFLATED) as z:
        for n in names: z.writestr(n, data[n])
    shutil.move(tmp, out_path)

if __name__=="__main__": main()
