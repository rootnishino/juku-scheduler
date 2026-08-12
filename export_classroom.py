#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
教室別時間割（1日1枚）を開校日数分生成
使い方: python3 export_classroom.py <テンプレ.xlsx> <バックアップ.json> <出力.xlsx> [履歴ラベル]
テンプレの「教室別時間割(月)」をマスターに複製（去年の完成形＝書式が最も整っているため）。
科目セルの表記は暫定ロジック（通し番号・学年接頭辞を除いた科目名の頭文字）。※後で決め直す前提
"""
import sys, json, re, datetime
from collections import Counter
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

SURNAME_LEN={"三ツ井":3,"篠":1}  # 名字が2文字でない生徒の例外（他は先頭2文字）
def _split(name):
    name=(name or "").strip()
    for k in SURNAME_LEN:
        if name.startswith(k): return k, name[len(k):]
    return name[:2], name[2:]
def _sei(s): return s.get("sei") or _split(s.get("name",""))[0]
def _mei(s): return s.get("mei") or _split(s.get("name",""))[1]

MASTER="教室別時間割(月)"
DOW="日月火水木金土"
KOMA_ROW=lambda ki: 4+4*ki          # 学年行。+1=学習者 +2=科目 +3=備考
DESK_COL=lambda j: 4+4*j            # デスクA..D の先頭列(D,H,L,P)
CLEAR_COLS=range(4,24)              # D..W（自習室含む）

def set_val(ws,r,c,v):
    cell=ws.cell(r,c)
    if not isinstance(cell,MergedCell): cell.value=v

def subject_display(code):
    s=re.sub(r"^\d+","",code)
    s=re.sub(r"^(中[1-3]|高[1-3]|高校|小[1-6]?)","",s)
    return s or code

def main():
    tpl,bkp,out=sys.argv[1],sys.argv[2],sys.argv[3]
    want=sys.argv[4] if len(sys.argv)>4 else None
    data=json.load(open(bkp)); app=data.get("app",data)
    season=next(s for s in app["seasons"] if s["id"]==app["currentId"])
    st=season["data"]
    entry=next((h for h in st["history"] if h["label"]==want), st["history"][0])
    units={u["id"]:u for u in st["units"]}
    students={s["id"]:s for s in st["students"]}
    # 教室別の名前表示: 基本は名字だけ。名字が他とかぶる場合のみ名前頭1文字を付ける
    seicnt=Counter(_sei(v) for v in students.values())
    def disp_name(x):
        s=students.get(x)
        if not s: return "?"
        sei=_sei(s); mei=_mei(s)
        return sei+(mei[:1] if seicnt.get(sei,0)>1 and mei else "")
    kidx={k["id"]:i for i,k in enumerate(st["komas"])}
    tidx={t["id"]:i for i,t in enumerate(st["teachers"])}
    start=datetime.datetime.strptime(st["startDate"],"%Y-%m-%d")
    start-=datetime.timedelta(days=(start.weekday()+1)%7)
    # 開校曜日＋日曜(0)。授業が無い日曜も白紙シートとして日付順に出力する
    all_dows=sorted(set(st["openDows"]) | {0})

    # (w,d) -> [session]
    days={}
    for s in entry["sessions"]:
        days.setdefault((s["w"],s["d"]),[]).append(s)

    wb=load_workbook(tpl)
    master=wb[MASTER]
    made=[]
    for w in range(st["weeks"]):
        for d in all_dows:
            date=start+datetime.timedelta(days=w*7+d)
            title=f"{date.month}.{date.day}({DOW[d]})"
            sh=wb.copy_worksheet(master)
            sh.title=title
            # ヘッダ: 年/月/日/曜・タイトル年
            set_val(sh,1,2,date.year); set_val(sh,1,3,date.month); set_val(sh,1,5,date.day)
            set_val(sh,1,6,f"日（{DOW[d]}）")
            t=sh.cell(1,20).value
            if isinstance(t,str): set_val(sh,1,20,re.sub(r"\d{4}",str(date.year),t,count=1))
            # 日付行(1行目)が高さ18/フォント18でギリギリ→見切れるので余裕を持たせ、年・月の列幅も広げる
            sh.row_dimensions[1].height=27
            sh.column_dimensions["B"].width=12    # 年
            sh.column_dimensions["C"].width=7     # 月
            # 講師名（5人全員。5人目=神崎は自習室ブロック=デスクEに入れる）
            for j,tc in enumerate(st["teachers"][:5]): set_val(sh,3,DESK_COL(j),tc["name"])
            set_val(sh,2,DESK_COL(4),"E")  # 5番目ブロックの見出しを「自習室」→「E」に
            # 中身クリア（学年/学習者/科目/備考 × D..W）
            for ki in range(8):
                r=KOMA_ROW(ki)
                for rr in range(r,r+4):
                    for c in CLEAR_COLS: set_val(sh,rr,c,None)
            # セッション書き込み
            cnt=0
            for s in days.get((w,d),[]):
                u=units.get(s["unitId"])
                if not u or s["k"] not in kidx: continue
                r=KOMA_ROW(kidx[s["k"]]); base=DESK_COL(tidx.get(u["teacherId"],0))
                mem=u["members"]+(s.get("extra") or [])
                subj=subject_display(u["subject"])[0]
                seats=mem[:3]+([mem[3:]] if len(mem)>3 else [])
                for i,m in enumerate(seats):
                    ids=m if isinstance(m,list) else [m]
                    names="・".join(disp_name(x) for x in ids)
                    grades="・".join(dict.fromkeys(students.get(x,{}).get("grade","") for x in ids))
                    set_val(sh,r,base+i,grades)
                    set_val(sh,r+1,base+i,names)
                    set_val(sh,r+2,base+i,subj)
                    cnt+=1
            made.append((title,cnt))
    for name in list(wb.sheetnames):
        if name not in [t for t,_ in made]: del wb[name]
    # シートを日付順に並べ替え（copy_worksheetは末尾追加なので既に日付順）
    wb.save(out)
    print(json.dumps({"days":len(made),"seats":sum(c for _,c in made),
                      "empty":[t for t,c in made if c==0],"label":entry["label"]},ensure_ascii=False))

if __name__=="__main__": main()
