"""Build the downloadable MPA database workbooks (zh and en) from the site's data files."""
import sys, math
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.worksheet.table import Table, TableStyleInfo

LANG = sys.argv[1]
import os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + "/"
DD = ROOT + ("data/" if LANG == "zh" else "en/data/")
OUT = ROOT + f"download/Taiwan_MPA_database_{LANG}.xlsx"
Z = LANG == "zh"
t = lambda zh, en: zh if Z else en

rd = lambda f: pd.read_csv(DD + f)
dr, ef, bh, cl, pr = rd("mpa_directory.csv"), rd("mpa_effectiveness.csv"), rd("biohab.csv"), rd("cluster_mpa.csv"), rd("pressure_mpa.csv")
cx, bc, bt, off = rd("cross_mpa_slim.csv"), rd("bluecarbon.csv"), rd("benthic_mpa.csv"), rd("offshore_mpa.csv")
en_names = pd.read_csv(ROOT + "en/data/mpa_directory.csv").set_index("no")["mpa"]
zh_names = pd.read_csv(ROOT + "data/mpa_directory.csv").set_index("no")["mpa"]
for d in (ef, bh, cl, pr, cx, bc): d.set_index("no", inplace=True)
name2no = dr.set_index("mpa")["no"].to_dict()

FONT = "Arial"
HDR = PatternFill("solid", fgColor="13315C"); HF = Font(name=FONT, bold=True, color="FFFFFF", size=10)
GRP = {"basic": "DCE6F0", "prot": "E4EFE9", "score": "FBEBD9", "calc": "F6E3E1"}
thin = Side(style="thin", color="C9D3DB")
BODY = Font(name=FONT, size=10)

def val(v):
    if v is None: return None
    if isinstance(v, float) and math.isnan(v): return None
    if isinstance(v, str) and v.strip() in ("", "—"): return None
    return v

wb = Workbook()

# ---------------- README ----------------
ws0 = wb.active; ws0.title = t("說明", "README")
ws0.sheet_view.showGridLines = False
rows0 = [
    (t("臺灣海洋保護區資料庫", "Taiwan Marine Protected Areas database"), "title"),
    (t("71 處海洋保護區的基本資料、保護等級與成效指標分數（v4，2026-11-04）", "Basic data, protection levels and effectiveness scores for 71 MPAs (v4, 2026-11-04)"), "sub"),
    ("", None),
    (t("基準年（用於計算成立年數）", "Reference year (for years since establishment)"), "year"),
    ("", None),
    (t("工作表", "Sheets"), "h"),
    (t("保護區總表：每處一列，含基本資料、保護等級、四面向分數，以及以公式計算的綜合分、等級與最弱面向。", "MPAs: one row per MPA with basic data, protection level, the four dimension scores, and the composite score, grade and weakest dimension computed by formula."), "p"),
    (t("分區保護等級：各保護區分區與 MPA Guide 保護等級。", "Zones: zones of each MPA and their MPA Guide protection levels."), "p"),
    (t("指標原始資料：水質、海廢、執法、生物多樣性分項、eDNA、珊瑚礁體檢、藍碳與生態系類型。", "Raw indicators: water quality, debris, enforcement, biodiversity components, eDNA, reef checks, blue carbon and ecosystem types."), "p"),
    (t("人口與漁業壓力：鄰近鄉鎮人口、海廢中漁業來源占比、巡守投入。", "Pressure: nearby population, fishing share of debris, patrol effort."), "p"),
    (t("輔助調查：底棲群聚與外海船舶測站（海保署 110–112 年重要生態系調查）。", "Supplementary surveys: benthic communities and offshore ship stations (OCA important-ecosystem survey, 2021–2023)."), "p"),
    (t("欄位說明：各欄位的定義與來源。", "Field guide: definition and source of each field."), "p"),
    ("", None),
    (t("等級門檻", "Grade thresholds"), "h"),
    (t("綜合分 ≥70 良好、55–69 尚可、40–54 待加強、<40 不佳；只有兩個面向有資料者標「暫評」，少於兩個面向為「資料不足」。綜合分為有資料面向的平均，以「四捨六入五成雙」取整（與網站一致）。", "Composite ≥70 good, 55–69 fair, 40–54 needs improvement, <40 poor; sites with only two dimensions are marked provisional, sites with fewer than two are insufficient data. The composite is the mean of the dimensions with data, rounded half to even (as on the website)."), "p"),
    ("", None),
    (t("資料來源", "Sources"), "h"),
    (t("海洋委員會海洋保育署（海洋保護區資訊平台、海域水質監測、海岸淨灘監測、魚類環境 DNA、潛水與船舶調查、巡守與執法通報、2024 藍碳調查）；農業部漁業署保育區公告；台灣生物多樣性網絡 TBN；珊瑚礁體檢（中研院 2008、台灣環境資訊協會 2023、Reef Check Taiwan 2024–26）；MPATable1 保護區分類檢核表；維基百科鄉鎮人口列表。", "Ocean Conservation Administration (MPA platform, water-quality and beach-litter monitoring, fish eDNA, dive and ship surveys, patrol and enforcement reports, 2024 blue-carbon survey); Fisheries Agency notices; Taiwan Biodiversity Network (TBN); reef checks (Academia Sinica 2008, TEIA 2023, Reef Check Taiwan 2024–26); MPATable1 classification checklist; Wikipedia township population list."), "p"),
    ("", None),
    (t("注意", "Notes"), "h"),
    (t("本資料庫為依公開資料所做的非官方評估，用於找出管理優先順序，不等同主管機關的官方成效分級。代表點座標依 eDNA 採樣站、公告座標或地名推定，僅供示意，不是正式範圍線。空白表示無資料。", "This is a non-official assessment based on public data, intended to identify management priorities; it is not the authorities' official grading. Representative coordinates come from eDNA stations, coordinates in notices or place names and are indicative only, not official boundaries. Blank cells mean no data."), "p"),
    (t("網站：https://juliah1121.github.io/Taiwan_MPA/", "Website: https://juliah1121.github.io/Taiwan_MPA/en/"), "p"),
]
ws0.column_dimensions["A"].width = 44; ws0.column_dimensions["B"].width = 110
r = 1
for text, kind in rows0:
    c = ws0.cell(r, 1, text if kind not in ("p",) else None)
    if kind == "title": c.font = Font(name=FONT, size=16, bold=True, color="13315C")
    elif kind == "sub": c.font = Font(name=FONT, size=11, color="5B6B7A")
    elif kind == "h": c.font = Font(name=FONT, size=11, bold=True, color="13315C")
    elif kind == "year":
        c.font = Font(name=FONT, size=10, bold=True); YR = ws0.cell(r, 2, 2026)
        YR.font = Font(name=FONT, size=10, color="0000FF"); YR.fill = PatternFill("solid", fgColor="FFF2CC"); YR.alignment = Alignment(horizontal="left")
        YEAR_REF = f"'{ws0.title}'!$B${r}"
        ws0.cell(r, 3, t("← 可修改", "← editable")).font = Font(name=FONT, size=9, color="5B6B7A")
    elif kind == "p":
        c2 = ws0.cell(r, 1, text); c2.font = BODY; c2.alignment = Alignment(wrap_text=True, vertical="top")
        ws0.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws0.row_dimensions[r].height = 15 * max(1, math.ceil(len(text) / (75 if Z else 150)))
    r += 1

TABN = []
FIELDS = []   # (sheet, header, definition)

def sheet(title, cols, records, widths=None, wrap=(), freeze="C2"):
    """cols: list of (key, header, group, fmt, definition). records: list of dicts. Returns ws and column index map."""
    ws = wb.create_sheet(title)
    for j, (k, h, g, f, defi) in enumerate(cols, 1):
        c = ws.cell(1, j, h); c.font = HF; c.fill = HDR
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = Border(bottom=Side(style="medium", color=GRP.get(g, "DCE6F0")))
        FIELDS.append((title, h, defi))
    idx = {k: j for j, (k, *_ ) in enumerate(cols, 1)}
    for i, rec in enumerate(records, 2):
        for j, (k, h, g, f, defi) in enumerate(cols, 1):
            v = rec.get(k)
            if callable(v): v = v(i)
            c = ws.cell(i, j, val(v)); c.font = BODY; c.border = Border(bottom=thin)
            c.alignment = Alignment(vertical="top", wrap_text=k in wrap)
            if f: c.number_format = f
    ws.row_dimensions[1].height = 42
    ws.freeze_panes = freeze
    for j, (k, *_ ) in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(j)].width = (widths or {}).get(k, 12)
    ref = f"A1:{get_column_letter(len(cols))}{len(records) + 1}"
    TABN.append(1); tab = Table(displayName=f"Table{len(TABN)}", ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(tab)
    return ws, idx

# ---------------- MPAs (main) ----------------
L = lambda zh, en: t(zh, en)
cols = [
    ("no", L("編號", "No."), "basic", "0", L("本資料庫編號（與 MPATable1 一致）", "Database number (as in MPATable1)")),
    ("mpa", L("保護區名稱", "MPA name"), "basic", None, L("簡稱，與網站一致", "Short name as on the website")),
    ("alt", L("英文名稱", "Chinese name"), "basic", None, L("作者自譯英文名稱", "Name in Chinese")),
    ("official_name", L("公告名稱", "Official name"), "basic", None, L("海保署平台／公告所列名稱", "Name on the OCA platform / in the notice (translated)")),
    ("cat_group", L("類別（六大類）", "Category (6 groups)"), "basic", None, L("依法源歸併的六大類", "Six groups by legal basis")),
    ("type", L("類型", "Type"), "basic", None, L("法定類型（MPATable1）", "Legal type (MPATable1)")),
    ("law", L("法源", "Legal basis"), "basic", None, L("劃設依據法規", "Law under which the MPA is designated")),
    ("authority", L("中央主管機關", "Competent authority"), "basic", None, L("海保署平台", "OCA MPA platform")),
    ("manager", L("管理機關", "Managing agency"), "basic", None, L("海保署平台", "OCA MPA platform")),
    ("county", L("縣市", "County"), "basic", None, L("所在縣市（鄉鎮）", "County (township)")),
    ("location", L("位置／範圍", "Location / extent"), "basic", None, L("公告範圍文字", "Boundary description in the notice")),
    ("year", L("成立年份", "Year established"), "basic", "0", L("首次公告或成立年（MPATable1）", "First notice or establishment year (MPATable1)")),
    ("age", L("成立年數", "Years since established"), "basic", "0", L("公式：基準年（說明頁）−成立年份", "Formula: reference year (README) − year established")),
    ("latest_notice", L("最新公告日期", "Latest notice"), "basic", None, L("民國年或西元日期，依公告", "ROC or Gregorian date as in the notice")),
    ("notice_no", L("公告文號", "Notice no."), "basic", None, L("公告文號", "Notice document number")),
    ("area_km2", L("海域面積（km²）", "Marine area (km²)"), "basic", "#,##0.0###", L("MPATable1 海域面積", "Marine area from MPATable1")),
    ("area_note", L("面積說明", "Area note"), "basic", None, L("公告或其他來源的面積說明", "Area as stated in notices or other sources")),
    ("lat", L("代表點緯度", "Latitude"), "basic", "0.000", L("代表點，僅供示意", "Representative point, indicative only")),
    ("lon", L("代表點經度", "Longitude"), "basic", "0.000", L("代表點，僅供示意", "Representative point, indicative only")),
    ("coord_basis", L("座標依據", "Coordinate basis"), "basic", None, L("eDNA 採樣站、公告座標或地名推定", "eDNA station, coordinates in notice, or place name")),
    ("targets", L("保護標的", "Protection targets"), "basic", None, L("公告保護對象", "Protected species / features in the notice")),
    ("restrictions", L("限制事項", "Restrictions"), "basic", None, L("公告主要限制", "Main restrictions in the notice")),
    ("note", L("備註", "Notes"), "basic", None, L("其他說明", "Other notes")),
    ("oca_url", L("海保署平台頁面", "OCA platform page"), "basic", None, L("海洋保護區資訊平台各區頁面（中文）", "Page on the OCA MPA platform (Chinese)")),
    ("iucn", L("IUCN 分類", "IUCN category"), "prot", None, L("MPATable1 建議值", "Suggested value in MPATable1")),
    ("best_level", L("最高保護等級", "Highest protection level"), "prot", None, L("MPA Guide，任一分區的最高等級", "MPA Guide, highest level of any zone")),
    ("dom_level", L("主要保護等級（最大分區）", "Main protection level (largest zone)"), "prot", None, L("MPA Guide，面積最大分區的等級", "MPA Guide, level of the largest zone")),
    ("fh_share", L("完全＋高度保護面積占比", "Share fully or highly protected"), "prot", "0%", L("完全或高度保護分區面積／總面積", "Area of fully or highly protected zones / total area")),
    ("stage", L("建立階段", "Establishment stage"), "prot", None, L("已劃設、已實施、積極管理（MPA Guide）", "Designated, implemented, actively managed (MPA Guide)")),
    ("grade_2022", L("2022 海保署管理成效分級", "OCA 2022 management grade"), "prot", None, L("海保署 2022 年官方評估（僅供對照）", "OCA official 2022 assessment (for reference)")),
    ("score_wq", L("水質分數", "Water quality score"), "score", "0", L("0–100，見網站評分方法", "0–100, see Methods on the website")),
    ("score_debris", L("海廢分數", "Marine debris score"), "score", "0", L("0–100", "0–100")),
    ("score_bio", L("生物多樣性分數", "Biodiversity score"), "score", "0", L("0–100，五分項加權", "0–100, weighted mean of five components")),
    ("score_enf", L("執法分數", "Enforcement score"), "score", "0", L("0–100", "0–100")),
    ("dims", L("有資料面向數", "Dimensions with data"), "calc", "0", L("公式：四項分數的個數", "Formula: count of the four scores")),
    ("score", L("綜合分", "Composite score"), "calc", "0", L("公式：有資料面向的平均，四捨六入五成雙", "Formula: mean of available dimensions, rounded half to even")),
    ("grade", L("成效等級", "Grade"), "calc", None, L("公式：依綜合分與面向數", "Formula: from composite and number of dimensions")),
    ("weakest", L("最弱面向", "Weakest dimension"), "calc", None, L("公式：四項中最低者（至少兩項有資料）", "Formula: lowest of the four (needs at least two)")),
    ("cluster_name", L("成效型態（集群）", "Effectiveness type (cluster)"), "calc", None, L("Ward 集群分析結果；資料不足者空白", "Ward cluster result; blank if too little data")),
]
ci = {k: j for j, (k, *_ ) in enumerate(cols, 1)}
Lc = {k: get_column_letter(j) for k, j in ci.items()}
dims_lbl = [L("水質", "Water quality"), L("海廢", "Marine debris"), L("生物多樣性", "Biodiversity"), L("執法", "Enforcement")]
G = [L("良好", "Good"), L("尚可", "Fair"), L("待加強", "Needs improvement"), L("不佳", "Poor"), L("資料不足", "Insufficient data"), L("（暫評）", " (provisional)")]
recs = []
for _, d in dr.iterrows():
    no = d["no"]; x = cx.loc[no]; e = ef.loc[no]
    rec = d.to_dict()
    rec["alt"] = en_names[no] if Z else zh_names[no]
    rec.update({k: e[k] for k in ("score_wq", "score_debris", "score_bio", "score_enf")})
    rec["fh_share"] = x["fh_share"]; rec["grade_2022"] = x["grade_2022"]
    sc = f"{Lc['score_wq']}{{r}}:{Lc['score_enf']}{{r}}"
    rec["age"] = lambda r: f'=IF({Lc["year"]}{r}="","",{YEAR_REF}-{Lc["year"]}{r})'
    rec["dims"] = lambda r: f"=COUNT({Lc['score_wq']}{r}:{Lc['score_enf']}{r})"
    rec["score"] = lambda r: (f"=IF({Lc['dims']}{r}=0,\"\",IF(MOD(ROUND(AVERAGE({Lc['score_wq']}{r}:{Lc['score_enf']}{r})*2,6),2)=1,"
                              f"2*ROUND(AVERAGE({Lc['score_wq']}{r}:{Lc['score_enf']}{r})/2,0),ROUND(AVERAGE({Lc['score_wq']}{r}:{Lc['score_enf']}{r}),0)))")
    rec["grade"] = lambda r: (f'=IF({Lc["dims"]}{r}<2,"{G[4]}",IF({Lc["score"]}{r}>=70,"{G[0]}",IF({Lc["score"]}{r}>=55,"{G[1]}",'
                              f'IF({Lc["score"]}{r}>=40,"{G[2]}","{G[3]}")))&IF({Lc["dims"]}{r}=2,"{G[5]}",""))')
    rec["weakest"] = lambda r: (f'=IF({Lc["dims"]}{r}<2,"",INDEX(DIMLBL,MATCH(MIN({Lc["score_wq"]}{r}:{Lc["score_enf"]}{r}),'
                                f'{Lc["score_wq"]}{r}:{Lc["score_enf"]}{r},0)))')
    recs.append(rec)
widths = {"no": 6, "mpa": 30, "alt": 30, "official_name": 30, "cat_group": 18, "type": 18, "law": 18, "authority": 22, "manager": 22,
          "county": 16, "location": 40, "latest_notice": 14, "notice_no": 22, "area_note": 30, "coord_basis": 14, "targets": 40,
          "restrictions": 48, "note": 30, "oca_url": 26, "best_level": 14, "dom_level": 16, "stage": 12, "grade_2022": 12,
          "grade": 16, "weakest": 14, "cluster_name": 20}
ws1, _ = sheet(L("保護區總表", "MPAs"), cols, recs, widths, wrap=("location", "targets", "restrictions", "note", "area_note"), freeze="C2")
# dimension labels for the weakest-dimension lookup (hidden helper row)
from openpyxl.workbook.defined_name import DefinedName
lab_row = len(recs) + 4
ws1.cell(lab_row, 1, L("最弱面向標籤（公式用）", "Labels for weakest-dimension formula")).font = Font(name=FONT, size=9, color="5B6B7A")
for j, k in enumerate(("score_wq", "score_debris", "score_bio", "score_enf")):
    ws1.cell(lab_row, ci[k], dims_lbl[j]).font = Font(name=FONT, size=9, color="5B6B7A")
wb.defined_names["DIMLBL"] = DefinedName("DIMLBL", attr_text=f"'{ws1.title}'!${Lc['score_wq']}${lab_row}:${Lc['score_enf']}${lab_row}")
for i in range(2, len(recs) + 2):
    c = ws1.cell(i, ci["oca_url"]);
    if c.value: c.hyperlink = c.value; c.font = Font(name=FONT, size=10, color="1C5CAB", underline="single")
n = len(recs) + 1
rng = f"{Lc['score_wq']}2:{Lc['score']}{n}"
ws1.conditional_formatting.add(rng, ColorScaleRule(start_type="num", start_value=0, start_color="F8D3CF", mid_type="num", mid_value=55, mid_color="FFFFFF", end_type="num", end_value=100, end_color="CDE8DA"))
gcol = f"{Lc['grade']}2:{Lc['grade']}{n}"
for key, color in ((G[0], "CDE8DA"), (G[1], "E9EDF0"), (G[2], "FBE7C6"), (G[3], "F8D3CF")):
    ws1.conditional_formatting.add(gcol, FormulaRule(formula=[f'LEFT({Lc["grade"]}2,{len(key)})="{key}"'], fill=PatternFill("solid", fgColor=color)))
ws1.cell(1, ci["score"]).comment = Comment(L("綜合分＝有資料面向的平均，以四捨六入五成雙取整（例：80.5→80、59.5→60），與網站一致。", "Composite = mean of available dimensions, rounded half to even (80.5→80, 59.5→60), matching the website."), "Taiwan MPA")
ws1.cell(1, ci["age"]).comment = Comment(L("成立年數＝說明頁的基準年−成立年份。", "Years since established = reference year on the README sheet − year established."), "Taiwan MPA")

# ---------------- Zones ----------------
zrecs = []
for _, d in dr.iterrows():
    z = val(d["zones"])
    if z:
        parts = z.split("；" if Z else "; ")
        for p in parts:
            f = [s.strip() for s in p.split("｜" if Z else " | ")]
            area = None
            if len(f) > 2:
                try: area = float(f[2].replace("km²", "").strip())
                except ValueError: area = None
            zrecs.append({"no": d["no"], "mpa": d["mpa"], "zone": f[0], "level": f[1] if len(f) > 1 else None, "area": area})
    else:
        zrecs.append({"no": d["no"], "mpa": d["mpa"], "zone": L("全區", "Whole site"), "level": d["best_level"], "area": d["area_km2"]})
zcols = [("no", L("編號", "No."), "basic", "0", L("保護區編號", "MPA number")), ("mpa", L("保護區名稱", "MPA name"), "basic", None, ""),
         ("zone", L("分區", "Zone"), "prot", None, L("依園區計畫或公告的分區；單一分區者列「全區」", "Zone per park plan or notice; single-zone sites listed as whole site")),
         ("level", L("保護等級（MPA Guide）", "Protection level (MPA Guide)"), "prot", None, L("完全、高度、輕度、低度保護", "Fully, highly, lightly or minimally protected")),
         ("area", L("分區面積（km²）", "Zone area (km²)"), "prot", "#,##0.0###", L("MPATable1 分區面積", "Zone area in MPATable1"))]
sheet(L("分區保護等級", "Zones"), zcols, zrecs, {"no": 6, "mpa": 32, "zone": 48, "level": 20, "area": 14}, wrap=("zone",))

# ---------------- Raw indicators ----------------
rcols = [("no", L("編號", "No."), "basic", "0", ""), ("mpa", L("保護區名稱", "MPA name"), "basic", None, ""),
 ("wq_status", L("水質現況", "Water status"), "score", None, L("近 3 年測站超標情形：良好／注意／不佳", "Last 3 years: good / caution / poor")),
 ("wq_trend", L("水質長期趨勢", "Water trend"), "score", None, L("2002–2026 長期趨勢", "Long-term trend 2002–2026")),
 ("wq_exceed", L("水質超標項目", "Exceedances"), "score", None, ""),
 ("debris_kg", L("海廢平均 kg／次", "Debris kg per survey"), "score", "0.00", L("淨灘監測 2020–2024 平均", "Beach-litter monitoring mean, 2020–2024")),
 ("debris_change_pct", L("海廢 2020→2024 變化 %", "Debris change 2020→2024 %"), "score", "0.0", ""),
 ("debris_trend", L("海廢趨勢", "Debris trend"), "score", None, ""),
 ("enf_type", L("執法型態", "Enforcement type"), "score", None, L("A–E 五型", "Types A–E")),
 ("patrol_hours", L("巡守時數", "Patrol hours"), "score", "#,##0", L("112–114 年合計", "Total 2023–2025")),
 ("cases", L("取締件數", "Violations detected"), "score", "0", ""), ("punished", L("處分件數", "Sanctioned"), "score", "0", ""),
 ("pun_rate", L("處分率", "Sanction rate"), "score", "0%", ""),
 ("bio_s_species", L("物種紀錄分", "Species-record score"), "score", "0", L("生物多樣性分項（權重 1）", "Biodiversity component (weight 1)")),
 ("bio_conf_sp", L("已確認物種數", "Confirmed species"), "score", "0", L("TBN 海洋魚類與蝦蟹類", "TBN marine fish and crustaceans")),
 ("bio_threatened", L("IUCN 受脅物種數", "IUCN threatened species"), "score", "0", ""),
 ("bio_edna_only", L("eDNA 獨有物種數", "eDNA-only species"), "score", "0", ""),
 ("s_edna_v2", L("eDNA 試驗分", "eDNA pilot score"), "score", "0", L("分項（權重 0.5）", "Component (weight 0.5)")),
 ("bio_s_coral", L("珊瑚健康分", "Coral health score"), "score", "0", L("分項（權重 1）", "Component (weight 1)")),
 ("bio_s_bluecarbon", L("藍碳分", "Blue carbon score"), "score", "0", L("分項（權重 1）", "Component (weight 1)")),
 ("bio_s_habitat", L("棲地多樣性分", "Habitat diversity score"), "score", "0", L("分項（權重 1）", "Component (weight 1)")),
 ("edna_samples", L("eDNA 採樣次數（保護區站）", "eDNA samples (MPA stations)"), "score", "0", ""),
 ("edna_taxa", L("eDNA 分類單元數", "eDNA taxa"), "score", "0", L("保護區站與潛水站合計", "MPA and dive stations combined")),
 ("edna_ms_taxa", L("eDNA 分類單元（保護區站）", "eDNA taxa (MPA stations)"), "score", "0", ""),
 ("edna_dive_taxa", L("eDNA 分類單元（潛水站）", "eDNA taxa (dive stations)"), "score", "0", ""),
 ("rc2008", L("珊瑚覆蓋 2008 %", "Coral cover 2008 %"), "score", "0.0", L("活珊瑚", "Live coral")),
 ("rc2023", L("珊瑚覆蓋 2023 %", "Coral cover 2023 %"), "score", "0.0", L("活珊瑚", "Live coral")),
 ("rc2025", L("珊瑚覆蓋 2024–26 %", "Coral cover 2024–26 %"), "score", "0.0", L("硬珊瑚", "Hard coral")),
 ("rc_trend", L("珊瑚趨勢", "Coral trend"), "score", None, ""), ("rc_grade", L("珊瑚等級", "Coral grade"), "score", None, ""),
 ("bc_tier", L("藍碳級", "Blue-carbon tier"), "score", None, L("A ≥1 萬 tC、B 0.1–1 萬、C 更少或未量化", "A ≥10,000 tC, B 1,000–10,000, C less or unquantified")),
 ("est_tC", L("估計藍碳（公噸 C）", "Estimated blue carbon (t C)"), "score", "#,##0", L("面積 × 全國平均碳密度", "Area × national mean carbon density")),
 ("mangrove_ha", L("紅樹林（公頃）", "Mangrove (ha)"), "score", "#,##0.00", ""), ("seagrass_ha", L("海草床（公頃）", "Seagrass (ha)"), "score", "#,##0.00", ""),
 ("coral", L("珊瑚礁", "Coral reef"), "basic", None, L("主要／有／空白", "Main / present / blank")), ("seagrass", L("海草床", "Seagrass"), "basic", None, ""),
 ("rocky", L("岩礁", "Rocky reef"), "basic", None, ""), ("mangrove", L("紅樹林", "Mangrove"), "basic", None, ""),
 ("mudflat", L("泥灘", "Mudflat"), "basic", None, ""), ("saltmarsh", L("鹽沼", "Salt marsh"), "basic", None, ""),
 ("n_types", L("生態系類型數", "Ecosystem types"), "basic", "0", ""), ("eco_note", L("生態系判定依據", "Ecosystem basis"), "basic", None, ""),
]
rrecs = []
for _, d in dr.iterrows():
    no = d["no"]; e = ef.loc[no].to_dict(); b = bh.loc[no].to_dict()
    rec = {"no": no, "mpa": d["mpa"], **e, **{k: b[k] for k in b if k in dict((c[0], 1) for c in rcols)}}
    rec["edna_taxa"] = b.get("edna_taxa")
    if no in bc.index:
        x = bc.loc[no]; rec.update(est_tC=x["est_tC"], mangrove_ha=x["mangrove_ha"], seagrass_ha=x["seagrass_ha"])
    rrecs.append(rec)
sheet(L("指標原始資料", "Raw indicators"), rcols, rrecs, {"no": 6, "mpa": 30, "wq_exceed": 18, "enf_type": 20, "eco_note": 46, "rc_trend": 12}, wrap=("eco_note",))

# ---------------- Pressure ----------------
pcols = [("no", L("編號", "No."), "basic", "0", ""), ("mpa", L("保護區名稱", "MPA name"), "basic", None, ""),
 ("townships", L("所在或鄰接鄉鎮", "Townships"), "basic", None, L("保護區所在或鄰接的鄉鎮市區", "Townships containing or adjoining the MPA")),
 ("pop", L("鄰近人口", "Nearby population"), "basic", "#,##0", L("維基百科鄉鎮列表（未標年份）", "Wikipedia township list (year not stated)")),
 ("density", L("人口密度（人／km²）", "Population density (per km²)"), "basic", "#,##0", ""),
 ("dens_band", L("人口密度分組", "Density group"), "basic", None, ""),
 ("fish_share", L("海廢漁業來源占比 %", "Fishing share of debris %"), "score", "0.0", L("淨灘垃圾中漁業與養殖相關品項占比", "Share of fishing and aquaculture items in beach litter")),
 ("fish_kg", L("漁業來源海廢 kg／次", "Fishing debris kg per survey"), "score", "0.00", ""),
 ("patrol_hours", L("巡守時數", "Patrol hours"), "score", "#,##0", L("112–114 年", "2023–2025")),
 ("members", L("巡守人數", "Patrol staff"), "score", "0", ""),
 ("hours_per_km2", L("每 km² 巡守時數", "Patrol hours per km²"), "score", "#,##0.0", ""),
 ("east", L("位於花東", "East coast"), "basic", "0", L("1＝花蓮、臺東", "1 = Hualien or Taitung")),
 ("eco", L("生態環境分數", "Ecological score"), "calc", "0.0", L("水質、海廢、生物多樣性平均（不含執法）", "Mean of water, debris and biodiversity (excl. enforcement)"))]
prs = [{"no": no, "mpa": dr.set_index("no").loc[no, "mpa"], **pr.loc[no].to_dict()} for no in dr["no"]]
sheet(L("人口與漁業壓力", "Pressure"), pcols, prs, {"no": 6, "mpa": 30, "townships": 40}, wrap=("townships",))

# ---------------- Supplementary ----------------
ws = wb.create_sheet(L("輔助調查", "Supplementary surveys"))
def block(ws, r0, title, cols, rows):
    ws.cell(r0, 1, title).font = Font(name=FONT, size=11, bold=True, color="13315C")
    for j, (k, h, f) in enumerate(cols, 1):
        c = ws.cell(r0 + 1, j, h); c.font = HF; c.fill = HDR; c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, rw in enumerate(rows, r0 + 2):
        for j, (k, h, f) in enumerate(cols, 1):
            c = ws.cell(i, j, val(rw.get(k))); c.font = BODY; c.border = Border(bottom=thin); c.alignment = Alignment(vertical="top", wrap_text=k == "sites")
            if f: c.number_format = f
    ws.row_dimensions[r0 + 1].height = 32
    return r0 + 2 + len(rows) + 1
btr = [{**x, "no": name2no.get(x["mpa"])} for x in bt.to_dict("records")]
r = block(ws, 1, L("底棲群聚（潛水調查，保護區內或 1 km 內站點）", "Benthic communities (dive surveys, stations inside or within 1 km)"),
          [("no", L("編號", "No."), "0"), ("mpa", L("保護區名稱", "MPA name"), None), ("n_sites", L("站數", "Stations"), "0"),
           ("taxa_mean", L("平均分類群數", "Mean taxa"), "0.0"), ("taxa_max", L("最多分類群數", "Max taxa"), "0"),
           ("density_med", L("密度中位數 ind./m²", "Median density ind./m²"), "#,##0"), ("biomass_med", L("生物量中位數（原表單位）", "Median biomass (source units)"), "0.0"),
           ("sites", L("站點", "Stations"), None)], btr)
ofr = [{**x, "no": name2no.get(x["mpa"])} for x in off.to_dict("records")]
r = block(ws, r + 1, L("外海背景（10 km 內船舶測站；百分位為 25 站中的排名）", "Offshore background (ship stations within 10 km; percentile among 25 stations)"),
          [("no", L("編號", "No."), "0"), ("mpa", L("保護區名稱", "MPA name"), None), ("near_km", L("最近測站 km", "Nearest station km"), "0.00"),
           ("wq_station", L("水質測站", "Water station"), None), ("wq_din", L("表層氮營養鹽 μM", "Surface DIN μM"), "0.00"), ("wq_din_pct", L("氮百分位", "DIN percentile"), "0"),
           ("wq_po4", L("表層磷酸 μM", "Surface PO₄ μM"), "0.000"), ("wq_po4_pct", L("磷百分位", "PO₄ percentile"), "0"),
           ("mp_station", L("塑膠微粒測站", "Microplastics station"), None), ("mp_mp", L("表層塑膠微粒 個／m³", "Surface microplastics per m³"), "#,##0"),
           ("mp_mp_pct", L("塑膠微粒百分位", "Microplastics percentile"), "0"), ("sf_station", L("海底廢棄物測站", "Seafloor litter station"), None),
           ("sf_sf_n", L("海底廢棄物數量密度", "Seafloor litter density"), "0.0"), ("sf_sf_n_pct", L("海底廢棄物百分位", "Seafloor litter percentile"), "0")], ofr)
for j, w in enumerate([6, 32, 10, 12, 14, 12, 14, 12, 14, 14, 12, 14, 14, 12], 1): ws.column_dimensions[get_column_letter(j)].width = w
ws.column_dimensions["H"].width = max(ws.column_dimensions["H"].width, 40)

# ---------------- Field guide ----------------
wf = wb.create_sheet(L("欄位說明", "Field guide"))
for j, h in enumerate([L("工作表", "Sheet"), L("欄位", "Field"), L("說明", "Definition")], 1):
    c = wf.cell(1, j, h); c.font = HF; c.fill = HDR
for i, (s, h, dfn) in enumerate(FIELDS, 2):
    for j, v in enumerate((s, h, dfn), 1):
        c = wf.cell(i, j, v); c.font = BODY; c.alignment = Alignment(wrap_text=True, vertical="top")
wf.column_dimensions["A"].width = 18; wf.column_dimensions["B"].width = 30; wf.column_dimensions["C"].width = 70; wf.freeze_panes = "A2"

for w in wb.worksheets:
    w.sheet_properties.tabColor = "13315C" if w.title in (ws1.title,) else None
wb.save(OUT); print("saved", OUT)
