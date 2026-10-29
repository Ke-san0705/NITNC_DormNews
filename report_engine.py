"""Microsoft Forms回答(CSV/XLSX)から寮生会新聞の本部部会データを生成する。

意図的にExcel/Formsの日本語データを推測して復元しない。
部署が認識できない、または日本語が ? に置換された場合は警告する。
"""
from __future__ import annotations

import csv
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime, date, timedelta
from pathlib import Path

# フォーム選択肢：16部署（秀峰寮東と秀峰寮西は独立）。キーはフォームの回答値。
FORM_DEPARTMENTS = [
    "三役", "翔峰寮", "栄峰寮", "光峰寮", "明峰寮", "清峰寮", "秀峰寮東", "秀峰寮西", "優峰寮",
    "風紀長", "総務長", "企画長", "学習支援部長", "会計長", "IT管理長", "庶務長",
]

# 上段左→右の紙面配置。三役は従来どおり右の独立欄を利用。
PAPER_ORDER = [
    "翔峰寮", "栄峰寮", "風紀長",
    "光峰寮", "明峰寮", "庶務長",
    "会計長", "清峰寮", "優峰寮",
    "総務長", "学習支援部長", "秀峰寮西",
    "秀峰寮東", "企画長", "IT管理長",
]
PAPER_TITLES = {
    "三役": "寮生会三役", "翔峰寮": "翔峰棟長", "栄峰寮": "栄峰棟長", "光峰寮": "光峰棟長",
    "明峰寮": "明峰棟長", "清峰寮": "清峰棟長", "秀峰寮東": "秀峰寮東", "秀峰寮西": "秀峰寮西", "優峰寮": "優峰棟長",
    "風紀長": "風紀長", "総務長": "総務長", "企画長": "企画長", "学習支援部長": "学習支援部長",
    "会計長": "会計長", "IT管理長": "IT管理長", "庶務長": "庶務長",
}

ALIASES = {name: name for name in FORM_DEPARTMENTS}
ALIASES.update({
    "寮生会三役": "三役", "三役": "三役",
    "翔峰棟長": "翔峰寮", "栄峰棟長": "栄峰寮", "光峰棟長": "光峰寮",
    "明峰棟長": "明峰寮", "清峰棟長": "清峰寮", "優峰棟長": "優峰寮",
    # 旧Formsの表記も履歴を保持するため、東西ごとに新しい正式名称へ変換する。
    "秀峰東寮": "秀峰寮東", "秀峰西寮": "秀峰寮西",
    "秀峰東棟長": "秀峰寮東", "秀峰西棟長": "秀峰寮西",
    "秀峰寮東棟長": "秀峰寮東", "秀峰寮西棟長": "秀峰寮西",
    "秀峰東": "秀峰寮東", "秀峰西": "秀峰寮西",
})


class InputProblem(ValueError):
    """入力が不正で掲載判断できない。"""


def norm(value) -> str:
    return re.sub(r"[\s\u3000]+", "", unicodedata.normalize("NFKC", str(value or ""))).lower()


def bad_text(s) -> bool:
    """文字化けした '???' 等。疑問文や通常の半角?1字は拒否しない。"""
    s = str(s or "").strip()
    if not s:
        return False
    if re.fullmatch(r"[?？]+", s):
        return True
    return s.count("?") >= 5 and s.count("?") >= max(5, len(s) * 0.35)


def read_xlsx_xml(path: Path) -> list[list]:
    """xlsx zip内のシート先頭を解析。sharedStrings・inlineStr・Excel日時数値に対応。"""
    from zipfile import ZipFile, BadZipFile
    from xml.etree import ElementTree as ET
    from posixpath import normpath, dirname, join

    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    q = lambda tag: "{" + main + "}" + tag

    try:
        with ZipFile(path) as z:
            files = set(z.namelist())
            if "xl/workbook.xml" not in files:
                raise InputProblem("このxlsxにはワークブックの定義が見つかりません。")
            shared = []
            if "xl/sharedStrings.xml" in files:
                root = ET.fromstring(z.read("xl/sharedStrings.xml"))
                for si in root.findall(q("si")):
                    shared.append("".join(x.text or "" for x in si.iter(q("t"))))
            book = ET.fromstring(z.read("xl/workbook.xml"))
            first = book.find(q("sheets") + "/" + q("sheet"))
            if first is None:
                raise InputProblem("xlsxのワークシートがありません。")
            rid = first.get("{" + rel_ns + "}id")
            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            rel = next((r for r in rels.findall("{" + pkg_ns + "}Relationship") if r.get("Id") == rid), None)
            if rel is None:
                raise InputProblem("xlsxのシート定義に対応するファイルを見つけられません。")
            target = rel.get("Target", "")
            sheet_path = target.lstrip("/") if target.startswith("/") else normpath(join("xl", target))
            if sheet_path not in files:
                raise InputProblem("xlsxの内部ファイルが見つかりません: " + sheet_path)
            root = ET.fromstring(z.read(sheet_path))
            values = []
            for row in root.findall(".//" + q("sheetData") + "/" + q("row")):
                r = []
                for cell in row.findall(q("c")):
                    cell_ref = cell.get("r", "A1")
                    letters = re.match(r"[A-Z]+", cell_ref)
                    if not letters:
                        continue
                    index = 0
                    for ch in letters.group():
                        index = index * 26 + (ord(ch) - ord("A") + 1)
                    index -= 1
                    if index > 60:
                        continue
                    while len(r) <= index:
                        r.append("")
                    typ = cell.get("t", "")
                    value_node = cell.find(q("v"))
                    val = value_node.text if value_node is not None else ""
                    if typ == "inlineStr":
                        inline = cell.find(q("is"))
                        val = "".join(n.text or "" for n in inline.iter(q("t"))) if inline is not None else ""
                    elif typ == "s" and val:
                        try:
                            val = shared[int(val)]
                        except (ValueError, IndexError):
                            raise InputProblem("xlsxの共有文字列を取得できません。")
                    elif typ in ("n", "") and val:
                        try:
                            val = float(val)
                        except ValueError:
                            pass
                    r[index] = val if val is not None else ""
                if r and any(str(x).strip() for x in r):
                    values.append(r)
            return values
    except BadZipFile as exc:
        raise InputProblem("xlsxとして読み取れません。元のExcelを保存し直してください。") from exc


def load_table(path: Path) -> list[list]:
    if path.suffix.lower() == ".csv":
        errors = []
        for enc in ("utf-8-sig", "cp932"):
            try:
                with path.open("r", encoding=enc, newline="") as f:
                    rows = list(csv.reader(f))
                if rows and sum(bad_text(h) for h in rows[0]) >= 2:
                    raise InputProblem(
                        "CSVの列名・本文が実際に『????』へ置換されています。文字コード変更では復元できません。"
                        "Microsoft Formsの［回答］→［Excelで開く］から元の.xlsxを取得してください。"
                    )
                return rows
            except UnicodeDecodeError as e:
                errors.append(f"{enc}: {e}")
        raise InputProblem("CSVの文字コードを認識できません: " + "; ".join(errors))
    if path.suffix.lower() == ".xlsx":
        # Microsoft Forms Excelの通常の.xlsxを標準ライブラリで読み取り（ネット接続不要）。
        # 回答元ブックには一切書き込まない。
        return read_xlsx_xml(path)
    raise InputProblem("入力ファイルは .xlsx または .csv を指定してください。")


def column_map(headers: list[str]) -> dict[str, int]:
    h = [norm(x) for x in headers]
    def find_one(*keywords):
        return next((i for i, name in enumerate(h) if any(norm(k) in name for k in keywords)), None)
    # 質問文をわずかに変更しても対応できるようキーワードで判定。
    result = {
        "id": find_one("id", "回答id", "応答id"),
        "start": find_one("開始時刻", "starttime"),
        "end": find_one("完了時刻", "終了時刻", "submittime", "completiontime"),
        "department": find_one("部署", "所属"),
        "activity": find_one("活動報告"),
        "request": find_one("連絡事項", "注意お願い", "注意・お願い"),
        "other": find_one("その他", "イベント情報"),
    }
    missing = [k for k in ("department", "activity", "request") if result[k] is None]
    if missing:
        raise InputProblem(
            "Forms回答の列名から必須項目を特定できません（" + ", ".join(missing) + "）。\n"
            "検出した列名: " + " / ".join(str(x) for x in headers) + "\n"
            "質問名に『部署』『活動報告』『連絡事項』を含めてください。"
        )
    if result["end"] is None:
        raise InputProblem("『完了時刻』の列が見つかりません。最新版の判定ができません。")
    return result


def parse_time(value) -> datetime:
    if isinstance(value, (int, float)) and 30000 < float(value) < 90000:
        return datetime(1899, 12, 30) + timedelta(days=float(value))
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    raw = str(value or "").strip()
    raw = raw.replace("年", "/").replace("月", "/").replace("日", " ")
    raw = raw.replace("午前", "AM ").replace("午後", "PM ")
    raw = raw.replace("T", " ").replace("Z", "")
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d %H:%M", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M",
                "%Y/%m/%d %I:%M:%S %p", "%Y/%m/%d %I:%M %p"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(raw)
    except ValueError:
        raise InputProblem(f"完了時刻『{raw}』を解釈できません。")


def get(row: list, idx: int | None) -> str:
    return str(row[idx] if idx is not None and idx < len(row) else "").strip()


@dataclass
class MergeResult:
    latest: dict
    audit: dict


def merge_rows(table: list[list], volume: int, max_chars: int = 105) -> MergeResult:
    if not table:
        raise InputProblem("回答データに行がありません。")
    index = column_map(table[0])
    if len(table) < 2:
        # 回答0件でも未提出表示ができる。
        rows = []
    else:
        rows = table[1:]
    records = []
    errors = []
    for index_row, row in enumerate(rows, 2):
        if not any(str(c or "").strip() for c in row):
            continue
        dep_raw = get(row, index["department"])
        activity = get(row, index["activity"])
        request = get(row, index["request"])
        other = get(row, index["other"])
        if any(bad_text(x) for x in (dep_raw, activity, request, other)):
            errors.append(f"{index_row}行目：日本語が『???』になっています。CSVの書き出しを確認してください。")
            continue
        dep = ALIASES.get(dep_raw)
        if dep is None:
            if dep_raw in ("秀峰寮", "秀峰棟長"):
                errors.append(f"{index_row}行目：『{dep_raw}』は東西を区別できません。Microsoft Formsの選択肢を『秀峰寮東』『秀峰寮西』に分け、回答を再送信してください。")
            else:
                errors.append(f"{index_row}行目：未定義の部署『{dep_raw}』を検出。誤掲載防止のため採用しません。")
            continue
        if not activity or not request:
            errors.append(f"{index_row}行目：部署『{dep}』の必須項目①または②が空欄です。")
            continue
        try:
            completed = parse_time(get(row, index["end"]))
        except InputProblem as exc:
            errors.append(f"{index_row}行目：{exc}")
            continue
        response_id = get(row, index["id"])
        records.append({"department": dep, "activity": activity, "request": request, "other": other,
                        "completed": completed, "response_id": response_id, "row": index_row})
    if errors:
        raise InputProblem("\n".join(errors))
    records.sort(key=lambda r: (r["completed"], int(r["response_id"]) if r["response_id"].isdigit() else -1, r["row"]))
    latest = {}
    versions = {dep: [] for dep in FORM_DEPARTMENTS}
    for record in records:
        latest[record["department"]] = record
        versions[record["department"]].append(record)
    out = {}
    over = []
    for dep in FORM_DEPARTMENTS:
        r = latest.get(dep)
        if r is None:
            out[dep] = {"department": dep, "title": PAPER_TITLES[dep], "activity": "", "request": "", "other": "",
                        "display": "（未提出）", "missing": True, "chars": 0, "completed": "", "response_id": ""}
            continue
        chunks = [f"①{r['activity']}", f"②{r['request']}"]
        if r["other"]:
            chunks.append(f"③{r['other']}")
        chars = sum(len(r[k]) for k in ("activity", "request", "other"))
        if chars > max_chars:
            over.append({"department": dep, "chars": chars, "recommended_max": max_chars})
        out[dep] = {"department": dep, "title": PAPER_TITLES[dep], "activity": r["activity"], "request": r["request"],
                    "other": r["other"], "display": "\n".join(chunks), "missing": False, "chars": chars,
                    "completed": r["completed"].isoformat(timespec="minutes"), "response_id": r["response_id"]}
    audit = {"volume": volume, "total_responses": len(records), "submitted": len(latest),
             "missing": [dep for dep in FORM_DEPARTMENTS if dep not in latest],
             "over_limit": over,
             "duplicates": [{"department": dep, "count": len(v)} for dep, v in versions.items() if len(v) > 1],
             "selection_rule": "同一部署の完了時刻が最も新しい回答を①②③一式で採用。以前の回答は履歴として残す。",
             "layout": "三役は右側、残り15部署は3列×5段（空き枠なし）。秀峰寮西は4段目右、秀峰寮東は5段目左に独立表示。"}
    return MergeResult(latest=out, audit=audit)


def load_and_merge(path: Path, volume: int, max_chars: int = 105) -> MergeResult:
    return merge_rows(load_table(path), volume, max_chars)
