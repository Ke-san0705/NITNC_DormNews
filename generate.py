"""寮生会新聞: Microsoft Forms回答Excel→確定2デザインのHTML/PDF/PNG。

Forms のアカウントに直接ログインするものではない。同期済み .xlsx を読込む。
発行前の校正は必ず人間が行う。
"""
from __future__ import annotations

import argparse
import base64
from html import escape
import json
import mimetypes
import re
import sys
import time
from datetime import datetime
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape
from report_engine import InputProblem, PAPER_ORDER, FORM_DEPARTMENTS, load_and_merge

ROOT = Path(__file__).resolve().parent
TEMPLATES = {"classic": "①新聞風", "blue": "④現行改良版"}


def validate_edition(edition_file: Path, volume: int) -> tuple[dict, list[dict]]:
    if not edition_file.is_file():
        raise InputProblem(f"紙面の設定ファイルがありません: {edition_file}")
    try:
        edition = json.loads(edition_file.read_text(encoding="utf-8-sig"))
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        raise InputProblem(f"edition.jsonを読み取れません: {e}") from e
    if edition.get("volume") != volume:
        raise InputProblem(f"号数の不一致：起動時はVol.{volume}ですが、edition.jsonはVol.{edition.get('volume')}です。")
    # 発行年月は年月を独立管理し、題字の日付と左下表記の不一致を防止する。
    # 旧版の issue_date のみを含む設定も読み込み可能にする。
    has_ym = "issue_year" in edition and "issue_month" in edition
    if has_ym:
        year, month = edition["issue_year"], edition["issue_month"]
    else:
        match = re.fullmatch(r"(\d{4})年(\d{1,2})月", str(edition.get("issue_date", "")).strip())
        if not match:
            raise InputProblem("発行年月が未設定です。edition.jsonのissue_year、issue_monthを指定してください。")
        year, month = (int(match.group(1)), int(match.group(2)))
    if type(year) is not int or not 2000 <= year <= 2100:
        raise InputProblem("issue_yearは2000〜2100の整数で指定してください。")
    if type(month) is not int or not 1 <= month <= 12:
        raise InputProblem("issue_monthは1〜12の整数で指定してください。")
    edition["issue_year"] = year
    edition["issue_month"] = month
    edition["issue_date"] = f"{year}年{month}月"
    for k in ("intro", "notices", "teacher1", "teacher2", "article1", "article2"):
        if k not in edition:
            raise InputProblem(f"edition.jsonの必須キー {k} がありません")
    if not isinstance(edition["notices"], list) or len(edition["notices"]) > 3:
        raise InputProblem("お知らせは最大3件です。edition.jsonのnoticesを確認してください。")
    articles = []
    for i in (1, 2):
        item = edition[f"article{i}"].copy()
        for required in ("category", "title", "body"):
            if not str(item.get(required, "")).strip():
                raise InputProblem(f"記事{i}の {required} が未入力です")
        img = str(item.get("image", "")).strip()
        item["img_src"] = None
        if img:
            filepath = (edition_file.parent / img).resolve()
            # 相対ディレクトリ外の写真参照は許可しない
            try:
                filepath.relative_to(edition_file.parent.resolve())
            except ValueError:
                raise InputProblem(f"画像はこのプロジェクト内に配置してください: {img}")
            if not filepath.is_file():
                raise InputProblem(f"記事{i}の画像が見つかりません: {filepath}")
            mime = mimetypes.guess_type(filepath.name)[0]
            if mime not in {"image/jpeg", "image/png", "image/webp"}:
                raise InputProblem("画像形式は.jpg、.png、.webpを使用してください。")
            if filepath.stat().st_size > 8_000_000:
                raise InputProblem(f"画像サイズが大きすぎます（8MB上限）: {img}")
            item["img_src"] = f"data:{mime};base64,{base64.b64encode(filepath.read_bytes()).decode('ascii')}"
        else:
            item["date_label"] = item.get("date_label", "PHOTO")
        articles.append(item)
    return edition, articles


def write_json(path: Path, item: dict):
    path.write_text(json.dumps(item, ensure_ascii=False, indent=2), encoding="utf-8")


def browser_render(html_file: Path, pdf_file: Path, png_file: Path, final: bool) -> list[dict]:
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        kwargs = {"headless": True}
        if Path("/usr/bin/chromium").exists():
            kwargs["executable_path"] = "/usr/bin/chromium"
        browser = p.chromium.launch(**kwargs)
        page = browser.new_page(viewport={"width": 794, "height": 1123}, device_scale_factor=1.7)
        page.set_content(html_file.read_text(encoding="utf-8"), wait_until="load")
        page.evaluate("document.fonts.ready")
        page.locator(".article img").evaluate_all("els => Promise.all(els.map(img => img.decode().catch(()=>{})))")
        problems = page.evaluate("window.inspectNewspaper()")
        page.locator(".paper").screenshot(path=str(png_file), animations="disabled")
        if problems:
            if pdf_file.exists():
                pdf_file.unlink()  # 古い成功PDFの誤配布を防ぐ
            browser.close()
            return problems
        page.emulate_media(media="print")
        page.pdf(path=str(pdf_file), print_background=True, prefer_css_page_size=True,
                 margin={"left": "0", "top": "0", "right": "0", "bottom": "0"})
        browser.close()
    return []


def compile_once(input_file: Path, edition_file: Path, out_root: Path, volume: int,
                 designs: list[str], publish: bool, max_chars: int) -> dict:
    out_root.mkdir(parents=True, exist_ok=True)
    result = load_and_merge(input_file, volume, max_chars)
    edition, articles = validate_edition(edition_file, volume)
    audit = dict(result.audit)
    audit.update({
        "source_file": input_file.name,
        "issue_year": edition["issue_year"],
        "issue_month": edition["issue_month"],
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "designs": designs,
        "publication_requested": publish,
        "status": "PREVIEW",
        "notes": "原稿の事実確認・掲載可否の最終判断は編集者が行うこと",
    })
    if publish and audit["missing"]:
        raise InputProblem("発行モードでは未提出部署があるため中止しました: " + "、".join(audit["missing"]))
    if publish and audit["over_limit"]:
        raise InputProblem("発行モードでは文字数の目安超過を解消してください: " + "、".join(x["department"] for x in audit["over_limit"]))
    if publish and any("編集担当" in str(edition.get(k, "")) for k in ("teacher1", "teacher2")):
        raise InputProblem("発行モードでは教員コメントの仮文言を実際の文章に変更してください。")

    reports = [result.latest[name] for name in PAPER_ORDER]
    leader = result.latest["三役"]
    write_json(out_root / "提出状況.json", audit)
    write_json(out_root / "反映データ.json", {
        "volume": volume,
        "issue_year": edition["issue_year"],
        "issue_month": edition["issue_month"],
        "paper_order": PAPER_ORDER,
        "reports": reports,
        "leader": leader,
        "notice": edition["notices"],
        "message": "個人メールアドレスや回答者名はこのファイルに含まれません",
    })
    # 回答の提出状況だけをブラウザで確認できる一覧（個人情報は含めない）
    status_rows = []
    for dep, item in result.latest.items():
        is_missing = item["missing"]
        status_rows.append(
            f"<tr><td>{escape(dep)}</td><td class={chr(34)+('missing' if is_missing else 'submitted')+chr(34)}>"
            f"{'未提出' if is_missing else '提出済み'}</td>"
            f"<td>{escape(item.get('completed') or '—')}</td><td>{item['chars']}</td></tr>"
        )
    status_html = """<!doctype html><html lang=ja><meta charset=utf-8><title>提出状況</title>
    <style>body{font-family:system-ui,sans-serif;max-width:820px;margin:35px auto;color:#203445}h1{border-bottom:3px solid #246c9b;padding-bottom:8px}table{width:100%;border-collapse:collapse}th,td{border-bottom:1px solid #dae3e8;padding:9px 12px;text-align:left}th{background:#edf5f9}td.submitted{color:#13754a;font-weight:bold}td.missing{color:#b34b3a;font-weight:bold}small{color:#667682}</style>
    <h1>寮生会新聞 Vol."""+str(volume)+"""｜提出状況</h1><p>提出 """+str(audit["submitted"])+"""/"""+str(len(FORM_DEPARTMENTS))+"""部署・総回答 """+str(audit["total_responses"])+""" 件（最新のものを採用）</p><table><thead><tr><th>部署</th><th>状態</th><th>最新の完了時刻</th><th>文字数</th></tr></thead><tbody>"""+"".join(status_rows)+"""</tbody></table><p><small>内容の承認や掲載可否は編集担当者が確認してください。</small></p></html>"""
    (out_root / "提出状況.html").write_text(status_html, encoding="utf-8")
    env = Environment(loader=FileSystemLoader(ROOT / "templates"), autoescape=select_autoescape(["html"]))
    template = env.get_template("newspaper.html")
    overflows = {}
    output_paths = {}
    for design in designs:
        dest = out_root / design
        dest.mkdir(parents=True, exist_ok=True)
        base_name = f"寮生会新聞Vol{volume}_{TEMPLATES[design]}"
        html_path = dest / (base_name + ".html")
        pdf_path = dest / (base_name + ".pdf")
        png_path = dest / (base_name + ".png")
        html = template.render(design=design, design_label=TEMPLATES[design], volume=volume,
                               issue=edition, articles=articles, reports=reports, leader=leader)
        html_path.write_text(html, encoding="utf-8")
        overflows[design] = browser_render(html_path, pdf_path, png_path, publish)
        output_paths[design] = {"html": str(html_path), "pdf": str(pdf_path) if pdf_path.exists() else None,
                                "png": str(png_path)}

    if any(overflows.values()):
        audit["status"] = "OVERFLOW: PDFを生成せずHTML/PNGで確認"
        audit["layout_overflows"] = overflows
        write_json(out_root / "提出状況.json", audit)
        raise InputProblem("本文が紙面からはみ出しています: " + json.dumps(overflows, ensure_ascii=False))
    audit["status"] = "READY_FOR_REVIEW: 発行前の内容確認が必要" if publish else "DRAFT: 校正前"
    audit["generated_files"] = output_paths
    write_json(out_root / "提出状況.json", audit)
    if (out_root / "エラー.txt").exists():
        (out_root / "エラー.txt").unlink()
    summary = f"Vol.{volume}：有効回答{audit['total_responses']}件・提出{audit['submitted']}/{len(FORM_DEPARTMENTS)}部署・未提出{len(audit['missing'])}部署"
    print("[正常終了]", summary)
    print("[重複]", [(d["department"], d["count"]) for d in audit["duplicates"]])
    for design, paths in output_paths.items():
        print(f"[{TEMPLATES[design]}] PDF → {paths['pdf']}")
    return audit


def main():
    p = argparse.ArgumentParser(description="寮生会新聞：Forms回答を①新聞風／④現行改良版に反映する")
    p.add_argument("--input", type=Path, help="Microsoft Forms のExcel (.xlsx) またはUTF-8のCSV")
    p.add_argument("--volume", type=int, default=5, help="対象の号数（既定5）")
    p.add_argument("--edition", type=Path, default=ROOT / "edition.json", help="記事・お知らせ編集用JSON")
    p.add_argument("--design", choices=["classic", "blue", "both"], default="both")
    p.add_argument("--watch", action="store_true", help="OneDriveで同期されたExcelの更新を監視する")
    p.add_argument("--interval", type=float, default=5, help="監視する秒間隔")
    p.add_argument("--publish", action="store_true", help="未提出や原稿未確認を厳密チェックする")
    p.add_argument("--demo", action="store_true", help="架空の回答データで動作確認")
    p.add_argument("--max-chars", type=int, default=95, help="部署別の文字数の目安。実際の収まりはブラウザ判定")
    p.add_argument("--out", type=Path, help="保存先。省略時output/VolX")
    a = p.parse_args()
    if not a.input and not a.demo:
        p.error("--input または --demo を指定してください")
    if a.demo and a.publish:
        p.error("デモ原稿を発行モードにはできません")
    if a.demo and a.watch:
        p.error("--demo と --watch は併用できません")
    file = (ROOT / "samples" / "動作確認用_架空回答.xlsx") if a.demo else a.input.expanduser().resolve()
    edition = a.edition.expanduser().resolve()
    out = a.out.expanduser().resolve() if a.out else ROOT / "output" / f"Vol{a.volume}"
    designs = ["classic", "blue"] if a.design == "both" else [a.design]

    def execute():
        try:
            return compile_once(file, edition, out, a.volume, designs, a.publish, a.max_chars)
        except (InputProblem, OSError, json.JSONDecodeError) as ex:
            out.mkdir(parents=True, exist_ok=True)
            message = str(ex)
            (out / "エラー.txt").write_text(message, encoding="utf-8")
            # エラー後に古い成功PDFを最新結果と誤認しないよう削除
            for design in designs:
                d = out / design
                if d.exists():
                    for stale_pdf in d.glob("寮生会新聞Vol*_*.pdf"):
                        stale_pdf.unlink()
            print("[生成中止]", message, file=sys.stderr)
            return None

    if not a.watch:
        if execute() is None:
            raise SystemExit(2)
        return
    print(f"[監視開始] {file}（OneDriveと同期されると新聞を再生成。終了Ctrl+C）")
    last_stamp = None
    while True:
        try:
            # ファイルがロック中の場合も再試行する。異常時は古いPDFを配布しない。
            stamp = (file.stat().st_mtime_ns, file.stat().st_size, edition.stat().st_mtime_ns)
            if stamp != last_stamp:
                if execute() is not None:
                    last_stamp = stamp
            time.sleep(max(1, a.interval))
        except (FileNotFoundError, PermissionError):
            print("[監視中] ファイルの同期を待っています")
            time.sleep(max(1, a.interval))
        except KeyboardInterrupt:
            print("[監視停止]")
            return

if __name__ == "__main__":
    main()
