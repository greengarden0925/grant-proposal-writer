#!/usr/bin/env python3
"""
研究計畫書合規檢查（grant-proposal-writer）

用法：
  python compliance_check.py proposal.docx --config proposal-config.md [--refs refs.ris] [--out report.md]

檢查項目：
  1. 實際頁數（以 LibreOffice 轉 PDF 後計算）與章節頁數
  2. 必填章節是否存在
  3. 殘留標記：【待查證】【需補文獻】【請填入】
  4. 參考文獻每筆是否有 DOI / PMID / URL / ISBN
  5. 正文 author-date 引用與參考文獻清單雙向比對（盡力而為，需人工確認）
  6. （選用）RIS 檔每筆是否有 DOI 或 PMID

結束碼：0 = 無硬性問題；1 = 有硬性問題（超頁、缺必填章節、殘留待查證標記、文獻缺識別碼）
需要：python-docx、pypdf；頁數檢查另需 LibreOffice（soffice）。
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata

try:
    import docx  # python-docx
except ImportError:
    sys.exit("缺少 python-docx：pip install python-docx")

MARKERS = ["【待查證】", "【需補文獻】", "【請填入】"]
HARD_MARKERS = ["【待查證】", "【需補文獻】"]
ID_PATTERNS = {
    "DOI": re.compile(r"10\.\d{4,9}/[^\s]+", re.I),
    "PMID": re.compile(r"PMID[:：]?\s*\d{5,9}", re.I),
    "URL": re.compile(r"https?://\S+", re.I),
    "ISBN": re.compile(r"ISBN[:：]?\s*[\d\-Xx]{10,17}", re.I),
}
# 英文 author-date：(Smith, 2020) (Smith & Lee, 2020) (Smith et al., 2020a; Chen, 2019)
CITE_EN = re.compile(
    r"([A-Z][A-Za-z'\-]+)(?:\s+(?:&|and)\s+[A-Z][A-Za-z'\-]+)?(?:\s+et al\.)?,?\s+((?:19|20)\d{2}[a-z]?)"
)
# 中文 author-date：（陳等人，2020）（陳、林，2019）
CITE_ZH = re.compile(r"([一-鿿]{1,4}?)(?:[、與和][一-鿿]{1,4}?)?(?:等人|等)?[，,]\s*((?:19|20)\d{2}[a-z]?)")


def norm(s):
    return unicodedata.normalize("NFKC", s or "").strip().lower()


def read_config(path):
    if not path:
        return {}
    text = open(path, encoding="utf-8").read()
    m = re.search(r"```compliance\s*\n(.*?)```", text, re.S)
    if not m:
        print("⚠ config 中找不到 ```compliance 區塊，只做不需參數的檢查。", file=sys.stderr)
        return {}
    try:
        return json.loads(m.group(1))
    except json.JSONDecodeError as e:
        sys.exit(f"config 的 compliance 區塊不是合法 JSON：{e}")


def heading_level(par):
    name = (par.style.name or "") if par.style is not None else ""
    m = re.match(r"(?:Heading|標題)\s*(\d+)", name, re.I)
    if m:
        return int(m.group(1))
    if name.lower() == "title":
        return 0
    return None


def load_docx(path):
    d = docx.Document(path)
    paras = []
    for p in d.paragraphs:
        paras.append({"text": p.text, "level": heading_level(p)})
    # 表格內文字也納入標記與引用檢查
    table_text = []
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                table_text.append(cell.text)
    return paras, table_text


def docx_to_pdf_pages(path):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None, "找不到 LibreOffice，無法計算實際頁數（請在 Word 中自行確認頁數）。"
    try:
        from pypdf import PdfReader
    except ImportError:
        return None, "缺少 pypdf，無法計算實際頁數：pip install pypdf"
    out = tempfile.mkdtemp(prefix="gpw_")
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", out, path],
            check=True, capture_output=True, timeout=180,
        )
        pdf = os.path.join(out, os.path.splitext(os.path.basename(path))[0] + ".pdf")
        reader = PdfReader(pdf)
        pages = [norm(pg.extract_text() or "") for pg in reader.pages]
        return pages, None
    except Exception as e:  # noqa: BLE001
        return None, f"轉檔失敗，無法計算實際頁數：{e}"
    finally:
        shutil.rmtree(out, ignore_errors=True)


def find_page(pages, text, start=0):
    key = re.sub(r"\s+", "", norm(text))[:40]
    if not key:
        return None
    for i in range(start, len(pages)):
        if key in re.sub(r"\s+", "", pages[i]):
            return i
    return None


def section_pages(paras, pages):
    """回傳 [(標題, 層級, 起始頁(1-based), 結束頁, 頁數)]"""
    heads = [(p["text"], p["level"]) for p in paras if p["level"] is not None and p["text"].strip()]
    located = []
    cursor = 0
    for text, lvl in heads:
        pg = find_page(pages, text, cursor)
        if pg is not None:
            cursor = pg
        located.append((text, lvl, pg))
    result = []
    for i, (text, lvl, pg) in enumerate(located):
        if pg is None:
            result.append((text, lvl, None, None, None))
            continue
        end = len(pages) - 1
        for text2, lvl2, pg2 in located[i + 1:]:
            if lvl2 is not None and lvl2 <= lvl and pg2 is not None:
                # 下一節若從新頁開頭，本節結束於前一頁；保守起見計入共用頁
                end = pg2
                break
        result.append((text, lvl, pg + 1, end + 1, end - pg + 1))
    return result


def split_references(paras, ref_heads):
    ref_heads = [norm(h) for h in ref_heads]
    idx, lvl = None, None
    for i, p in enumerate(paras):
        if p["level"] is not None and any(h in norm(p["text"]) for h in ref_heads):
            idx, lvl = i, p["level"]
    if idx is None:
        return paras, [], False
    refs = []
    for p in paras[idx + 1:]:
        if p["level"] is not None and p["level"] <= lvl:
            break
        if p["text"].strip():
            refs.append(p["text"].strip())
    body = paras[:idx]
    return body, refs, True


def citation_keys(text):
    keys = set()
    for m in re.finditer(r"[（(]([^（）()]{3,300})[)）]", text):
        inner = m.group(1)
        for a, y in CITE_EN.findall(inner):
            keys.add((a.lower(), y[:4]))
        for a, y in CITE_ZH.findall(inner):
            keys.add((a, y[:4]))
    # 敘述式引用：Smith et al. (2020)
    for m in re.finditer(r"([A-Z][A-Za-z'\-]+)(?:\s+(?:&|and)\s+[A-Z][A-Za-z'\-]+)?(?:\s+et al\.)?\s*[（(]((?:19|20)\d{2})[a-z]?[)）]", text):
        keys.add((m.group(1).lower(), m.group(2)))
    return keys


def ref_key(ref):
    year = re.search(r"((?:19|20)\d{2})", ref)
    y = year.group(1) if year else None
    m = re.match(r"\s*(?:\[\d+\]\s*)?([A-Z][A-Za-z'\-]+)", ref)
    if m:
        return (m.group(1).lower(), y)
    m = re.match(r"\s*([一-鿿]{1,4})", ref)
    if m:
        return (m.group(1), y)
    return (None, y)


def same_author(a, b):
    if a == b:
        return True
    if a and b and re.match(r"[\u4e00-\u9fff]", a) and re.match(r"[\u4e00-\u9fff]", b):
        return a.startswith(b) or b.startswith(a)
    return False


def check_ris(path):
    issues, n = [], 0
    rec = {}
    for line in open(path, encoding="utf-8", errors="replace"):
        tag = line[:2]
        val = line[6:].strip() if len(line) > 6 else ""
        if tag == "TY":
            rec, n = {}, n + 1
        if tag in ("DO", "AN", "TI", "T1", "AU", "A1", "PY", "Y1", "UR", "SN", "C1", "M1"):
            rec.setdefault(tag, []).append(val)
        if tag == "ER":
            has_doi = any(ID_PATTERNS["DOI"].search(v) for v in rec.get("DO", []))
            has_pmid = any(re.fullmatch(r"\d{5,9}", v) for v in rec.get("AN", []) + rec.get("C1", []) + rec.get("M1", []))
            if not (has_doi or has_pmid):
                title = (rec.get("TI") or rec.get("T1") or ["(無標題)"])[0]
                issues.append(title[:80])
    return n, issues


def main():
    ap = argparse.ArgumentParser(description="研究計畫書合規檢查")
    ap.add_argument("docx")
    ap.add_argument("--config")
    ap.add_argument("--refs", help="RIS 檔（選用）")
    ap.add_argument("--out", help="報告輸出路徑（.md）")
    args = ap.parse_args()

    cfg = read_config(args.config)
    paras, table_text = load_docx(args.docx)
    all_text = "\n".join(p["text"] for p in paras) + "\n" + "\n".join(table_text)
    hard, report = [], []

    def add(line):
        report.append(line)

    add(f"# 合規檢查報告：{os.path.basename(args.docx)}\n")

    # 1. 頁數
    add("## 1. 頁數")
    pages, err = docx_to_pdf_pages(args.docx)
    if err:
        add(f"- ⚠ {err}")
    else:
        total = len(pages)
        lim = cfg.get("total_page_limit")
        status = "" if lim is None else ("✅" if total <= lim else "❌")
        add(f"- 全文實際頁數：**{total}** 頁" + (f"（上限 {lim}）{status}" if lim is not None else ""))
        if lim is not None and total > lim:
            hard.append(f"全文超頁：{total} > {lim}")
        secs = section_pages(paras, pages)
        if secs:
            add("\n| 章節 | 起始頁 | 結束頁 | 頁數 | 上限 | 狀態 |\n|---|---|---|---|---|---|")
            limits = {norm(k): v for k, v in (cfg.get("section_page_limits") or {}).items()}
            matched_limits = set()
            for text, lvl, s, e, n in secs:
                lim_s = next((v for k, v in limits.items() if k in norm(text)), None)
                if lim_s is not None:
                    matched_limits.add(next(k for k in limits if k in norm(text)))
                if s is None:
                    add(f"| {'　' * (lvl or 0)}{text[:40]} | 找不到 | | | {lim_s or ''} | ⚠ |")
                    continue
                st = ""
                if lim_s is not None:
                    st = "✅" if n <= lim_s else "❌"
                    if n > lim_s:
                        hard.append(f"章節超頁：{text[:30]} {n} > {lim_s}")
                add(f"| {'　' * (lvl or 0)}{text[:40]} | {s} | {e} | {n} | {lim_s or ''} | {st} |")
            for k in set(limits) - matched_limits:
                add(f"- ⚠ config 設定了「{k}」的頁數上限，但找不到對應的標題。")
            add("\n註：章節頁數以標題所在頁計算，相鄰章節共用的頁會同時計入兩節，屬保守估計。標題須使用 Word 標題樣式（Heading 1/2/3）才能辨識。")

    # 2. 必填章節
    add("\n## 2. 必填章節")
    req = cfg.get("required_sections") or []
    heads = [norm(p["text"]) for p in paras if p["level"] is not None]
    if not req:
        add("- （config 未設定 required_sections）")
    for r in req:
        if any(norm(r) in h for h in heads):
            add(f"- ✅ {r}")
        elif norm(r) in norm(all_text):
            add(f"- ⚠ {r}：出現在內文，但不是標題樣式")
        else:
            add(f"- ❌ {r}：找不到")
            hard.append(f"缺少必填章節：{r}")

    # 3. 殘留標記
    add("\n## 3. 殘留標記")
    any_marker = False
    for mk in MARKERS:
        hits = [p["text"][:60] for p in paras if mk in p["text"]] + [t[:60] for t in table_text if mk in t]
        if hits:
            any_marker = True
            icon = "❌" if mk in HARD_MARKERS else "⚠"
            add(f"- {icon} {mk}：{len(hits)} 處")
            for h in hits[:10]:
                add(f"  - {h}…")
            if mk in HARD_MARKERS:
                hard.append(f"殘留 {mk} {len(hits)} 處")
    if not any_marker:
        add("- ✅ 無殘留標記")

    # 4. 參考文獻識別碼
    add("\n## 4. 參考文獻識別碼")
    body, refs, found = split_references(paras, cfg.get("reference_heading") or ["參考文獻", "References"])
    if not found:
        add("- ❌ 找不到參考文獻章節（需以標題樣式設定）")
        hard.append("找不到參考文獻章節")
    else:
        missing = [r for r in refs if not any(p.search(r) for p in ID_PATTERNS.values())]
        add(f"- 參考文獻共 {len(refs)} 筆；缺 DOI/PMID/URL/ISBN：{len(missing)} 筆")
        for r in missing[:20]:
            add(f"  - ❌ {r[:100]}")
        if missing:
            hard.append(f"{len(missing)} 筆參考文獻缺識別碼")

    # 5. 引用雙向比對
    add("\n## 5. 正文引用與參考文獻比對（盡力而為，請人工確認）")
    if found:
        body_text = "\n".join(p["text"] for p in body)
        cited = citation_keys(body_text)
        ref_keys = [ref_key(r) for r in refs]
        ref_set = {k for k in ref_keys if k[0]}
        not_in_refs = sorted(k for k in cited if not any(same_author(k[0], a) and k[1] == y for a, y in ref_set))
        not_cited = [r for r, k in zip(refs, ref_keys) if k[0] and not any(same_author(k[0], a) and k[1] == y for a, y in cited)]
        add(f"- 正文辨識出 {len(cited)} 個 作者-年份 引用")
        add(f"- 正文有引用、清單找不到：{len(not_in_refs)}")
        for a, y in not_in_refs[:30]:
            add(f"  - ⚠ {a} ({y})")
        add(f"- 清單有、正文未引用：{len(not_cited)}")
        for r in not_cited[:30]:
            add(f"  - ⚠ {r[:100]}")
        add("- 註：數字編號引用格式（如 [1]）不在此比對範圍內。")

    # 6. RIS
    if args.refs:
        add("\n## 6. RIS 檔檢查")
        n, issues = check_ris(args.refs)
        add(f"- RIS 共 {n} 筆；缺 DOI 或 PMID：{len(issues)} 筆")
        for t in issues[:20]:
            add(f"  - ⚠ {t}")

    add("\n## 結論")
    if hard:
        add("❌ 有硬性問題，修正後再進入審查：")
        for h in hard:
            add(f"- {h}")
    else:
        add("✅ 無硬性問題。仍請完成 review-panel.md 的人工合規清單。")

    out = "\n".join(report)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(out)
    print(out)
    sys.exit(1 if hard else 0)


if __name__ == "__main__":
    main()
