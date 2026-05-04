"""
テニス選手（コーチ）の経費扱い 調査・文書生成スクリプト

調査テーマ：
  テニスクラブ等に雇用されたテニスコーチ・選手が
  業務上支出した経費（交通費・用具費・試合参加費等）の
  労働法・税法上の取り扱いと、未払い経費の請求方法
"""

from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree

FONT_NAME = "MS明朝"
FONT_SIZE = Pt(10.5)


def fw(text):
    """半角ASCII文字をすべて全角に変換（e内容証明対応）。"""
    result = []
    for ch in text:
        code = ord(ch)
        if 0x21 <= code <= 0x7E:
            result.append(chr(code + 0xFEE0))
        elif code == 0x20:
            result.append('　')
        else:
            result.append(ch)
    return ''.join(result)


def _apply_font_xml(rPr):
    tag = qn("w:rFonts")
    rFonts = rPr.find(tag)
    if rFonts is None:
        rFonts = etree.SubElement(rPr, tag)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), FONT_NAME)


def set_font(run):
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE
    run.font.bold = False
    run.font.italic = False
    run.font.underline = False
    run.font.color.rgb = RGBColor(0, 0, 0)
    rPr = run._r.get_or_add_rPr()
    _apply_font_xml(rPr)
    for tag in (qn("w:sz"), qn("w:szCs")):
        el = rPr.find(tag)
        if el is None:
            el = etree.SubElement(rPr, tag)
        el.set(qn("w:val"), "21")


def configure_doc_defaults(doc):
    normal = doc.styles["Normal"]
    normal.font.name = FONT_NAME
    normal.font.size = FONT_SIZE
    normal.font.bold = False
    normal.font.italic = False
    normal.font.underline = False
    nfPr = normal.element.get_or_add_rPr()
    _apply_font_xml(nfPr)
    doc_elm = doc.element
    styles_elm = doc_elm.find(qn("w:styles"))
    if styles_elm is None:
        return
    docDefaults = styles_elm.find(qn("w:docDefaults"))
    if docDefaults is None:
        docDefaults = etree.SubElement(styles_elm, qn("w:docDefaults"))
    rPrDefault = docDefaults.find(qn("w:rPrDefault"))
    if rPrDefault is None:
        rPrDefault = etree.SubElement(docDefaults, qn("w:rPrDefault"))
    rPr = rPrDefault.find(qn("w:rPr"))
    if rPr is None:
        rPr = etree.SubElement(rPrDefault, qn("w:rPr"))
    _apply_font_xml(rPr)
    for rFonts_el in doc_elm.iter(qn("w:rFonts")):
        for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            key = qn(attr)
            if key in rFonts_el.attrib:
                del rFonts_el.attrib[key]


def add_para(doc, text, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent_cm=0):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    if first_indent_cm:
        pf.first_line_indent = Cm(first_indent_cm)
    run = p.add_run(text)
    set_font(run)
    return p


# ============================================================
# 調査内容（テニス選手・コーチの経費扱い）
# ============================================================
RESEARCH = """
【テニス選手（コーチ）の経費扱い 調査まとめ】

■ 1. 労働者として雇用されている場合の原則

テニスクラブや企業に雇用されたテニスコーチ・選手は「労働者」であり、
業務遂行上必要な経費は使用者が負担すべきものが原則です（民法485条・
労働契約法3条）。

使用者が業務上必要な経費を労働者に立て替えさせた場合、
これは「立替金」として返還請求できます。

■ 2. 主な経費の種類と法的根拠

(1) 交通費（業務のための移動）
  → 会社が業務指示をした試合・遠征・出張に伴う交通費は
    業務上の経費であり、使用者が負担義務を負う。
    明示的な取り決めがなくても、業務命令がある以上
    立替費用の返還を求めうる（民法650条1項）。

(2) 試合参加費（大会エントリー費）
  → 会社の業務として参加させられた試合・大会の
    エントリー費用は業務経費に該当し得る。
    会社が参加を指示・推奨した場合は特に強い。

(3) 用具費（ラケット・シューズ・ウェア等）
  → 業務上必須の道具であれば経費計上が認められる場合がある。
    ただし、個人としても使用する物品は按分が問題となる。
    会社指定ユニフォームや消耗品は会社負担が原則。

(4) 宿泊費（遠征・合宿）
  → 会社の業務上の遠征・合宿に伴う宿泊費は
    業務経費として使用者負担が原則。

■ 3. 所得税法上の「特定支出控除」（参考）

給与所得者が業務上必要な支出をした場合、
所得税法第57条の2に基づく「特定支出控除」を
確定申告で適用できる（給与所得控除の1/2超の部分）。

特定支出の対象（同法57条の2第2項）：
  ・通勤費
  ・転居費
  ・研修費
  ・資格取得費
  ・帰宅旅費
  ・勤務必要経費（図書費・衣服費・交際費等、65万円限度）

テニスコーチの場合、コーチング資格取得費・
専門書籍費・テニスウェア（業務限定）等が該当し得る。

■ 4. 未払い経費の請求方法

(1) 内容証明郵便による請求
  → 未払い経費の具体的金額・費目を列挙し、
    支払期限を定めて使用者に請求する。

(2) 労働基準監督署への申告
  → 経費の未払いが賃金の不払いに相当する場合
    （歩合給に含む等の合意があった場合を除く）
    労基署への申告が可能。

(3) 労働審判・民事訴訟
  → 60万円以下の場合は少額訴訟（1回期日）も利用可能。
    立替金の証拠（領収書・交通系ICカード履歴等）を保全する。

■ 5. 証拠として有効なもの

  ・領収書（正本）
  ・交通系ICカードの利用履歴（Suica・PiTaPa等）
  ・銀行振込明細
  ・大会エントリー確認メール
  ・会社からの業務指示（LINE・メール等）
  ・業務日誌・タイムカード

■ 6. 消滅時効

  立替金の返還請求権：5年（民法166条1項1号）
  賃金未払い（賃金に該当する場合）：3年（労働基準法115条）

■ 参考法令

  ・民法第650条第1項（受任者の費用償還請求権）
  ・労働基準法第15条（労働条件の明示）
  ・労働基準法第115条（時効）
  ・所得税法第57条の2（特定支出控除）
  ・雇用保険法施行規則第17条

"""


def main():
    print(RESEARCH)

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.bottom_margin = Cm(7.0)
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    for hdr_para in section.header.paragraphs:
        for run in hdr_para.runs:
            run.text = ""
    for ftr_para in section.footer.paragraphs:
        for run in ftr_para.runs:
            run.text = ""

    configure_doc_defaults(doc)

    C = WD_ALIGN_PARAGRAPH.CENTER
    R = WD_ALIGN_PARAGRAPH.RIGHT
    L = WD_ALIGN_PARAGRAPH.LEFT

    add_para(doc, "テニス選手（コーチ）の経費扱い　調査報告書", align=C)
    add_para(doc, "")
    add_para(doc, fw("令和8年5月4日"), align=R)
    add_para(doc, "")

    sections_data = [
        ("１．労働者として雇用されている場合の原則", [
            "テニスクラブや企業に雇用されたテニスコーチ・選手は「労働者」であり、"
            "業務遂行上必要な経費は使用者が負担すべきものが原則です（民法第485条・"
            "労働契約法第3条）。",
            "　使用者が業務上必要な経費を労働者に立て替えさせた場合、これは「立替金」"
            "として返還請求できます（民法第650条第1項）。",
        ]),
        ("２．主な経費の種類と法的根拠", [
            fw("（1）交通費（業務のための移動）"),
            "　会社が業務指示をした試合・遠征・出張に伴う交通費は業務上の経費であり、"
            "使用者が負担義務を負います。業務命令がある以上、立替費用の返還を求めるこ"
            "とができます（民法第650条第1項）。",
            fw("（2）試合参加費（大会エントリー費）"),
            "　会社の業務として参加を指示・推奨された試合・大会のエントリー費用は"
            "業務経費に該当します。",
            fw("（3）用具費（ラケット・シューズ・ウェア等）"),
            "　業務上必須の道具は経費計上が認められます。会社指定のユニフォームや"
            "業務専用消耗品は使用者負担が原則です。",
            fw("（4）宿泊費（遠征・合宿）"),
            "　会社の業務上の遠征・合宿に伴う宿泊費は使用者負担が原則です。",
        ]),
        ("３．所得税法上の特定支出控除（参考）", [
            "　給与所得者が業務上必要な支出をした場合、所得税法第57条の2に基づく"
            "「特定支出控除」を確定申告で適用できます（給与所得控除の2分の1超の部分）。",
            "　テニスコーチの場合、コーチング資格取得費・専門書籍費・"
            "業務限定のテニスウェア等が対象となり得ます。",
        ]),
        ("４．未払い経費の請求方法", [
            fw("（1）内容証明郵便による請求"),
            "　未払い経費の具体的金額・費目を列挙し、支払期限を定めて使用者に"
            "請求します。",
            fw("（2）労働基準監督署への申告"),
            "　経費の未払いが賃金の不払いに相当する場合、労働基準監督署への"
            "申告が可能です。",
            fw("（3）少額訴訟・労働審判"),
            fw("　60万円以下の場合は少額訴訟（1回期日）が利用できます。"),
        ]),
        ("５．証拠として有効なもの", [
            "　・領収書（正本）",
            "　・交通系ICカードの利用履歴（Suica・PiTaPa等）",
            "　・銀行振込明細",
            "　・大会エントリー確認メール",
            "　・会社からの業務指示（LINE・メール等）",
            "　・業務日誌・タイムカード",
        ]),
        (fw("６．消滅時効"), [
            "　立替金の返還請求権：5年（民法第166条第1項第1号）",
            "　賃金に該当する場合：3年（労働基準法第115条）",
        ]),
    ]

    for heading, paras in sections_data:
        add_para(doc, heading)
        for para_text in paras:
            add_para(doc, para_text, first_indent_cm=0)
        add_para(doc, "")

    add_para(doc, "以上", align=R)

    out = "テニス選手_経費扱い_調査報告書.docx"
    doc.save(out)
    print(f"\n保存: {out}")

    bad = []
    check_doc = Document(out)
    for i, para in enumerate(check_doc.paragraphs):
        for run in para.runs:
            rPr = run._r.find(qn("w:rPr"))
            if rPr is None:
                continue
            rFonts = rPr.find(qn("w:rFonts"))
            if rFonts is None:
                continue
            for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                val = rFonts.get(qn(attr))
                if val and val != FONT_NAME:
                    bad.append(f"  段落{i+1} run='{run.text[:10]}' {attr}={val}")
    if bad:
        print("⚠ 異フォント検出:")
        for b in bad:
            print(b)
    else:
        print("✓ 全フォント MS明朝 で統一されています")


if __name__ == "__main__":
    main()
