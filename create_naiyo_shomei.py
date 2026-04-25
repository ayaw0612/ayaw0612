"""
e内容証明（電子内容証明）テンプレート生成スクリプト
日本郵便 e内容証明サービス書式規定に準拠

書式規定:
- 用紙: A4縦置き・横書き
- 余白: 上・左・右 1.5cm以上、下 7.0cm以上（認証印スペース）
- フォント: MS明朝（または MS P明朝、MSゴシック、MS Pゴシック）
- 文字サイズ: 10.5pt以上
- 目安: 1ページ約1,584字（10.5pt・標準行間）
"""

from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy


FONT_NAME = "MS明朝"
FONT_SIZE = Pt(10.5)


def set_font(run, bold=False):
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE
    run.font.bold = bold
    # 日本語フォント設定（eastAsia）
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), FONT_NAME)
    rFonts.set(qn("w:ascii"), FONT_NAME)
    rFonts.set(qn("w:hAnsi"), FONT_NAME)


def add_paragraph(doc, text="", alignment=WD_ALIGN_PARAGRAPH.LEFT, bold=False, indent_cm=0):
    p = doc.add_paragraph()
    p.alignment = alignment
    if indent_cm:
        p.paragraph_format.first_line_indent = Cm(indent_cm)
    if text:
        run = p.add_run(text)
        set_font(run, bold=bold)
    else:
        # 空行でもフォントサイズを設定してスペースを確保
        run = p.add_run(" ")
        set_font(run)
    return p


def set_doc_defaults(doc):
    """ドキュメント全体のデフォルトフォントを設定"""
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = FONT_NAME
    normal.font.size = FONT_SIZE
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), FONT_NAME)


def main():
    doc = Document()

    # ページ設定: A4縦
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)

    # 余白: 上・左・右 1.5cm、下 7cm（認証印スペース）
    section.top_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.bottom_margin = Cm(7.0)

    set_doc_defaults(doc)

    # ---- 文書本文 ----

    # 宛先（右寄せ）
    add_paragraph(doc, "○○株式会社", alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    add_paragraph(doc, "代表取締役　○○　○○　殿", alignment=WD_ALIGN_PARAGRAPH.RIGHT)

    add_paragraph(doc)  # 空行

    # 日付（右寄せ）
    add_paragraph(doc, "令和　　年　　月　　日", alignment=WD_ALIGN_PARAGRAPH.RIGHT)

    add_paragraph(doc)  # 空行

    # 差出人（右寄せ）
    add_paragraph(doc, "〒○○○－○○○○", alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    add_paragraph(doc, "○○県○○市○○町○丁目○番○号", alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    add_paragraph(doc, "○○　○○　㊞", alignment=WD_ALIGN_PARAGRAPH.RIGHT)

    add_paragraph(doc)  # 空行

    # タイトル
    add_paragraph(doc, "通　知　書", alignment=WD_ALIGN_PARAGRAPH.CENTER, bold=True)

    add_paragraph(doc)  # 空行

    # 本文
    body_texts = [
        "拝啓　時下ますますご清栄のこととお慶び申し上げます。",
        "",
        "　さて、私（以下「通知人」といいます。）は、貴社との間において、令和　　年　　月　　日付けで○○契約（以下「本契約」といいます。）を締結しておりますが、下記の事由により、本契約を解除する旨を通知いたします。",
        "",
        "記",
        "",
        "１　解除事由",
        "　貴社は、本契約第○条に基づき、令和　　年　　月　　日までに○○の義務を履行すべきところ、同日を経過した現在においても、当該義務を履行しておりません。",
        "",
        "２　解除の意思表示",
        "　以上の事由により、通知人は、本書面をもって本契約を解除いたします。",
        "",
        "３　対応のお願い",
        "　本書面到達後○日以内に、○○○○○○（金　　　　円也）を下記口座へお振り込みくださいますようお願いいたします。",
        "",
        "　なお、上記期限内にご対応いただけない場合は、法的手段を講じることも検討いたしますので、ご了承ください。",
        "",
        "　　　　　　　振込先口座",
        "　　　　　　　○○銀行　○○支店",
        "　　　　　　　普通預金　口座番号　○○○○○○○",
        "　　　　　　　口座名義　○○　○○",
        "",
        "　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　以上",
    ]

    for line in body_texts:
        if line == "記":
            add_paragraph(doc, line, alignment=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
        elif line == "":
            add_paragraph(doc)
        else:
            add_paragraph(doc, line)

    # ファイル保存
    out_path = "内容証明郵便_テンプレート.docx"
    doc.save(out_path)
    print(f"保存完了: {out_path}")


if __name__ == "__main__":
    main()
