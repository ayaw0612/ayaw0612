"""
e内容証明（電子内容証明）文書生成スクリプト
日本郵便 e内容証明サービス書式規定に完全準拠

修正ポイント（よくあるエラー対策）:
- MS明朝 に完全統一（ascii / hAnsi / eastAsia / cs すべて指定）
  → 半角数字・ハイフン・アルファベット(LINE等)を Century 等で
    レンダリングするのを防ぐ。これが「文字装飾エラー」の最大原因。
- 太字(bold)・下線(underline)・斜体(italic)は一切なし
- ヘッダー・フッター・表・テキストボックス・図形なし
- 余白: 上・左・右 1.5cm / 下 7.0cm（認証印スペース）
- フォントサイズ: 10.5pt 統一（サイズ混在もエラー原因）
- スタイルデフォルト・テーマフォントも MS明朝 に上書き
"""

from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree


FONT_NAME = "MS明朝"
FONT_SIZE = Pt(10.5)


def fw(text):
    """半角ASCII文字をすべて全角に変換する。
    e内容証明は半角文字（数字・英字・ハイフン等）を一切受け付けないため必須。
    0x21-0x7E の ASCII 印字可能文字 → 0xFF01-0xFF5E（全角対応文字）
    半角スペース → 全角スペース（U+3000）
    """
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
    """w:rFonts の ascii / hAnsi / eastAsia / cs をすべて MS明朝 に設定する。
    python-docx の run.font.name だけでは ascii/hAnsi しか変わらない場合があるため
    XML を直接操作して確実に統一する。"""
    tag = qn("w:rFonts")
    rFonts = rPr.find(tag)
    if rFonts is None:
        rFonts = etree.SubElement(rPr, tag)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), FONT_NAME)


def set_font(run):
    """Run の全フォント属性を統一する（装飾なし）。"""
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE
    run.font.bold = False
    run.font.italic = False
    run.font.underline = False
    run.font.color.rgb = RGBColor(0, 0, 0)

    rPr = run._r.get_or_add_rPr()
    _apply_font_xml(rPr)

    # sz / szCs を明示（サイズ混在エラー防止）
    sz_val = "210"  # 10.5pt = 21 half-points → 210? 実際は 10.5 * 2 = 21
    # python-docx の Pt(10.5) は 133350 EMU = 10.5pt 正しい
    # w:sz は half-points 単位: 10.5pt * 2 = 21
    for tag in (qn("w:sz"), qn("w:szCs")):
        el = rPr.find(tag)
        if el is None:
            el = etree.SubElement(rPr, tag)
        el.set(qn("w:val"), "21")  # 21 half-points = 10.5pt


def configure_doc_defaults(doc):
    """ドキュメントレベルのデフォルトフォント設定を MS明朝 で統一する。
    Word はスタイルの継承でフォントを上書きするため、ここで押さえておく。"""
    # Normal スタイルのフォント設定
    normal = doc.styles["Normal"]
    normal.font.name = FONT_NAME
    normal.font.size = FONT_SIZE
    normal.font.bold = False
    normal.font.italic = False
    normal.font.underline = False

    nfPr = normal.element.get_or_add_rPr()
    _apply_font_xml(nfPr)

    # w:docDefaults > w:rPrDefault > w:rPr
    doc_elm = doc.element
    body = doc_elm.find(qn("w:body"))
    # settings ではなく styles 直下の docDefaults を操作
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

    # テーマフォント (w:theme の majorFont / minorFont) を無効化するため
    # w:rFonts に w:asciiTheme / w:hAnsiTheme が残っていれば削除
    for rFonts_el in doc_elm.iter(qn("w:rFonts")):
        for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            key = qn(attr)
            if key in rFonts_el.attrib:
                del rFonts_el.attrib[key]


def add_para(doc, text, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent_cm=0):
    """段落を追加して run のフォントを確実に設定する。"""
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    if first_indent_cm:
        pf.first_line_indent = Cm(first_indent_cm)

    if text:
        run = p.add_run(text)
        set_font(run)
    else:
        # 空段落にもフォントを設定し、高さを 10.5pt に固定
        run = p.add_run("")
        set_font(run)
    return p


def main():
    doc = Document()

    # ──────────────────────────────
    # ページ設定
    # ──────────────────────────────
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.bottom_margin = Cm(7.0)   # 認証印スペース
    # ヘッダー・フッターを確実に無効化
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    for hdr_para in section.header.paragraphs:
        for run in hdr_para.runs:
            run.text = ""
    for ftr_para in section.footer.paragraphs:
        for run in ftr_para.runs:
            run.text = ""

    configure_doc_defaults(doc)

    R = WD_ALIGN_PARAGRAPH.RIGHT
    C = WD_ALIGN_PARAGRAPH.CENTER
    L = WD_ALIGN_PARAGRAPH.LEFT

    # ──────────────────────────────
    # 文書本文
    # ──────────────────────────────

    # タイトル
    add_para(doc, "通　知　書", align=C)
    add_para(doc, "")

    # 日付
    add_para(doc, fw("令和8年4月25日"), align=R)
    add_para(doc, "")

    # 宛先
    add_para(doc, fw("〒665-0035　兵庫県宝塚市逆瀬川1-4-19-202"), align=R)
    add_para(doc, "有限会社サラジュ　代表取締役　岩崎真也　殿", align=R)
    add_para(doc, "")

    # 差出人
    add_para(doc, fw("〒663-8141　兵庫県西宮市高須町1丁目7-15-809"), align=R)
    add_para(doc, "通知人：上島　彩", align=R)
    add_para(doc, "")

    # 前文
    add_para(doc, fw(
        "私は、貴社を令和8年3月31日をもって退職いたしましたが、現在、退職後の"
        "諸手続きを進めております。"))
    add_para(doc, fw(
        "つきましては、法令に基づき貴社に交付義務がある以下の書面について、"
        "本通知受領後7日以内に送付するよう請求いたします。"))
    add_para(doc,
        "なお、本通知と入れ違いで既に発送済みのものがある場合は、重ねての請求と"
        "なる旨ご容赦ください。")
    add_para(doc, "")

    # 1項目
    add_para(doc, fw("１．労働条件通知書（または雇用契約書写し）の交付"))
    add_para(doc, fw(
        "　労働基準法第15条第1項は、使用者が労働者に対し、賃金・労働時間等の"
        "労働条件を書面で明示することを義務付けています。令和8年4月24日付の"
        "貴社回答（LINE）において「書面の作成・保管はない」との説明がありました"
        "が、同条に基づく明示義務の存否は貴社の内部管理状況にかかわらず法定の"
        "ものであり、貴社の説明の当否は関係機関が判断することです。"))
    add_para(doc,
        "　つきましては、在職中に適用された労働条件を記載した書面を速やかに"
        "交付してください。")
    add_para(doc, "")

    # 2項目
    add_para(doc, fw("２．離職票（紙）および退職証明書の交付"))
    add_para(doc, fw(
        "　雇用保険法施行規則第17条および労働基準法第22条第1項に基づき、"
        "請求します。離職票については令和8年4月3日に、退職証明書についても"
        "同日に交付を求めましたが、いずれも長期間対応されませんでした。"
        "令和8年4月25日付の貴社回答において「本日送付した」との説明がありました"
        "が、追跡番号のない普通郵便とのことであり、到達の確認が取れない状況です。"))
    add_para(doc, fw(
        "　本通知は、万一未着の場合に備えた公式の交付請求として提出するものです。"
        "未着の場合は、追跡可能な方法（レターパック等）にて速やかに再送してください。"))
    add_para(doc, "")

    # 結び
    add_para(doc, "以上", align=R)

    # ──────────────────────────────
    # 保存
    # ──────────────────────────────
    out = "内容証明郵便_上島彩.docx"
    doc.save(out)
    print(f"保存: {out}")

    # ──────────────────────────────
    # 検証: MS明朝 以外のフォントが残っていないか確認
    # ──────────────────────────────
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
