# -*- coding: utf-8 -*-
"""
生成《生物數據統計分析專案工作交接手冊》Word 文件 (.docx)
注重：
  1. 統計方法挑選的原因、邏輯與檢定注意事項
  2. 深入淺出，以非統計專業一般人也能完全理解的白話說明
  3. 人性化、研究團隊交接的口吻，避免生硬 AI 八股文
  4. 內嵌 Seaborn 高解析度箱形圖與完整結果表格
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    """設定表格單元格背景顏色"""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    """設定單元格邊距 (單位: dxa, 20 dxa = 1 pt)"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_callout_box(doc, title, text, bg_color="F2F5F9", border_color="1B365D"):
    """建立類似 Note / Alert 的左側粗邊框重點提示方框"""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Inches(6.5)
    
    cell = table.cell(0, 0)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=140, bottom=140, left=180, right=140)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:left w:val="single" w:sz="36" w:space="0" w:color="{border_color}"/>'
        f'  <w:top w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'  <w:bottom w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.2
    
    run_title = p.add_run(f"📌 {title}\n")
    run_title.font.name = '微軟正黑體'
    run_title.font.size = Pt(10.5)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    run_text = p.add_run(text)
    run_text.font.name = '微軟正黑體'
    run_text.font.size = Pt(9.5)
    run_text.font.color.rgb = RGBColor(0x2C, 0x3E, 0x50)
    
    # 增加段後間距
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(4)

def format_cell_text(cell, text, bold=False, color=(0x33, 0x33, 0x33), size=9.5, align=WD_ALIGN_PARAGRAPH.LEFT):
    """格式化表格單元格內的文字"""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    run.font.name = '微軟正黑體'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color)

def build_document():
    doc = Document()
    base_dir = "/Users/tenyacute/Desktop/bio"
    
    # 版面設定 (標準邊界 1 英吋)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.different_first_page_header_footer = False
        
        # 頁首頁尾
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        frun = fp.add_run("統計分析專案工作交接手冊 | 內部資料")
        frun.font.name = '微軟正黑體'
        frun.font.size = Pt(8.5)
        frun.font.color.rgb = RGBColor(0x95, 0xA5, 0xA6)
    
    # 全域樣式設定
    normal_style = doc.styles['Normal']
    normal_style.font.name = '微軟正黑體'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor(0x2C, 0x3E, 0x50)
    normal_style.paragraph_format.line_spacing = 1.35
    normal_style.paragraph_format.space_after = Pt(6)

    # -------------------------------------------------------------
    # 封面 / 標頭
    # -------------------------------------------------------------
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(6)
    r_t = p_title.add_run("統計分析與資料視覺化 專案交接手冊")
    r_t.font.name = '微軟正黑體'
    r_t.font.size = Pt(22)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(18)
    r_s = p_sub.add_run("無母數統計檢定實務應用、Seaborn 箱形圖繪製規範與白話結果判讀指南")
    r_s.font.name = '微軟正黑體'
    r_s.font.size = Pt(12)
    r_s.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    # 專案中繼資訊表
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.columns[0].width = Inches(1.8)
    meta_table.columns[1].width = Inches(4.7)
    
    meta_info = [
        ("交接專案", "生醫組織標記、運動員薪資與生理特徵無母數統計分析"),
        ("交接對象", "接手資料分析同仁、研究助理、專案主管及跨領域協作夥伴"),
        ("涵蓋資料集", "Mann whitney.xlsx、KW_Pro_salary.xlsx、KW_Blood type.xlsx"),
        ("核心統計方法", "曼-惠特尼 U 檢定 (Mann-Whitney U)、克魯斯卡爾-沃利斯檢定 (Kruskal-Wallis H)"),
        ("程式與產出檔案", "run_statistical_analysis.py、statistical_results_summary.csv、4張高清圖表")
    ]
    
    for i, (k, v) in enumerate(meta_info):
        c0 = meta_table.cell(i, 0)
        c1 = meta_table.cell(i, 1)
        set_cell_background(c0, "EBF3FA")
        set_cell_background(c1, "FDFDFD")
        set_cell_margins(c0, top=70, bottom=70, left=100, right=100)
        set_cell_margins(c1, top=70, bottom=70, left=100, right=100)
        format_cell_text(c0, k, bold=True, color=(0x1B, 0x36, 0x5D), size=9.5)
        format_cell_text(c1, v, bold=False, color=(0x33, 0x33, 0x33), size=9.5)
    
    doc.add_paragraph() # 空白間隔

    # -------------------------------------------------------------
    # 前言
    # -------------------------------------------------------------
    add_callout_box(
        doc,
        "交接說明：這份手冊寫給誰看？",
        "這份手冊是為了讓接手的同仁（無論你過去有沒有修過高等統計學）都能在 10 分鐘內完全搞懂這批資料是怎麼算的、為什麼要這樣算，以及對外報告或寫論文時該怎麼講。\n"
        "我們在處理真實世界（特別是生醫或社會經濟）數據時，最常犯的錯誤就是『不管三七二十一直接套 t 檢定或 ANOVA 算平均值』。本手冊會著重交代為什麼這次分析『非用無母數方法不可』，並教你如何看懂箱形圖與向他人清楚解說結果。"
    )

    # -------------------------------------------------------------
    # 第一章：統計方法的挑選與核心觀念
    # -------------------------------------------------------------
    h1 = doc.add_heading("第一章、為什麼這次分析「必須」使用無母數統計？", level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(6)
    for r in h1.runs:
        r.font.name = '微軟正黑體'
        r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph()
    p.add_run("很多人在學校學統計，最熟悉的就是「獨立樣本 t 檢定」和「單因子變異數分析 (ANOVA)」。這類方法叫做")
    r_bold = p.add_run("「母數統計 (Parametric Statistics)」")
    r_bold.bold = True
    p.add_run("，它們好用，但有一個非常嚴苛的前提：")
    r_crit = p.add_run("你的資料必須符合『常態分佈（鐘形曲線）』，而且樣本最好足夠大（通常 N > 30）。")
    r_crit.bold = True
    r_crit.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)

    p2 = doc.add_paragraph()
    p2.add_run("然而，在我們這次接手的這三個資料集中，這種假設完全崩潰了：")
    
    bullet_items = [
        ("理由一：樣本數太少，常態性根本不可信", "在 `Mann whitney.xlsx` 中，乳癌組只有 8 筆、肺腺癌組只有 9 筆。在樣本數只有個位數的情況下，任何常態分佈檢定（如 Shapiro-Wilk 檢定）的檢定力都極低，貿然使用 t 檢定極易產生誤判。"),
        ("理由二：極端值拉扯，平均值會說謊", "在 `KW_Pro_salary.xlsx` 中，大谷翔平年薪 7,000 萬美元，而 WNBA 選手年薪約 20 萬美元，兩者相差 350 倍！在薪資這種極度右偏的數據中，少數天價選手會把整隊的「平均值」大幅拉高。這時候看平均值完全沒有代表性，必須看「中位數」與「等級秩」。"),
        ("理由三：無母數檢定是保護研究結論的防波堤", "無母數統計（Non-parametric tests）不假設母體分佈型態，因此即使資料有極端值、離群值或偏斜嚴重，它做出的統計推論依然穩健可靠。")
    ]
    for b_title, b_desc in bullet_items:
        bp = doc.add_paragraph(style='List Bullet')
        bp.paragraph_format.space_after = Pt(4)
        r_bt = bp.add_run(b_title + "：")
        r_bt.bold = True
        bp.add_run(b_desc)

    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(6)
    p3.add_run("那麼，無母數統計到底在算什麼？用一句話白話解釋：")
    p3_quote = doc.add_paragraph()
    p3_quote.paragraph_format.left_indent = Inches(0.4)
    r_quote = p3_quote.add_run("「它不看原始數字相差多少，而是把所有人從小到大排排隊，看哪一組的排隊名次（Rank，等級秩）整體排在比較前面！」")
    r_quote.bold = True
    r_quote.font.color.rgb = RGBColor(0x29, 0x80, 0xB9)
    p3_quote.add_run(" 就算第一名拿到了 1 億美元，他在隊伍裡也只是『第 1 名』，不會因為他的天文數字把整組的名次結構給扭曲。這就是無母數檢定之所以強大抗干擾的核心原理。")

    # 母數與無母數對照表
    comp_table = doc.add_table(rows=3, cols=4)
    comp_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    comp_table.columns[0].width = Inches(1.5)
    comp_table.columns[1].width = Inches(1.6)
    comp_table.columns[2].width = Inches(1.8)
    comp_table.columns[3].width = Inches(1.6)
    
    headers = ["分析比較情境", "傳統母數方法 (需常態)", "無母數方法 (本次採用)", "本次對應檔案"]
    for j, h in enumerate(headers):
        c = comp_table.cell(0, j)
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        format_cell_text(c, h, bold=True, color=(0xFF, 0xFF, 0xFF), size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        
    rows_data = [
        ("兩組獨立樣本比較\n(如：乳癌 vs 肺腺癌)", "獨立樣本 t 檢定\n(Independent t-test)", "曼-惠特尼 U 檢定\n(Mann-Whitney U Test)", "Mann whitney.xlsx\n(OUTPUT(1).pdf)"),
        ("三組以上獨立樣本\n(如：4大聯盟、4種血型)", "單因子變異數分析\n(One-way ANOVA)", "克魯斯卡爾-沃利斯檢定\n(Kruskal-Wallis Test)", "KW_Pro_salary.xlsx\nKW_Blood type.xlsx")
    ]
    for i, r_data in enumerate(rows_data):
        bg = "F9FBFD" if i % 2 == 0 else "FFFFFF"
        for j, val in enumerate(r_data):
            c = comp_table.cell(i+1, j)
            set_cell_background(c, bg)
            set_cell_margins(c, 60, 60, 80, 80)
            align = WD_ALIGN_PARAGRAPH.CENTER if j in [1, 2, 3] else WD_ALIGN_PARAGRAPH.LEFT
            format_cell_text(c, val, bold=(j==2), color=(0x1B, 0x36, 0x5D) if j==2 else (0x33, 0x33, 0x33), size=9, align=align)

    doc.add_paragraph()

    # -------------------------------------------------------------
    # 第二章：三大數據集分析結果與白話判讀
    # -------------------------------------------------------------
    h2 = doc.add_heading("第二章、三大數據集實證結果與白話報告", level=1)
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(6)
    for r in h2.runs:
        r.font.name = '微軟正黑體'
        r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    # ---------------- 案例 1 ----------------
    h2_1 = doc.add_heading("2.1 乳癌與肺腺癌組織 HER2 表現量（Mann-Whitney U 檢定）", level=2)
    for r in h2_1.runs:
        r.font.name = '微軟正黑體'
        r.font.color.rgb = RGBColor(0x2C, 0x3E, 0x50)

    p_case1 = doc.add_paragraph()
    p_case1.add_run("• 研究背景：HER2 是一種促進細胞生長的原致癌基因蛋白。在臨床上，HER2 過量表現是乳癌標靶治療（如赫賽汀 Trastuzumab）的重要生物標記。本分析比較乳癌組（BC, n=8）與肺腺癌組（LUAD, n=9）組織中 HER2 的相對定量表現量。\n"
                    "• 數據統計指標：\n"
                    "   - 乳癌組 (BC)：中位數 446.50（平均 434.50 ± 202.05，範圍 156 ~ 666）\n"
                    "   - 肺腺癌組 (LUAD)：中位數 118.00（平均 128.44 ± 25.62，範圍 111 ~ 187）\n"
                    "• 統計檢定值：Mann-Whitney U = 1.0, Wilcoxon W = 46.0, Z = -3.372, p = 0.0009 (< 0.001***)，效應值 r = 0.818。")

    add_callout_box(
        doc,
        "向非專業主管/聽眾白話解讀 2.1：",
        "「我們比較了 8 位乳癌患者和 9 位肺腺癌患者的腫瘤組織，結果發現乳癌組織的 HER2 蛋白質含量非常顯著地高於肺腺癌（p < 0.001）。\n"
        "肺腺癌組的數值全部擠在 111 到 187 的極低區間；而乳癌組的數值幾乎全部在 250 以上，最高甚至達到 666。統計檢定顯示這種差異是由隨機運氣造成的機率小於千分之一，且效應值 r = 0.82 屬於極強效應，證明 HER2 在這兩種癌症中的表現具有顯著的特異性。」"
    )

    # 插入圖 1
    p1_img = os.path.join(base_dir, "boxplot_1_her2_expression.png")
    if os.path.exists(p1_img):
        doc.add_paragraph()
        p_img1 = doc.add_paragraph()
        p_img1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_img1 = p_img1.add_run()
        run_img1.add_picture(p1_img, width=Inches(5.2))
        p_cap1 = doc.add_paragraph()
        p_cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap1.paragraph_format.space_after = Pt(12)
        r_c1 = p_cap1.add_run("圖 1：乳癌與肺腺癌組織 HER2 表現量之 Seaborn 箱形圖（附個別樣本分佈與顯著性標註）")
        r_c1.font.size = Pt(9)
        r_c1.font.italic = True
        r_c1.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)

    # ---------------- 案例 2 ----------------
    h2_2 = doc.add_heading("2.2 四大職業運動聯盟年薪比較（Kruskal-Wallis 檢定與事後分析）", level=2)
    for r in h2_2.runs:
        r.font.name = '微軟正黑體'
        r.font.color.rgb = RGBColor(0x2C, 0x3E, 0x50)

    p_case2 = doc.add_paragraph()
    p_case2.add_run("• 研究背景：本分析取樣美職四大聯賽（NFL 美式足球、NBA 籃球、MLB 棒球、WNBA 女子籃球）各前 10 名頂尖選手的年薪（共 40 筆），探討跨聯盟間的待遇差異。\n"
                    "• 數據統計指標：\n"
                    "   - NFL (n=10)：中位數 5,275 萬美元（平均 5,246 萬美元，最高 6,000 萬美元）\n"
                    "   - NBA (n=10)：中位數 4,662 萬美元（平均 4,546 萬美元，最高 5,576 萬美元）\n"
                    "   - MLB (n=10)：中位數 3,850 萬美元（平均 4,237 萬美元，最高大谷翔平 7,000 萬美元）\n"
                    "   - WNBA (n=10)：中位數 20.35 萬美元（平均 19.99 萬美元，最高 25.2 萬美元）\n"
                    "• 整體檢定：Kruskal-Wallis H = 27.611, df = 3, p < 0.001***，效應值 η² = 0.684。\n"
                    "• 事後兩兩檢定（採用 Bonferroni 校正，避免偽陽性）：\n"
                    "   - WNBA vs NFL / NBA / MLB：校正後 p 值皆小於 0.001***（存在巨大性別與聯盟收益鴻溝）。\n"
                    "   - NFL vs MLB：校正後 p = 0.030*（NFL 頂尖四分衛年薪顯著高於棒球頂尖群）。\n"
                    "   - NFL vs NBA (p = 0.153) 與 NBA vs MLB (p = 1.000)：未達統計顯著差異。")

    add_callout_box(
        doc,
        "向非專業主管/聽眾白話解讀 2.2：",
        "「四大聯盟的頂尖薪資存在壓倒性的顯著不平等。WNBA 女籃前 10 名選手的『平均年薪約 20 萬美元』，而男子三大聯盟前 10 名都是『4,000 萬到 5,000 萬美元起跳』，整整差了 200 到 250 倍。\n"
        "另外在男子三聯盟中，美式足球（NFL）主力四分衛的年薪集中度與頂標又比棒球（MLB）更為優渥且顯著高出一截。這份數據以無母數檢定證實了極端值並不影響整體聯盟位階的排定。」"
    )

    # 插入圖 2
    p2_img = os.path.join(base_dir, "boxplot_2_pro_salary.png")
    if os.path.exists(p2_img):
        doc.add_paragraph()
        p_img2 = doc.add_paragraph()
        p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_img2 = p_img2.add_run()
        run_img2.add_picture(p2_img, width=Inches(5.4))
        p_cap2 = doc.add_paragraph()
        p_cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap2.paragraph_format.space_after = Pt(12)
        r_c2 = p_cap2.add_run("圖 2：四大職業運動聯盟頂尖年薪之 Seaborn 箱形圖（附事後檢定統計方塊）")
        r_c2.font.size = Pt(9)
        r_c2.font.italic = True
        r_c2.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)

    # ---------------- 案例 3 ----------------
    h2_3 = doc.add_heading("2.3 血型與身高關聯分析（Kruskal-Wallis 檢定）", level=2)
    for r in h2_3.runs:
        r.font.name = '微軟正黑體'
        r.font.color.rgb = RGBColor(0x2C, 0x3E, 0x50)

    p_case3 = doc.add_paragraph()
    p_case3.add_run("• 研究背景：探討一般坊間傳聞「某種血型的人發育比較好或長得比較高」是否具有科學根據。收集 44 位成人的 ABO 血型與身高（cm）數據。\n"
                    "• 數據統計指標：\n"
                    "   - A 型 (n=12)：平均 174.00 cm，中位數 173.50 cm（158 ~ 185 cm）\n"
                    "   - B 型 (n=10)：平均 171.50 cm，中位數 173.00 cm（157 ~ 184 cm）\n"
                    "   - O 型 (n=9)：平均 170.44 cm，中位數 170.00 cm（154 ~ 187 cm）\n"
                    "   - AB 型 (n=13)：平均 169.92 cm，中位數 172.00 cm（154 ~ 185 cm）\n"
                    "• 檢定結果：Kruskal-Wallis H = 1.507, df = 3, p = 0.6806 (> 0.05, 標記為 ns, not significant)，效應值 η² = 0.000。")

    add_callout_box(
        doc,
        "向非專業主管/聽眾白話解讀 2.3（如何漂亮解讀『不顯著』）：",
        "「在統計上，『不顯著（p = 0.681）』不是失敗，而是一個非常明確的科學結論！\n"
        "這代表：『血型跟身高完全沒有任何關聯』。各血型的平均身高都在 170 ~ 174 公分之間，各組內部都有高有矮，彼此之間的微小差異純粹是隨機抽樣的正常波動。因此，不需要進行任何事後檢定，可以放心地在報告中寫下：無證據支持血型會影響身高。」"
    )

    # 插入圖 3
    p3_img = os.path.join(base_dir, "boxplot_3_height_blood_type.png")
    if os.path.exists(p3_img):
        doc.add_paragraph()
        p_img3 = doc.add_paragraph()
        p_img3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_img3 = p_img3.add_run()
        run_img3.add_picture(p3_img, width=Inches(5.2))
        p_cap3 = doc.add_paragraph()
        p_cap3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap3.paragraph_format.space_after = Pt(12)
        r_c3 = p_cap3.add_run("圖 3：不同血型族群之身高分佈 Seaborn 箱形圖（四組中位數一致，統計未達顯著）")
        r_c3.font.size = Pt(9)
        r_c3.font.italic = True
        r_c3.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)

    # -------------------------------------------------------------
    # 第三章：圖表解讀指南——教接手同仁看懂 Seaborn Box Plot
    # -------------------------------------------------------------
    h3 = doc.add_heading("第三章、視覺化指南：5 分鐘看懂 Seaborn 箱形圖", level=1)
    h3.paragraph_format.space_before = Pt(14)
    h3.paragraph_format.space_after = Pt(6)
    for r in h3.runs:
        r.font.name = '微軟正黑體'
        r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_guide = doc.add_paragraph()
    p_guide.add_run("傳統的長條圖（Bar Chart）只會畫一個平均值高度，完全隱藏了真實資料的分佈與離群值，在現代頂尖期刊（如 Nature, Cell, Lancet）中已被嚴格限制使用。我們本次產出的箱形圖（Box Plot）搭配了散布點（Stripplot），包含了以下 5 個重要解讀元件：")

    box_elements = [
        ("1. 中間的彩色盒子 (The Box)", "代表「中間 50% 人的落點」（四分位距 IQR，從 25% 的 Q1 到 75% 的 Q3）。盒子越扁，表示該組人的數值越集中；盒子越長，表示差異越大。"),
        ("2. 盒子正中央的粗黑橫線 (Median)", "這是「中位數」（正中間那個人）。如果粗黑線偏向盒子底部，代表多數人偏低（右偏分佈）；如果偏向頂部，代表多數人偏高。"),
        ("3. 黃色菱形記號 (Mean, ♦)", "我們特別額外加上了「平均值記號」。若黃色菱形和粗黑橫線重疊，代表分佈對稱；若黃色菱形明顯高於黑線（如大谷翔平讓 MLB 平均值高於中位數），一眼就能看出資料受到極端值的強力拉扯！"),
        ("4. 上下延伸的黑鬚 (Whiskers)", "延伸到非離群值的最大值與最小值（通常為 1.5 倍 IQR 範圍），代表絕大多數樣本的合理邊界。"),
        ("5. 盒子上的浮動黑點 (Stripplot Jitter)", "每一個點都是一個真實的病人或球員！透過加上隨機微小左右平移（Jitter），讓讀者一眼看清樣本到底是只有 8 個點（如乳癌）還是很多點，兼具誠實度與科學透明度。")
    ]
    for e_title, e_desc in box_elements:
        ep = doc.add_paragraph(style='List Bullet')
        ep.paragraph_format.space_after = Pt(4)
        r_et = ep.add_run(e_title + "：")
        r_et.bold = True
        ep.add_run(e_desc)

    # 插入綜合面板大圖
    p4_img = os.path.join(base_dir, "boxplot_combined_panel.png")
    if os.path.exists(p4_img):
        doc.add_paragraph()
        p_img4 = doc.add_paragraph()
        p_img4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_img4 = p_img4.add_run()
        run_img4.add_picture(p4_img, width=Inches(6.4))
        p_cap4 = doc.add_paragraph()
        p_cap4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap4.paragraph_format.space_after = Pt(12)
        r_c4 = p_cap4.add_run("圖 4：生物醫學與數據統計分析 Box Plot 彙整面板（三合一發表級寬面板）")
        r_c4.font.size = Pt(9)
        r_c4.font.italic = True
        r_c4.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)

    # -------------------------------------------------------------
    # 第四章：技術交接與操作重現手冊
    # -------------------------------------------------------------
    h4 = doc.add_heading("第四章、技術交接：如何在一分鐘內重跑與維護程式", level=1)
    h4.paragraph_format.space_before = Pt(14)
    h4.paragraph_format.space_after = Pt(6)
    for r in h4.runs:
        r.font.name = '微軟正黑體'
        r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_tech = doc.add_paragraph()
    p_tech.add_run("為了確保未來的研究員能 100% 重現所有統計數據與圖表，本專案已模組化封裝。所有檔案結構與指令說明如下：")

    # 檔案目錄表
    file_table = doc.add_table(rows=8, cols=3)
    file_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    file_table.columns[0].width = Inches(2.2)
    file_table.columns[1].width = Inches(1.3)
    file_table.columns[2].width = Inches(3.0)
    
    f_headers = ["檔案名稱", "檔案類型", "用途說明"]
    for j, h in enumerate(f_headers):
        c = file_table.cell(0, j)
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 70, 70, 80, 80)
        format_cell_text(c, h, bold=True, color=(0xFF, 0xFF, 0xFF), size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        
    f_rows = [
        ("Mann whitney.xlsx", "Excel 原始資料", "乳癌與肺腺癌組織 HER2 表現量原始數據 (N=17)"),
        ("KW_Pro_salary.xlsx", "Excel 原始資料", "四大職業運動選手年薪數據 (N=40)"),
        ("KW_Blood type.xlsx", "Excel 原始資料", "成人血型與身高原始數據 (N=44)"),
        ("run_statistical_analysis.py", "Python 主程式", "一鍵執行全體統計檢定、事後比較與 Seaborn 出圖"),
        ("statistical_results_summary.csv", "CSV 統計報表", "整合所有組別之平均值、中位數、IQR、檢定量與 p 值"),
        ("boxplot_1 ~ 3.png", "PNG 圖片檔", "三張獨立 300 DPI 高畫質箱形圖 (附顯著性標示)"),
        ("boxplot_combined_panel.png", "PNG 圖片檔", "1x3 橫向寬幅三合一綜合展示大圖")
    ]
    for i, r_data in enumerate(f_rows):
        bg = "F9FBFD" if i % 2 == 0 else "FFFFFF"
        for j, val in enumerate(r_data):
            c = file_table.cell(i+1, j)
            set_cell_background(c, bg)
            set_cell_margins(c, 50, 50, 80, 80)
            align = WD_ALIGN_PARAGRAPH.CENTER if j == 1 else WD_ALIGN_PARAGRAPH.LEFT
            format_cell_text(c, val, bold=(j==0), color=(0x1B, 0x36, 0x5D) if j==0 else (0x33, 0x33, 0x33), size=8.5, align=align)

    doc.add_paragraph()

    # 指令說明
    p_cmd = doc.add_paragraph()
    p_cmd.add_run("【一鍵重跑指令】\n").bold = True
    p_cmd.add_run("接手同仁只要在 Mac 終端機（Terminal）進入專案目錄，並使用配置好的 Python 虛擬環境執行即可：")
    
    p_code = doc.add_paragraph()
    p_code.paragraph_format.left_indent = Inches(0.3)
    set_cell_background(doc.add_table(rows=1, cols=1).cell(0,0), "2D3436")
    # 代替 code box:
    t_code = doc.tables[-1]
    t_code.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_code.columns[0].width = Inches(6.5)
    c_code = t_code.cell(0, 0)
    set_cell_background(c_code, "2D3436")
    set_cell_margins(c_code, 100, 100, 150, 150)
    p_cin = c_code.paragraphs[0]
    p_cin.paragraph_format.space_before = Pt(0)
    p_cin.paragraph_format.space_after = Pt(0)
    r_c = p_cin.add_run("cd /Users/tenyacute/Desktop/bio\n/Users/tenyacute/Desktop/test(bio)/.venv/bin/python3 run_statistical_analysis.py")
    r_c.font.name = 'Courier New'
    r_c.font.size = Pt(9.5)
    r_c.font.color.rgb = RGBColor(0x00, 0xFF, 0xCC)
    r_c.bold = True

    doc.add_paragraph()

    # -------------------------------------------------------------
    # 第五章：常見疑問與避坑指南 (Q&A)
    # -------------------------------------------------------------
    h5 = doc.add_heading("第五章、口試與對外報告避坑指南（同仁問答集）", level=1)
    h5.paragraph_format.space_before = Pt(14)
    h5.paragraph_format.space_after = Pt(6)
    for r in h5.runs:
        r.font.name = '微軟正黑體'
        r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    qa_list = [
        ("Q1：如果審查委員或長官問『為什麼不用 t 檢定或單因子 ANOVA』，怎麼防禦？",
         "答：回答關鍵三點：\n"
         "1.「樣本數考量」：乳癌與肺腺癌組樣本數僅 8 與 9 筆，小樣本無法可靠驗證常態分佈假設。\n"
         "2.「抗偏態與抗極端值」：運動員年薪橫跨 20 萬至 7,000 萬美元，高度右偏，違反變異數同質性與常態性。\n"
         "3.「研究嚴謹度」：使用無母數統計得出的結論更保守、更穩健，絕不會因為換了檢定方法就讓結論翻盤。"),
        ("Q2：SPSS 報表上的 U 值跟 Python 算出來的 U 值差在哪？有算錯嗎？",
         "答：完全沒有算錯！Mann-Whitney 檢定中，兩組的 U 值總和固定等於 n1 × n2（本例為 8 × 9 = 72）。Python Scipy 預設回傳的是 U1 = 71.0，而 SPSS 習慣回傳較小值 U = min(U1, U2) = 1.000。兩者在數學上本質相同，算出的 Z 值（-3.372）與 p 值（0.0009）是百分之百完全吻合的。手冊與腳本中已統一轉換為標準報告格式 U = 1.0。"),
        ("Q3：寫英文論文或報告時，Method 與 Result 段落該如何撰寫？",
         "答：可直接參考以下標準學術範本：\n"
         "• Methods: \"Continuous variables were expressed as median with interquartile range (IQR) and mean ± standard deviation (SD). Differences between two independent groups were evaluated using the Mann-Whitney U test. Comparisons across three or more groups were performed using the Kruskal-Wallis H test, followed by pairwise post-hoc comparisons with Bonferroni correction. A two-tailed p-value < 0.05 was considered statistically significant. All analyses and visualizations were conducted using Python (Seaborn and SciPy).\"\n"
         "• Results (HER2): \"HER2 expression levels were significantly higher in the breast cancer group (median: 446.50, IQR: 343.25) than in the LUAD group (median: 118.00, IQR: 13.00; Mann-Whitney U = 1.0, Z = -3.372, p < 0.001, effect size r = 0.818).\"")
    ]
    for q, a in qa_list:
        add_callout_box(doc, q, a, bg_color="F8F9FA", border_color="34495E")

    # 結尾簽署
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(20)
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_sign = p_sign.add_run("交接製作：生物數據統計分析專案小組\n日期：2026 年 10 月 15 日")
    r_sign.font.size = Pt(9.5)
    r_sign.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)

    # 存檔
    output_docx_path = os.path.join(base_dir, "生物醫學數據統計分析交接手冊.docx")
    doc.save(output_docx_path)
    print(f"成功生成 Word 交接手冊：{output_docx_path}")

if __name__ == "__main__":
    build_document()
