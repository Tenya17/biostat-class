# 生物統計實習與數據分析專案 (Biostatistics & Biomedical Data Analysis)

![Python](https://img.shields.io/badge/Python-3.9+-3776AB?style=flat&logo=python&logoColor=white)
![Data Analysis](https://img.shields.io/badge/Data%20Analysis-Pandas%20%7C%20SciPy-blue)
![Visualization](https://img.shields.io/badge/Visualization-Matplotlib%20%7C%20Seaborn-orange)
![R](https://img.shields.io/badge/R-4.4+-276DC3?style=flat&logo=r&logoColor=white)
![Status](https://img.shields.io/badge/Status-Completed-success)

本儲存庫收錄生物統計實習專案之臨床醫學數據分析、無母數統計檢定（Non-parametric Tests）、流行病學田野調查報告，以及 Python 自動化分析與報表生成腳本。

---

## 📂 專案目錄結構 (Repository Structure)

```plaintext
biostat-class/
├── 1029生統/                               # 新藥臨床試驗 Kaplan-Meier 存活分析專案
│   ├── Survival test example.xlsx          # 新藥試驗原始追蹤數據集 (N=12)
│   ├── km_survival_analysis.R              # R 存活分析、Log-Rank 檢定與出版級繪圖腳本
│   ├── kaplan_meier_curve.png              # 300 DPI 出版級存活曲線圖 (含 Number at Risk 表格)
│   ├── kaplan_meier_curve.pdf              # 向量圖檔 (高解析論文發表用)
│   ├── kaplan_meier_curve_with_CI.png      # 附 95% 信賴區間帶 (CI) 存活曲線圖
│   ├── kaplan_meier_interactive.html       # 現代化互動式存活分析網頁報告
│   ├── survival_summary_table.csv          # 詳細存活生命表 (Life Table)
│   └── README.md                           # 存活分析專案說明與生統解讀
│
├── 1022 bio/                               # 流行病學調查專案
│   ├── 橘郡異常事件綜合調查報告.docx/.pdf    # 整合版流行病學與水質污染調查報告
│   ├── 橘郡異常事件調查報告.md/.docx        # 居民流行病學調查分析
│   ├── 橘郡異常事件調查報告（第二份）.md/.docx # 深入水源與病理關聯報告
│   ├── 1022生統數據.docx                    # SPSS 統計原始報表
│   └── 複本 調查情報20260331.xlsx           # 現場個案情報原始資料（N=669）
│
├── bio/                                    # 生物醫學統計分析與無母數檢定模組
│   ├── run_statistical_analysis.py         # 核心統計計算與圖表生成腳本
│   ├── generate_handover_docx.py           # 自動化生成 Word 交接手冊腳本
│   ├── 生物醫學數據統計分析交接手冊.docx/.pdf # 完整統計推論與結果交接文件
│   ├── statistical_results_summary.csv     # 統計檢定結果彙整表（Mean, SD, Median, p-value）
│   ├── boxplot_combined_panel.png          # 多組箱型圖整合面板
│   ├── boxplot_*.png                       # 各分析主題獨立圖表
│   ├── OUTPUT(1).pdf / OUTPUT(2).pdf       # 分析執行紀錄與輸出
│   └── KW_*.xlsx / Mann whitney.xlsx       # 各組檢定原始數據集
│
├── test(bio)/                              # 糖尿病（DM）臨床專案與擴充驗證
│   ├── DM.xls                              # 糖尿病臨床指標原始數據集
│   ├── DM資料分析專案交接與統計方法手冊.pdf   # 專案交接技術手冊
│   ├── analysis.py / create_handover_docx.py
│   └── dm_boxplot_*.png                    # 臨床指標分組箱型圖
│
└── .gitignore                              # 系統暫存與虛擬環境排除清單
```

---

## 🔬 主要研究模組與統計方法 (Methodologies)

### 1. 臨床生物標誌物比較 (HER2 Expression in Cancer)
- **檢定方法**：曼-惠特尼 U 檢定（Mann-Whitney U Test，無母數雙獨立樣本）
- **分析標的**：乳癌（Breast Cancer, BC）與肺腺癌（LUAD）組織之 HER2 表現量
- **統計發現**：乳癌組表現量（Median = 446.50）極顯著高於肺腺癌組（Median = 118.00），$U = 1.0, Z = -3.372, p = 0.0009$（$p < 0.001$），效應值 $r = 0.818$（大效應）。

### 2. 多組獨立樣本無母數分析 (Kruskal-Wallis Test)
- **職業運動員年薪比較**：
  - 比較 NFL、NBA、MLB 與 WNBA 四大職業運動之年薪分佈。
  - 檢定結果達極顯著差異（$H = 27.611, df = 3, p = 4.38 \times 10^{-6}, \eta^2 = 0.684$）。
  - 搭配 Dunn-Bonferroni 事後檢定（Post-hoc test）釐清各聯盟間顯著差異。
- **血型與身高相關性檢定**：
  - 分析 A、B、O、AB 四種血型個體身高是否具差異。
  - 檢定結果未達顯著（$H = 1.507, p = 0.681$），支持血型與身高無相關性之虛無假說。

### 3. 流行病學異常事件調查 (Orange County Incident Investigation)
- **樣本規模**：$N = 669$ 名居民個案完整田野與臨床紀錄。
- **分析內容**：
  - 性別、年齡分佈等基本人口學描述統計。
  - 飲用水來源（深井水、淺井水、自來水）與患病嚴重程度之卡方檢定與風險比（Odds Ratio / Relative Risk）。
  - 疾病致死率與空間傳播特徵推論。

### 4. 糖尿病臨床指標關聯分析 (DM Clinical Analysis)
- 探索糖尿病患者之性別、年齡組別在各項生化與臨床指標上的中位數與四分位距（IQR）分佈。

### 5. 新藥臨床試驗存活分析 (Kaplan-Meier Survival Analysis & Log-Rank Test)
- **分析標的**：新藥治療組（Group 1, $n=6$）與對照組（Group 2, $n=6$）之整體存活期（Overall Survival）。
- **分析方法**：
  - **Kaplan-Meier 乘積極限法**：估算各時間點之存活機率 $\hat{S}(t)$ 與 Greenwood 標準誤。
  - **Log-Rank 檢定**：檢定兩組存活曲線差異（$\chi^2 = 2.30, p = 0.1291$）。
  - **Cox 等比例風險回歸**：計算風險比 Hazard Ratio（$HR = 0.356$，降低 64.4% 事件風險）。
- **核心發現**：新藥組中位存活時間為 6.50 個月，明顯優於對照組的 5.25 個月（延長 +1.25 個月），且研究截止時仍有 44.4% 受試者存活設限。

---

## 📊 統計結果摘要 (Summary of Statistical Findings)

| 分析主題 | 檢定方法 | 分組 | 樣本數 $N$ | 中位數 (IQR) | 統計量 | $p$ 值 | 顯著性與效應值 |
| :--- | :--- | :--- | :---: | :---: | :---: | :---: | :---: |
| **HER2 表現量** | Mann-Whitney U | 乳癌 vs 肺腺癌 | 8 vs 9 | 446.50 vs 118.00 | $U = 1.0$ | **0.0009** | 極顯著 ($r = 0.818$) |
| **職業選手薪資** | Kruskal-Wallis | NFL / NBA / MLB / WNBA | 各 10 | \$52.8M / \$46.6M / \$38.5M / \$0.2M | $H = 27.611$ | **< 0.001** | 極顯著 ($\eta^2 = 0.684$) |
| **血型與身高** | Kruskal-Wallis | A / B / O / AB | 12 / 10 / 9 / 13 | 173.5 / 173.0 / 170.0 / 172.0 (cm) | $H = 1.507$ | 0.6806 | 不顯著 ($\eta^2 = 0.000$) |
| **新藥存活分析** | Kaplan-Meier / Log-Rank | 新藥組 vs 對照組 | 6 vs 6 | 中位 6.50m vs 5.25m | $\chi^2 = 2.30$ | **0.1291** | HR = 0.356 (小樣本未達顯著) |

---

## 🛠️ 自動化腳本執行說明 (How to Run)

### 環境需求
- Python 3.9 或更高版本
- 必要相依套件：
  ```bash
  pip install numpy pandas scipy matplotlib seaborn python-docx openpyxl xlrd
  ```

### 執行統計分析與繪圖
```bash
cd "bio"
python run_statistical_analysis.py
```
執行後將自動重新計算各項無母數檢定數值，並輸出高品質箱型圖及 `statistical_results_summary.csv`。

### 生成交接手冊 Word 報告
```bash
python generate_handover_docx.py
```
自動抓取最新統計檢定數據與分析圖檔，排版組裝成標準格式之交付手冊（`.docx`）。

---

## 📜 授權說明 (License)
本專案為學術實習與教學研究用途，資料與腳本僅供學術交流參考。
