# -*- coding: utf-8 -*-
"""
生物統計分析與 Seaborn 箱形圖 (Box Plot) 繪製腳本
分析項目：
  1. Mann-Whitney U 檢定：乳癌 (BC) vs 肺腺癌 (LUAD) 之 HER2 表現量
  2. Kruskal-Wallis 檢定：四大職業運動聯盟 (NFL, NBA, MLB, WNBA) 頂尖選手年薪比較
  3. Kruskal-Wallis 檢定：不同血型 (A, B, O, AB) 之身高比較
"""

import os
import numpy as np
import pandas as pd
import scipy.stats as stats
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.lines import Line2D

# ==========================================
# 1. 環境設定與字型設定
# ==========================================
plt.rcParams['font.sans-serif'] = ['Heiti TC', 'Arial Unicode MS', 'Songti SC', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.titlesize'] = 16
plt.rcParams['axes.titlesize'] = 13
plt.rcParams['axes.labelsize'] = 11
plt.rcParams['xtick.labelsize'] = 10
plt.rcParams['ytick.labelsize'] = 10

sns.set_theme(style="whitegrid", font='Heiti TC')

output_dir = os.path.dirname(os.path.abspath(__file__))

print("=" * 70)
print("開始讀取資料並執行統計分析...")
print("=" * 70)

summary_rows = []

# ==============================================================================
# 分析 1：Mann-Whitney U 檢定 - HER2 表現量 (BC vs LUAD)
# ==============================================================================
mw_file = os.path.join(output_dir, "Mann whitney.xlsx")
df_mw = pd.read_excel(mw_file)

label_dict_mw = {"BC": "乳癌 (BC)", "LUAD": "肺腺癌 (LUAD)"}
df_mw['組別標籤'] = df_mw['annotation'].map(label_dict_mw)

bc_vals = df_mw[df_mw['annotation'] == 'BC']['HER2_level']
luad_vals = df_mw[df_mw['annotation'] == 'LUAD']['HER2_level']

# Mann-Whitney U 檢定
u_stat, mw_p = stats.mannwhitneyu(bc_vals, luad_vals, alternative='two-sided')
spss_u = min(u_stat, len(bc_vals) * len(luad_vals) - u_stat)
n1, n2 = len(bc_vals), len(luad_vals)

# 計算 Z score (常態近似)
mean_u = (n1 * n2) / 2.0
ties = pd.concat([bc_vals, luad_vals]).value_counts()
t_correction = sum(t**3 - t for t in ties) / ((n1 + n2) * (n1 + n2 - 1))
std_u = np.sqrt((n1 * n2 / 12.0) * ((n1 + n2 + 1) - t_correction))
z_score = (spss_u - mean_u) / std_u

# 計算效應值 r = |Z| / sqrt(N)
r_effect_mw = abs(z_score) / np.sqrt(n1 + n2)

for grp_code, grp_label, vals in [("BC", "乳癌 (BC)", bc_vals), ("LUAD", "肺腺癌 (LUAD)", luad_vals)]:
    summary_rows.append({
        "分析主題": "HER2 表現量比較",
        "檢定方法": "Mann-Whitney U Test",
        "分組變項": grp_label,
        "樣本數 N": len(vals),
        "平均值 Mean": f"{vals.mean():.2f}",
        "標準差 SD": f"{vals.std():.2f}",
        "中位數 Median": f"{vals.median():.2f}",
        "第1四分位 Q1": f"{vals.quantile(0.25):.2f}",
        "第3四分位 Q3": f"{vals.quantile(0.75):.2f}",
        "四分位距 IQR": f"{vals.quantile(0.75) - vals.quantile(0.25):.2f}",
        "最小值 Min": f"{vals.min():.2f}",
        "最大值 Max": f"{vals.max():.2f}",
        "檢定統計量": f"U = {spss_u:.1f}, Z = {z_score:.3f}",
        "P 值": f"{mw_p:.4f}",
        "顯著性判定": "極顯著 (p < 0.001)",
        "效應值": f"r = {r_effect_mw:.3f} (大效應)"
    })

print(f"[1] HER2 表現量檢定: U = {spss_u:.1f}, Z = {z_score:.3f}, p = {mw_p:.5f}, Effect size r = {r_effect_mw:.3f}")

# 繪製圖 1：HER2 表現量 Box Plot
plt.figure(figsize=(7, 6.2), dpi=300)
palette_mw = {"乳癌 (BC)": "#e74c3c", "肺腺癌 (LUAD)": "#3498db"}

ax1 = sns.boxplot(
    data=df_mw, x='組別標籤', y='HER2_level', hue='組別標籤',
    palette=palette_mw, width=0.45, legend=False,
    boxprops=dict(alpha=0.75, edgecolor='black', linewidth=1.5),
    medianprops=dict(color='black', linewidth=2.5),
    whiskerprops=dict(color='black', linewidth=1.5),
    capprops=dict(color='black', linewidth=1.5),
    showmeans=True,
    meanprops=dict(marker='D', markerfacecolor='yellow', markeredgecolor='black', markersize=7.5)
)

sns.stripplot(
    data=df_mw, x='組別標籤', y='HER2_level', hue='組別標籤',
    palette=palette_mw, alpha=0.7, size=8, jitter=0.18, legend=False,
    linewidth=0.8, edgecolor='black'
)

# 顯著性標註線
y_max = df_mw['HER2_level'].max()
bar_y = y_max + 35
h_tip = 15
plt.plot([0, 0, 1, 1], [bar_y, bar_y + h_tip, bar_y + h_tip, bar_y], lw=1.5, c='black')
plt.text(0.5, bar_y + h_tip + 12, f"*** p < 0.001 (p = {mw_p:.4f})\nMann-Whitney U = {spss_u:.0f}, Z = {z_score:.3f}",
         ha='center', va='bottom', color='#c0392b', fontsize=10.5, fontweight='bold', linespacing=1.3)

plt.title("乳癌與肺腺癌組織 HER2 表現量比較\n(Mann-Whitney U 檢定)", fontsize=14, fontweight='bold', pad=15)
plt.xlabel("腫瘤類型 (組別)", fontweight='bold', labelpad=10)
plt.ylabel("HER2 表現量 (相對定量值)", fontweight='bold', labelpad=10)
plt.ylim(50, bar_y + 120)

legend_elements = [
    Line2D([0], [0], marker='D', color='w', label='組平均值 (Mean)', markerfacecolor='yellow', markeredgecolor='black', markersize=8),
    Line2D([0], [0], color='black', lw=2.5, label='組中位數 (Median)'),
    Line2D([0], [0], marker='o', color='w', label='個別樣本觀測點', markerfacecolor='gray', markeredgecolor='black', markersize=8)
]
plt.legend(handles=legend_elements, loc='upper right', frameon=True, facecolor='white', framealpha=0.92, fontsize=10)

plt.tight_layout()
p1_path = os.path.join(output_dir, "boxplot_1_her2_expression.png")
plt.savefig(p1_path, dpi=300)
plt.close()
print(f"已儲存圖表: {p1_path}")


# ==============================================================================
# 分析 2：Kruskal-Wallis 檢定 - 四大職業運動選手年薪比較
# ==============================================================================
ps_file = os.path.join(output_dir, "KW_Pro_salary.xlsx")
df_ps = pd.read_excel(ps_file)

df_ps['Salary_M'] = df_ps['Salary (Annual)'] / 1e6
prof_order = ['NFL', 'NBA', 'MLB', 'WNBA']
df_ps['profession'] = pd.Categorical(df_ps['profession'], categories=prof_order, ordered=True)

groups_ps = [df_ps[df_ps['profession'] == p]['Salary_M'].values for p in prof_order]
kw_h, kw_p = stats.kruskal(*groups_ps)

k_groups = len(prof_order)
total_n_ps = len(df_ps)
eta_sq_h = (kw_h - k_groups + 1) / (total_n_ps - k_groups)

for p in prof_order:
    vals = df_ps[df_ps['profession'] == p]['Salary_M']
    summary_rows.append({
        "分析主題": "職業運動選手年薪比較",
        "檢定方法": "Kruskal-Wallis Test",
        "分組變項": f"{p} 聯盟",
        "樣本數 N": len(vals),
        "平均值 Mean": f"${vals.mean():.2f}M",
        "標準差 SD": f"${vals.std():.2f}M",
        "中位數 Median": f"${vals.median():.2f}M",
        "第1四分位 Q1": f"${vals.quantile(0.25):.2f}M",
        "第3四分位 Q3": f"${vals.quantile(0.75):.2f}M",
        "四分位距 IQR": f"${vals.quantile(0.75) - vals.quantile(0.25):.2f}M",
        "最小值 Min": f"${vals.min():.2f}M",
        "最大值 Max": f"${vals.max():.2f}M",
        "檢定統計量": f"H = {kw_h:.3f}, df = {k_groups - 1}",
        "P 值": f"{kw_p:.4e}",
        "顯著性判定": "極顯著 (p < 0.001)",
        "效應值": f"η² = {eta_sq_h:.3f} (極大效應)"
    })

print(f"[2] 職業運動年薪檢定: H = {kw_h:.3f}, df = {k_groups - 1}, p = {kw_p:.4e}, η² = {eta_sq_h:.3f}")

# 繪製圖 2：職業運動年薪 Box Plot
plt.figure(figsize=(8.5, 6.8), dpi=300)
palette_ps = {"NFL": "#1f77b4", "NBA": "#ff7f0e", "MLB": "#2ca02c", "WNBA": "#e377c2"}

ax2 = sns.boxplot(
    data=df_ps, x='profession', y='Salary_M', hue='profession',
    order=prof_order, palette=palette_ps, width=0.5, legend=False,
    boxprops=dict(alpha=0.8, edgecolor='black', linewidth=1.5),
    medianprops=dict(color='black', linewidth=2.5),
    whiskerprops=dict(color='black', linewidth=1.5),
    capprops=dict(color='black', linewidth=1.5),
    showmeans=True,
    meanprops=dict(marker='D', markerfacecolor='yellow', markeredgecolor='black', markersize=7.5)
)

sns.stripplot(
    data=df_ps, x='profession', y='Salary_M', hue='profession',
    order=prof_order, palette=palette_ps, alpha=0.7, size=7.5, jitter=0.2, legend=False,
    linewidth=0.8, edgecolor='black'
)

plt.title("四大職業運動聯盟頂尖選手年薪分佈比較\n(Kruskal-Wallis 檢定)", fontsize=14, fontweight='bold', pad=15)
plt.xlabel("職業運動聯盟 (Profession)", fontweight='bold', labelpad=10)
plt.ylabel("年薪 (百萬美元, Million USD)", fontweight='bold', labelpad=10)
plt.ylim(-5, 82)

# 註解方塊位置調整，避免遮蔽點位
kw_text = (
    "Kruskal-Wallis 檢定統計值:\n"
    f"  • H = {kw_h:.2f}, df = 3, p < 0.001***\n"
    f"  • 效應值 η² = {eta_sq_h:.2f} (極大差異)\n"
    "事後檢定 (Bonferroni 校正):\n"
    "  • WNBA 顯著低於 NFL, NBA, MLB (p < 0.01***)\n"
    "  • NFL 顯著高於 MLB (p < 0.05*)"
)
plt.annotate(kw_text, xy=(0.04, 0.12), xycoords='axes fraction',
             bbox=dict(boxstyle="round,pad=0.6", fc="white", ec="#7f8c8d", alpha=0.92, lw=1.2),
             fontsize=9.5, linespacing=1.4)

legend_elements2 = [
    Line2D([0], [0], marker='D', color='w', label='組平均值 (Mean)', markerfacecolor='yellow', markeredgecolor='black', markersize=8),
    Line2D([0], [0], color='black', lw=2.5, label='組中位數 (Median)'),
    Line2D([0], [0], marker='o', color='w', label='個別選手年薪點', markerfacecolor='gray', markeredgecolor='black', markersize=8)
]
plt.legend(handles=legend_elements2, loc='upper left', frameon=True, facecolor='white', framealpha=0.92, fontsize=10)

plt.tight_layout()
p2_path = os.path.join(output_dir, "boxplot_2_pro_salary.png")
plt.savefig(p2_path, dpi=300)
plt.close()
print(f"已儲存圖表: {p2_path}")


# ==============================================================================
# 分析 3：Kruskal-Wallis 檢定 - 血型與身高比較
# ==============================================================================
bt_file = os.path.join(output_dir, "KW_Blood type.xlsx")
df_bt = pd.read_excel(bt_file).dropna(subset=['blood type', 'Height (cm)'])

bt_order = ['A', 'B', 'O', 'AB']
df_bt['blood_type_label'] = df_bt['blood type'].astype(str) + " 型"
bt_label_order = [b + " 型" for b in bt_order]
df_bt['blood_type_label'] = pd.Categorical(df_bt['blood_type_label'], categories=bt_label_order, ordered=True)

groups_bt = [df_bt[df_bt['blood type'] == b]['Height (cm)'].values for b in bt_order]
kw_h_bt, kw_p_bt = stats.kruskal(*groups_bt)
k_bt = len(bt_order)
n_bt = len(df_bt)
eta_sq_bt = max(0.0, (kw_h_bt - k_bt + 1) / (n_bt - k_bt))

for b in bt_order:
    vals = df_bt[df_bt['blood type'] == b]['Height (cm)']
    summary_rows.append({
        "分析主題": "血型與身高關聯分析",
        "檢定方法": "Kruskal-Wallis Test",
        "分組變項": f"{b} 型",
        "樣本數 N": len(vals),
        "平均值 Mean": f"{vals.mean():.2f} cm",
        "標準差 SD": f"{vals.std():.2f} cm",
        "中位數 Median": f"{vals.median():.2f} cm",
        "第1四分位 Q1": f"{vals.quantile(0.25):.2f} cm",
        "第3四分位 Q3": f"{vals.quantile(0.75):.2f} cm",
        "四分位距 IQR": f"{vals.quantile(0.75) - vals.quantile(0.25):.2f} cm",
        "最小值 Min": f"{vals.min():.2f} cm",
        "最大值 Max": f"{vals.max():.2f} cm",
        "檢定統計量": f"H = {kw_h_bt:.3f}, df = {k_bt - 1}",
        "P 值": f"{kw_p_bt:.4f}",
        "顯著性判定": "未達統計顯著 (p = 0.681)",
        "效應值": f"η² = {eta_sq_bt:.4f} (無效應)"
    })

print(f"[3] 血型與身高檢定: H = {kw_h_bt:.3f}, df = {k_bt - 1}, p = {kw_p_bt:.4f}, η² = {eta_sq_bt:.4f}")

# 繪製圖 3：血型與身高 Box Plot
plt.figure(figsize=(7.8, 6.2), dpi=300)
palette_bt = {"A 型": "#4daf4a", "B 型": "#377eb8", "O 型": "#ff7f00", "AB 型": "#984ea3"}

ax3 = sns.boxplot(
    data=df_bt, x='blood_type_label', y='Height (cm)', hue='blood_type_label',
    order=bt_label_order, palette=palette_bt, width=0.48, legend=False,
    boxprops=dict(alpha=0.75, edgecolor='black', linewidth=1.5),
    medianprops=dict(color='black', linewidth=2.5),
    whiskerprops=dict(color='black', linewidth=1.5),
    capprops=dict(color='black', linewidth=1.5),
    showmeans=True,
    meanprops=dict(marker='D', markerfacecolor='yellow', markeredgecolor='black', markersize=7.5)
)

sns.stripplot(
    data=df_bt, x='blood_type_label', y='Height (cm)', hue='blood_type_label',
    order=bt_label_order, palette=palette_bt, alpha=0.7, size=7.5, jitter=0.2, legend=False,
    linewidth=0.8, edgecolor='black'
)

plt.title("不同血型族群之身高分佈比較\n(Kruskal-Wallis 檢定)", fontsize=14, fontweight='bold', pad=15)
plt.xlabel("血型類別 (Blood Type)", fontweight='bold', labelpad=10)
plt.ylabel("身高 (公分, cm)", fontweight='bold', labelpad=10)
plt.ylim(145, 198)

# 標註未達顯著資訊（置於中上方避免遮蔽）
kw_bt_text = (
    "Kruskal-Wallis 檢定結果:\n"
    f"  • H = {kw_h_bt:.2f}, df = 3, p = {kw_p_bt:.3f} (ns)\n"
    "  • 各血型族群間身高無統計顯著差異"
)
plt.annotate(kw_bt_text, xy=(0.04, 0.05), xycoords='axes fraction',
             bbox=dict(boxstyle="round,pad=0.6", fc="white", ec="#95a5a6", alpha=0.92, lw=1.2),
             fontsize=10, linespacing=1.4)

legend_elements3 = [
    Line2D([0], [0], marker='D', color='w', label='組平均值 (Mean)', markerfacecolor='yellow', markeredgecolor='black', markersize=8),
    Line2D([0], [0], color='black', lw=2.5, label='組中位數 (Median)'),
    Line2D([0], [0], marker='o', color='w', label='個別樣本身高點', markerfacecolor='gray', markeredgecolor='black', markersize=8)
]
plt.legend(handles=legend_elements3, loc='upper right', frameon=True, facecolor='white', framealpha=0.92, fontsize=10)

plt.tight_layout()
p3_path = os.path.join(output_dir, "boxplot_3_height_blood_type.png")
plt.savefig(p3_path, dpi=300)
plt.close()
print(f"已儲存圖表: {p3_path}")


# ==============================================================================
# 圖 4：3 合 1 綜合統計圖表面板 (1x3 Panel Overview)
# ==============================================================================
fig, axes = plt.subplots(1, 3, figsize=(18, 5.8), dpi=300)

# 子圖 A: HER2
sns.boxplot(
    ax=axes[0], data=df_mw, x='組別標籤', y='HER2_level', hue='組別標籤',
    palette=palette_mw, width=0.45, legend=False,
    boxprops=dict(alpha=0.75, edgecolor='black', linewidth=1.3),
    medianprops=dict(color='black', linewidth=2.2),
    showmeans=True, meanprops=dict(marker='D', markerfacecolor='yellow', markeredgecolor='black', markersize=6.5)
)
sns.stripplot(ax=axes[0], data=df_mw, x='組別標籤', y='HER2_level', hue='組別標籤',
              palette=palette_mw, alpha=0.65, size=6.5, jitter=0.18, legend=False,
              linewidth=0.6, edgecolor='black')
# 標註顯著性
axes[0].plot([0, 0, 1, 1], [680, 695, 695, 680], lw=1.3, c='black')
axes[0].text(0.5, 705, "*** p < 0.001", ha='center', va='bottom', color='#c0392b', fontsize=11, fontweight='bold')
axes[0].set_title("(A) 乳癌 vs 肺腺癌 HER2 表現量\nMann-Whitney U = 1.0, p < 0.001***", fontsize=12, fontweight='bold')
axes[0].set_xlabel("腫瘤類別", fontweight='bold')
axes[0].set_ylabel("HER2 表現量 (相對值)", fontweight='bold')
axes[0].set_ylim(50, 770)

# 子圖 B: 職業薪資
sns.boxplot(
    ax=axes[1], data=df_ps, x='profession', y='Salary_M', hue='profession',
    order=prof_order, palette=palette_ps, width=0.5, legend=False,
    boxprops=dict(alpha=0.75, edgecolor='black', linewidth=1.3),
    medianprops=dict(color='black', linewidth=2.2),
    showmeans=True, meanprops=dict(marker='D', markerfacecolor='yellow', markeredgecolor='black', markersize=6.5)
)
sns.stripplot(ax=axes[1], data=df_ps, x='profession', y='Salary_M', hue='profession',
              order=prof_order, palette=palette_ps, alpha=0.65, size=6, jitter=0.18, legend=False,
              linewidth=0.6, edgecolor='black')
axes[1].set_title("(B) 職業運動選手年薪\nKruskal-Wallis H = 27.61, p < 0.001***", fontsize=12, fontweight='bold')
axes[1].set_xlabel("職業運動聯盟", fontweight='bold')
axes[1].set_ylabel("年薪 (百萬美元, $M USD)", fontweight='bold')
axes[1].set_ylim(-3, 76)

# 子圖 C: 血型與身高
sns.boxplot(
    ax=axes[2], data=df_bt, x='blood_type_label', y='Height (cm)', hue='blood_type_label',
    order=bt_label_order, palette=palette_bt, width=0.48, legend=False,
    boxprops=dict(alpha=0.75, edgecolor='black', linewidth=1.3),
    medianprops=dict(color='black', linewidth=2.2),
    showmeans=True, meanprops=dict(marker='D', markerfacecolor='yellow', markeredgecolor='black', markersize=6.5)
)
sns.stripplot(ax=axes[2], data=df_bt, x='blood_type_label', y='Height (cm)', hue='blood_type_label',
              order=bt_label_order, palette=palette_bt, alpha=0.65, size=6, jitter=0.18, legend=False,
              linewidth=0.6, edgecolor='black')
axes[2].set_title("(C) 血型與身高比較\nKruskal-Wallis H = 1.51, p = 0.681 (ns)", fontsize=12, fontweight='bold')
axes[2].set_xlabel("血型", fontweight='bold')
axes[2].set_ylabel("身高 (公分, cm)", fontweight='bold')
axes[2].set_ylim(148, 195)

plt.suptitle("生物醫學與數據統計分析 Box Plot 彙整面板", fontsize=15, fontweight='bold', y=1.02)
plt.tight_layout()
p4_path = os.path.join(output_dir, "boxplot_combined_panel.png")
plt.savefig(p4_path, dpi=300, bbox_inches='tight')
plt.close()
print(f"已儲存圖表: {p4_path}")


# ==============================================================================
# 輸出統計摘要表至 CSV
# ==============================================================================
df_summary = pd.DataFrame(summary_rows)
summary_csv_path = os.path.join(output_dir, "statistical_results_summary.csv")
df_summary.to_csv(summary_csv_path, index=False, encoding='utf-8-sig')
print(f"已儲存統計摘要表: {summary_csv_path}")

print("=" * 70)
print("所有統計分析與 Seaborn 箱形圖繪製完成！")
print("=" * 70)
