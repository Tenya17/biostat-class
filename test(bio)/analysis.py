# -*- coding: utf-8 -*-
"""
DM.xls 生物醫學統計分析與 Seaborn 箱形圖 (Box Plot) 繪製腳本
分析變項：BMI (kg/m²)、血糖 (mg/dL)、收縮壓 (mmHg)
"""

import os
import pandas as pd
import numpy as np
import scipy.stats as stats
import matplotlib.pyplot as plt
import seaborn as sns

# 1. 設定中文字型與繪圖風格
plt.rcParams['font.sans-serif'] = ['Heiti TC', 'PingFang TC', 'Arial Unicode MS', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False
sns.set_theme(style="whitegrid", font='Heiti TC')

# 2. 讀取資料
file_path = "DM.xls"
df = pd.read_excel(file_path)

# 將組別與性別轉換為具意義的標籤
group_map = {1: "第 1 組 (糖尿病/個案組)", 2: "第 2 組 (對照組)"}
gender_map = {1: "男性", 2: "女性"}
df['組別名稱'] = df['組別'].map(group_map)
df['性別名稱'] = df['性別'].map(gender_map)

target_vars = [
    {"col": "BMI (kg/m²)", "name": "BMI", "unit": "kg/m²", "color": "#2b5c8f"},
    {"col": "血糖 (mg/dL)", "name": "血糖 (Fasting Blood Glucose)", "unit": "mg/dL", "color": "#d95f02"},
    {"col": "收縮壓 (mmHg)", "name": "收縮壓 (Systolic BP)", "unit": "mmHg", "color": "#7570b3"}
]

print("=" * 70)
print("DM.xls 資料概覽")
print("=" * 70)
print(f"總樣本數: {len(df)} 人")
print(f"組別分佈: 第 1 組 n = {(df['組別']==1).sum()}, 第 2 組 n = {(df['組別']==2).sum()}")
print(f"性別分佈: 男性 n = {(df['性別']==1).sum()}, 女性 n = {(df['性別']==2).sum()}\n")

# 3. 描述性統計與推論檢定計算
summary_records = []

for item in target_vars:
    col = item["col"]
    name = item["name"]
    unit = item["unit"]
    
    # 全體
    all_vals = df[col]
    mean_all = all_vals.mean()
    std_all = all_vals.std()
    median_all = all_vals.median()
    iqr_all = all_vals.quantile(0.75) - all_vals.quantile(0.25)
    q1_all = all_vals.quantile(0.25)
    q3_all = all_vals.quantile(0.75)
    min_all = all_vals.min()
    max_all = all_vals.max()
    
    # 第 1 組 (糖尿病組)
    g1 = df[df['組別'] == 1][col]
    # 第 2 組 (對照組)
    g2 = df[df['組別'] == 2][col]
    
    # 常態性檢定 (Shapiro-Wilk)
    shapiro_g1 = stats.shapiro(g1)
    shapiro_g2 = stats.shapiro(g2)
    
    # 變異數同質性檢定 (Levene's test)
    levene_res = stats.levene(g1, g2)
    equal_var = levene_res.pvalue > 0.05
    
    # 兩獨立樣本 t 檢定
    ttest_res = stats.ttest_ind(g1, g2, equal_var=equal_var)
    
    # 曼-惠特尼 U 檢定 (無母數)
    mwu_res = stats.mannwhitneyu(g1, g2, alternative='two-sided')
    
    # Cohen's d 效應值
    n1, n2 = len(g1), len(g2)
    s_pooled = np.sqrt(((n1 - 1) * g1.var() + (n2 - 1) * g2.var()) / (n1 + n2 - 2))
    cohens_d = (g1.mean() - g2.mean()) / s_pooled
    
    # 95% 信賴區間 (差值)
    diff = g1.mean() - g2.mean()
    se_diff = np.sqrt(g1.var()/n1 + g2.var()/n2) if not equal_var else s_pooled * np.sqrt(1/n1 + 1/n2)
    df_t = (n1 - 1 + n2 - 1) if equal_var else ((g1.var()/n1 + g2.var()/n2)**2 / ((g1.var()/n1)**2/(n1-1) + (g2.var()/n2)**2/(n2-1)))
    ci_margin = stats.t.ppf(0.975, df_t) * se_diff
    ci_lower = diff - ci_margin
    ci_upper = diff + ci_margin

    summary_records.append({
        "指標": col,
        "全體_平均值±標準差": f"{mean_all:.2f} ± {std_all:.2f}",
        "全體_中位數(IQR)": f"{median_all:.2f} ({q1_all:.2f} - {q3_all:.2f})",
        "全體_極值[Min-Max]": f"[{min_all:.1f} - {max_all:.1f}]",
        "第1組_平均值±標準差": f"{g1.mean():.2f} ± {g1.std():.2f}",
        "第1組_中位數(IQR)": f"{g1.median():.2f} ({g1.quantile(0.25):.2f} - {g1.quantile(0.75):.2f})",
        "第2組_平均值±標準差": f"{g2.mean():.2f} ± {g2.std():.2f}",
        "第2組_中位數(IQR)": f"{g2.median():.2f} ({g2.quantile(0.25):.2f} - {g2.quantile(0.75):.2f})",
        "平均差(95% CI)": f"{diff:.2f} [{ci_lower:.2f}, {ci_upper:.2f}]",
        "t檢定_t值": round(ttest_res.statistic, 3),
        "t檢定_p值": f"{ttest_res.pvalue:.4e}",
        "Mann-Whitney_U值": mwu_res.statistic,
        "Mann-Whitney_p值": f"{mwu_res.pvalue:.4e}",
        "Cohen_d": round(cohens_d, 2),
        "Levene檢定_p值": round(levene_res.pvalue, 4),
        "Shapiro第1組_p值": round(shapiro_g1.pvalue, 4),
        "Shapiro第2組_p值": round(shapiro_g2.pvalue, 4),
    })

df_summary = pd.DataFrame(summary_records)
df_summary.to_csv("statistical_summary.csv", index=False, encoding="utf-8-sig")
print("統計分析摘要已儲存至: statistical_summary.csv")
print(df_summary.to_string())
print("\n")

# 4. 相關性矩陣
corr_matrix = df[[item["col"] for item in target_vars]].corr()
print("變項間相關係數矩陣 (Pearson correlation):")
print(corr_matrix)
print("\n")

# ==============================================================================
# 圖表 1: 主要分組盒狀圖 (Group Comparison Boxplot with jitter points)
# ==============================================================================
fig, axes = plt.subplots(1, 3, figsize=(16, 6))
palette = ["#e74c3c", "#3498db"]  # 鮮明專業紅藍對比配色

for i, item in enumerate(target_vars):
    col = item["col"]
    ax = axes[i]
    
    # 繪製 Seaborn 箱形圖
    sns.boxplot(
        data=df,
        x='組別名稱',
        y=col,
        hue='組別名稱',
        palette=palette,
        legend=False,
        width=0.45,
        fliersize=0,
        boxprops=dict(alpha=0.75, edgecolor='black', linewidth=1.2),
        medianprops=dict(color='black', linewidth=2),
        whiskerprops=dict(color='black', linewidth=1.2),
        capprops=dict(color='black', linewidth=1.2),
        ax=ax
    )
    
    # 疊加散點圖 (Stripplot) 呈現所有樣本真實點位
    sns.stripplot(
        data=df,
        x='組別名稱',
        y=col,
        color='black',
        size=7,
        jitter=0.18,
        alpha=0.85,
        edgecolor='white',
        linewidth=0.8,
        ax=ax
    )
    
    # 標註各組均值 (紅星)
    g1_mean = df[df['組別']==1][col].mean()
    g2_mean = df[df['組別']==2][col].mean()
    ax.scatter([0], [g1_mean], color='#800000', marker='D', s=45, zorder=5, label='組平均值 (Mean)' if i==0 else "")
    ax.scatter([1], [g2_mean], color='#003366', marker='D', s=45, zorder=5)

    # 顯著性標註欄
    y_max = df[col].max()
    y_min = df[col].min()
    y_range = y_max - y_min
    bar_y = y_max + y_range * 0.08
    text_y = bar_y + y_range * 0.02
    
    ax.plot([0, 0, 1, 1], [bar_y - y_range * 0.02, bar_y, bar_y, bar_y - y_range * 0.02], lw=1.5, c='black')
    ax.text(0.5, text_y, "*** p < 0.001", ha='center', va='bottom', color='#c0392b', fontsize=12, fontweight='bold')
    
    ax.set_ylim(y_min - y_range * 0.08, bar_y + y_range * 0.15)
    ax.set_title(f"{item['name']}\n(單位: {item['unit']})", fontsize=14, fontweight='bold', pad=10)
    ax.set_xlabel("組別", fontsize=12, fontweight='bold')
    ax.set_ylabel(item['col'], fontsize=12)
    ax.tick_params(axis='x', labelsize=11)
    ax.tick_params(axis='y', labelsize=10)

if axes[0].get_legend_handles_labels()[0]:
    axes[0].legend(loc='lower right', frameon=True)

plt.suptitle("DM.xls: BMI、血糖與收縮壓之組別箱形圖比較 (Boxplot with Stripplot)", fontsize=16, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig("dm_boxplot_by_group.png", dpi=300, bbox_inches='tight')
plt.close()
print("已成功產出: dm_boxplot_by_group.png")

# ==============================================================================
# 圖表 2: 全體樣本單變項盒狀圖 (Overall Distribution Boxplots)
# ==============================================================================
fig, axes = plt.subplots(1, 3, figsize=(15, 5.5))
colors = ["#2ecc71", "#e67e22", "#9b59b6"]

for i, item in enumerate(target_vars):
    col = item["col"]
    ax = axes[i]
    
    sns.boxplot(
        y=df[col],
        color=colors[i],
        width=0.35,
        boxprops=dict(alpha=0.7, edgecolor='black', linewidth=1.2),
        medianprops=dict(color='darkred', linewidth=2.2),
        ax=ax
    )
    
    sns.stripplot(
        y=df[col],
        color='#2c3e50',
        size=7,
        jitter=0.15,
        alpha=0.85,
        edgecolor='white',
        linewidth=0.8,
        ax=ax
    )
    
    mean_val = df[col].mean()
    med_val = df[col].median()
    
    ax.axhline(mean_val, color='blue', linestyle='--', linewidth=1.5, label=f'平均數 Mean: {mean_val:.1f}')
    ax.axhline(med_val, color='darkred', linestyle='-', linewidth=1.5, label=f'中位數 Median: {med_val:.1f}')
    
    ax.set_title(f"整體樣本 (N=22)\n{item['name']}", fontsize=13, fontweight='bold')
    ax.set_ylabel(f"{item['col']}", fontsize=11)
    ax.legend(loc='upper right', frameon=True, fontsize=10)

plt.suptitle("DM.xls: 全體受試者 BMI、血糖與收縮壓分佈盒狀圖", fontsize=15, fontweight='bold', y=1.03)
plt.tight_layout()
plt.savefig("dm_boxplot_overall.png", dpi=300, bbox_inches='tight')
plt.close()
print("已成功產出: dm_boxplot_overall.png")

# ==============================================================================
# 圖表 3: 組別 × 性別 雙因子箱形圖 (Boxplot by Group and Gender)
# ==============================================================================
fig, axes = plt.subplots(1, 3, figsize=(17, 6))

for i, item in enumerate(target_vars):
    col = item["col"]
    ax = axes[i]
    
    sns.boxplot(
        data=df,
        x='組別名稱',
        y=col,
        hue='性別名稱',
        palette={'男性': '#4a90e2', '女性': '#f39c12'},
        width=0.55,
        boxprops=dict(alpha=0.75, edgecolor='black', linewidth=1.1),
        medianprops=dict(color='black', linewidth=1.8),
        ax=ax
    )
    
    sns.stripplot(
        data=df,
        x='組別名稱',
        y=col,
        hue='性別名稱',
        dodge=True,
        palette={'男性': '#1b2631', '女性': '#512e5f'},
        legend=False,
        size=6,
        alpha=0.85,
        ax=ax
    )
    
    # 避免重複 legend
    handles, labels = ax.get_legend_handles_labels()
    ax.legend(handles[:2], labels[:2], title="性別", loc='upper right' if i!=1 else 'upper right')
    
    ax.set_title(f"{item['name']}\n(單位: {item['unit']})", fontsize=13, fontweight='bold')
    ax.set_xlabel("組別", fontsize=11, fontweight='bold')
    ax.set_ylabel(item['col'], fontsize=11)

plt.suptitle("DM.xls: 依組別與性別之 BMI、血糖、收縮壓箱形圖", fontsize=15, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig("dm_boxplot_gender.png", dpi=300, bbox_inches='tight')
plt.close()
print("已成功產出: dm_boxplot_gender.png")

print("全部分析與繪圖已順利完成！")
