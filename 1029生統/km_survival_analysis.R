# ==============================================================================
# Kaplan-Meier 存活分析與曲線繪製 (Kaplan-Meier Survival Analysis)
# 新藥試驗數據 (Survival test example.xlsx)
# ==============================================================================

Sys.setlocale("LC_ALL", "en_US.UTF-8")

suppressPackageStartupMessages({
  library(readxl)
  library(survival)
  library(ggplot2)
  library(grid)
  library(gtable)
  library(scales)
})

# 1. 讀取與整理數據
cat("--> 讀取數據中...\n")
excel_file <- "Survival test example.xlsx"
raw_data <- read_excel(excel_file)
colnames(raw_data) <- c("ID", "Group", "Time", "Status")

# 建立分組標籤
raw_data$GroupFactor <- factor(
  raw_data$Group,
  levels = c(1, 2),
  labels = c("Group 1 (New Drug 新藥組)", "Group 2 (Control 對照組)")
)

cat("--> 數據摘要:\n")
print(raw_data)

# 2. Kaplan-Meier 存活分析
cat("\n--> 建立 Kaplan-Meier 模型...\n")
km_fit <- survfit(Surv(Time, Status) ~ GroupFactor, data = raw_data)
cat("\n--- KM 模型摘要 ---\n")
print(summary(km_fit))

# 3. Log-Rank 檢定 (檢驗兩組存活曲線是否有統計顯著差異)
logrank_test <- survdiff(Surv(Time, Status) ~ GroupFactor, data = raw_data)
chisq_val <- logrank_test$chisq
df_val <- length(logrank_test$n) - 1
p_val <- 1 - pchisq(chisq_val, df_val)

cat("\n--- Log-Rank Test 結果 ---\n")
print(logrank_test)
cat(sprintf("Chisq = %.2f, df = %d, p-value = %.4f\n", chisq_val, df_val, p_val))

# 4. Cox 等比例風險模型 (計算 Hazard Ratio)
cox_model <- coxph(Surv(Time, Status) ~ Group, data = raw_data)
cox_sum <- summary(cox_model)
# Group 1 vs Group 2 (以對照組 Group 2 為參考基準，或以新藥組為基準)
# Group 2 vs 1: HR = exp(coef) = 2.809
# Group 1 vs 2: HR = exp(-coef) = 0.356
hr_val <- exp(-cox_sum$coefficients[1, "coef"])
hr_lower <- exp(-cox_sum$coefficients[1, "coef"] - 1.96 * cox_sum$coefficients[1, "se(coef)"])
hr_upper <- exp(-cox_sum$coefficients[1, "coef"] + 1.96 * cox_sum$coefficients[1, "se(coef)"])

cat(sprintf("Hazard Ratio (Group 1 vs Group 2): %.3f (95%% CI: %.3f - %.3f)\n", hr_val, hr_lower, hr_upper))

# 5. 輸出詳細 Life Table 至 CSV
summary_df <- data.frame(
  Group = rep(names(km_fit$strata), km_fit$strata),
  Time = km_fit$time,
  N_Risk = km_fit$n.risk,
  N_Event = km_fit$n.event,
  N_Censor = km_fit$n.censor,
  Survival = round(km_fit$surv, 4),
  Std_Error = round(km_fit$std.err, 4),
  Lower_95_CI = round(km_fit$lower, 4),
  Upper_95_CI = round(km_fit$upper, 4)
)
summary_df$Group <- gsub("GroupFactor=", "", summary_df$Group)
write.csv(summary_df, "survival_summary_table.csv", row.names = FALSE)
cat("--> 已儲存詳細存活表至 survival_summary_table.csv\n")

# 6. 整理畫圖所需的階梯資料 (Step function data)
build_km_df <- function(fit) {
  groups <- names(fit$strata)
  df_list <- list()
  
  idx_start <- 1
  for (i in seq_along(groups)) {
    g_name <- gsub("GroupFactor=", "", groups[i])
    n_pts <- fit$strata[i]
    idx_end <- idx_start + n_pts - 1
    
    times <- fit$time[idx_start:idx_end]
    survs <- fit$surv[idx_start:idx_end]
    lowers <- fit$lower[idx_start:idx_end]
    uppers <- fit$upper[idx_start:idx_end]
    censors <- fit$n.censor[idx_start:idx_end]
    events <- fit$n.event[idx_start:idx_end]
    
    # 加入時間 0
    df_g <- data.frame(
      Time = c(0, times),
      Survival = c(1.0, survs),
      Lower = c(1.0, lowers),
      Upper = c(1.0, uppers),
      Censor = c(0, censors),
      Event = c(0, events),
      Group = g_name
    )
    df_list[[i]] <- df_g
    idx_start <- idx_end + 1
  }
  do.call(rbind, df_list)
}

plot_data <- build_km_df(km_fit)

# 篩選出設限 (Censored) 點資料
censor_pts <- plot_data[plot_data$Censor > 0, ]

# 7. 計算 Number at Risk 表格資料 (時間點 0 到 8 個月)
time_ticks <- seq(0, 8, by = 1)
risk_summary <- summary(km_fit, times = time_ticks)
risk_table_df <- data.frame(
  Time = risk_summary$time,
  Risk = risk_summary$n.risk,
  Group = gsub("GroupFactor=", "", risk_summary$strata)
)

# 補齊可能缺少的時間點 (例如 Group 2 在 t=8 時 risk=0)
all_combinations <- expand.grid(Time = time_ticks, Group = unique(plot_data$Group))
risk_table_df <- merge(all_combinations, risk_table_df, by = c("Time", "Group"), all.x = TRUE)
risk_table_df$Risk[is.na(risk_table_df$Risk)] <- 0
risk_table_df$Group <- factor(risk_table_df$Group, levels = c("Group 2 (Control 對照組)", "Group 1 (New Drug 新藥組)"))

# 8. 設定配色與主題
color_palette <- c(
  "Group 1 (New Drug 新藥組)" = "#1E88E5",  # 專業醫療藍
  "Group 2 (Control 對照組)"  = "#E53935"   # 對照組珊瑚紅
)

# 9. 繪製主存活曲線圖
p_main <- ggplot(plot_data, aes(x = Time, y = Survival, color = Group, linetype = Group)) +
  # 階梯曲線
  geom_step(linewidth = 1.4, direction = "hv") +
  # 設限標記 (Censored ticks: + 字號)
  geom_point(data = censor_pts, aes(x = Time, y = Survival), shape = 3, size = 4, stroke = 1.8, show.legend = FALSE) +
  # 50% 存活率參考線 (中位存活時間)
  geom_hline(yintercept = 0.5, linetype = "dashed", color = "gray50", linewidth = 0.7) +
  # 中位存活時間垂直標示線
  annotate("segment", x = 6.5, xend = 6.5, y = 0, yend = 0.5, linetype = "dotted", color = "#1E88E5", linewidth = 0.9) +
  annotate("segment", x = 5.25, xend = 5.25, y = 0, yend = 0.5, linetype = "dotted", color = "#E53935", linewidth = 0.9) +
  # 標記中位存活時間數值
  annotate("text", x = 6.6, y = 0.26, label = "Median: 6.50 mo", color = "#1E88E5", fontface = "bold", size = 4.0, hjust = 0) +
  annotate("text", x = 5.35, y = 0.14, label = "Median: 5.25 mo", color = "#E53935", fontface = "bold", size = 4.0, hjust = 0) +
  # 統計檢定資訊框 (Annotated Statistics Box)
  annotate(
    "label",
    x = 0.25, y = 0.38,
    hjust = 0, vjust = 1,
    label = sprintf(
      "Log-Rank Test:\n  Chi-Square (\u03c7\u00b2) = %.2f (df = %d)\n  p-value = %.4f (NS)\n\nCox Proportional Hazards:\n  Hazard Ratio (HR) = %.2f\n  95%% CI: [%.2f, %.2f]\n\nCensored Observations: +",
      chisq_val, df_val, p_val, hr_val, hr_lower, hr_upper
    ),
    fill = alpha("white", 0.92),
    color = "#2C3E50",
    size = 3.6,
    label.r = unit(0.3, "lines"),
    label.size = 0.35,
    label.padding = unit(0.5, "lines")
  ) +
  scale_color_manual(values = color_palette) +
  scale_linetype_manual(values = c("Group 1 (New Drug 新藥組)" = "solid", "Group 2 (Control 對照組)" = "dashed")) +
  scale_x_continuous(breaks = time_ticks, limits = c(0, 8.2), expand = c(0.015, 0)) +
  scale_y_continuous(breaks = seq(0, 1, 0.2), limits = c(-0.02, 1.02), labels = percent_format(accuracy = 1)) +
  labs(
    title = "Kaplan-Meier Survival Curve (新藥試驗存活曲線)",
    subtitle = "Comparison of Overall Survival: Group 1 (New Drug) vs Group 2 (Control)",
    x = "Follow-up Months (追蹤月份)",
    y = "Survival Probability (存活率)",
    color = "Treatment Group",
    linetype = "Treatment Group"
  ) +
  theme_bw(base_size = 13, base_family = "PingFang SC") +
  theme(
    plot.title = element_text(face = "bold", size = 16, color = "#1A252C", margin = margin(b = 5)),
    plot.subtitle = element_text(size = 11.5, color = "#546E7A", margin = margin(b = 12)),
    legend.position = "inside",
    legend.position.inside = c(0.74, 0.88),
    legend.background = element_rect(fill = alpha("white", 0.92), color = "#CFD8DC", linewidth = 0.4),
    legend.title = element_text(face = "bold", size = 10.5),
    legend.text = element_text(size = 10),
    legend.key.width = unit(2.2, "lines"),
    panel.grid.minor = element_blank(),
    panel.grid.major = element_line(color = "#ECEFF1", linewidth = 0.5),
    panel.border = element_rect(color = "#B0BEC5", fill = NA, linewidth = 0.8),
    axis.title = element_text(face = "bold", size = 11.5, color = "#263238"),
    axis.text = element_text(size = 10.5, color = "#37474F"),
    plot.margin = margin(t = 14, r = 20, b = 4, l = 20)
  )

# 10. 繪製 Number at Risk 表格
p_table <- ggplot(risk_table_df, aes(x = Time, y = Group, label = Risk)) +
  geom_text(aes(color = Group), size = 4.2, fontface = "bold") +
  scale_color_manual(values = color_palette, guide = "none") +
  scale_x_continuous(breaks = time_ticks, limits = c(0, 8.2), expand = c(0.015, 0)) +
  scale_y_discrete(labels = c("Control (對照組)", "New Drug (新藥組)")) +
  labs(
    x = "Follow-up Months (追蹤月份)",
    y = "Number at Risk\n(在險受試人數)"
  ) +
  theme_classic(base_size = 11, base_family = "PingFang SC") +
  theme(
    axis.title.x = element_text(face = "bold", size = 11, color = "#263238", margin = margin(t = 6)),
    axis.title.y = element_text(face = "bold", size = 10, color = "#263238", angle = 0, vjust = 0.5, hjust = 1),
    axis.text.x = element_blank(),
    axis.ticks.x = element_blank(),
    axis.text.y = element_text(size = 10, face = "bold", color = c("#E53935", "#1E88E5")),
    axis.line = element_blank(),
    axis.ticks.y = element_blank(),
    plot.margin = margin(t = 2, r = 20, b = 12, l = 20)
  )

# 11. 結合主圖與風險表 (使用 gtable 對齊 x 軸)
cat("--> 對齊圖表並輸出高畫質圖檔...\n")
p_main_notitle <- p_main + theme(
  axis.title.x = element_blank(),
  plot.margin = margin(t = 14, r = 20, b = 2, l = 20)
)

g_main <- ggplotGrob(p_main_notitle)
g_table <- ggplotGrob(p_table)

# 統一兩者的寬度以精確對齊 x 軸
max_widths <- unit.pmax(g_main$widths, g_table$widths)
g_main$widths <- max_widths
g_table$widths <- max_widths

draw_combined <- function() {
  grid.newpage()
  pushViewport(viewport(layout = grid.layout(nrow = 2, heights = unit(c(0.77, 0.23), c("null", "null")))))
  pushViewport(viewport(layout.pos.row = 1))
  grid.draw(g_main)
  popViewport()
  pushViewport(viewport(layout.pos.row = 2))
  grid.draw(g_table)
  popViewport()
}

# 儲存 PNG (300 DPI 高畫質)
png_filename <- "kaplan_meier_curve.png"
png(png_filename, width = 3000, height = 2400, res = 300, family = "PingFang SC")
draw_combined()
dev.off()
cat(sprintf("--> 已輸出高解析度圖檔: %s\n", png_filename))

# 儲存 PDF (向量圖)
pdf_filename <- "kaplan_meier_curve.pdf"
quartz(type = "pdf", file = pdf_filename, width = 10, height = 8, family = "PingFang SC")
draw_combined()
dev.off()
cat(sprintf("--> 已輸出向量圖檔: %s\n", pdf_filename))

# 12. 繪製附有 95% 信賴區間 (Confidence Intervals) 的版本
p_ci <- ggplot(plot_data, aes(x = Time, y = Survival, color = Group, fill = Group)) +
  geom_ribbon(aes(ymin = Lower, ymax = Upper), alpha = 0.15, linetype = 0, show.legend = FALSE) +
  geom_step(linewidth = 1.3, direction = "hv") +
  geom_point(data = censor_pts, aes(x = Time, y = Survival), shape = 3, size = 3.5, stroke = 1.5, show.legend = FALSE) +
  geom_hline(yintercept = 0.5, linetype = "dashed", color = "gray55", linewidth = 0.7) +
  scale_color_manual(values = color_palette) +
  scale_fill_manual(values = color_palette) +
  scale_x_continuous(breaks = time_ticks, limits = c(0, 8.2), expand = c(0.01, 0)) +
  scale_y_continuous(breaks = seq(0, 1, 0.2), limits = c(-0.02, 1.02), labels = percent_format(accuracy = 1)) +
  labs(
    title = "Kaplan-Meier Survival Curve with 95% CI (附95%信賴區間)",
    subtitle = "Shaded bands represent 95% Pointwise Confidence Intervals",
    x = "Follow-up Months (追蹤月份)",
    y = "Survival Probability (存活率)",
    color = "Treatment Group"
  ) +
  theme_bw(base_size = 13, base_family = "PingFang SC") +
  theme(
    plot.title = element_text(face = "bold", size = 15, color = "#1A252C"),
    plot.subtitle = element_text(size = 11, color = "#546E7A", margin = margin(b = 10)),
    legend.position = "inside",
    legend.position.inside = c(0.72, 0.88),
    legend.background = element_rect(fill = alpha("white", 0.9), color = "#CFD8DC", linewidth = 0.3),
    panel.grid.minor = element_blank(),
    panel.grid.major = element_line(color = "#ECEFF1", linewidth = 0.5),
    axis.title = element_text(face = "bold", size = 11.5, color = "#263238"),
    axis.text = element_text(size = 10, color = "#37474F")
  )

png("kaplan_meier_curve_with_CI.png", width = 2700, height = 2100, res = 300, family = "PingFang SC")
print(p_ci)
dev.off()
cat("--> 已輸出 95% CI 存活曲線: kaplan_meier_curve_with_CI.png\n")

cat("\n=======================================================\n")
cat("全部 Kaplan-Meier 分析圖表與數據表已成功生成！\n")
cat("=======================================================\n")
