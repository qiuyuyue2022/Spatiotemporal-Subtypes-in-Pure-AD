library(readxl)
library(dplyr)
library(writexl)
library(tidyr)

# 读取数据
df <- read_xlsx("/Users/qiuyuyue/Documents/MRI/data/SuStaIn/mri_from4500_renamed_with_zscores.xlsx")
df <- df %>%
  mutate(age_group = case_when(
    age <= 54 ~ "<=54",
    age >= 55 & age <= 59 ~ "55-59",
    age >= 60 & age <= 64 ~ "60-64",
    age >= 65 & age <= 69 ~ "65-69",
    age >= 70 & age <= 74 ~ "70-74",
    age >= 75 & age <= 79 ~ "75-79",
    age >= 80 & age <= 84 ~ "80-84",
    age >= 85 ~ ">=85"
  ))

mean_data <- read_xlsx("/Users/qiuyuyue/Documents/MRI/data/Vmri_mean_sd202507.xlsx", sheet = "mean")
# 将均值数据转换为长格式，方便后续匹配
mean_long <- mean_data %>%
  pivot_longer(cols = -1, names_to = "Brain_Region", values_to = "Mean") %>%
  rename(age_group = 1)  # 将第一列重命名为 age_group


# 定义函数：加权zcombined
weighted_zcombined <- function(df, regions, suffix = "_z", mean_long) {
  z_scores <- df %>% select(paste0(regions, suffix))
  age_groups <- df$age_group
  
  weights_matrix <- matrix(NA, nrow = nrow(df), ncol = length(regions))
  
  for (i in seq_along(regions)) {
    region <- regions[i]
    # 针对每个患者，找到该年龄组 + 区域对应的标准均值
    weights_matrix[, i] <- sapply(seq_along(age_groups), function(j) {
      weight <- mean_long %>%
        filter(age_group == age_groups[j], Brain_Region == region) %>%
        pull(Mean)
      if (length(weight) == 0) NA else weight
    })
  }
  
  sum_weights <- rowSums(weights_matrix, na.rm = TRUE)
  z_scores_mat <- as.matrix(z_scores)
  weighted_z <- rowSums(weights_matrix * z_scores_mat, na.rm = TRUE)
  
  return(weighted_z / sum_weights)
}

# 开始计算各个指标
df <- df %>%
  mutate(
    CC_zcombined = weighted_zcombined(df, c("CC_Anterior_pct", "CC_Mid_Anterior_pct", "CC_Central_pct", "CC_Mid_Posterior_pct", "CC_Posterior_pct"), mean_long = mean_long),
    PCC_zcombined_L = weighted_zcombined(df, c("Cingulum_Post_L_pct", "Isthmus_of_cingulate_L_pct"), mean_long = mean_long),
    PCC_zcombined_R = weighted_zcombined(df, c("Cingulum_Post_R_pct", "Isthmus_of_cingulate_R_pct"), mean_long = mean_long),
    
    BG_zcombined_L = weighted_zcombined(df, c("Caudate_L_pct", "Pallidum_L_pct", "Putamen_L_pct", "Thalamus_L_pct", "Accumbens_Area_L_pct"), mean_long = mean_long),
    BG_zcombined_R = weighted_zcombined(df, c("Caudate_R_pct", "Pallidum_R_pct", "Putamen_R_pct", "Thalamus_R_pct", "Accumbens_Area_R_pct"), mean_long = mean_long),
    
    Temp_Neo_zcombined_L = weighted_zcombined(df, c("Temporal_Sup_L_pct", "Temporal_Sup_Banks_L_pct", "Temporal_Mid_L_pct", "Temporal_Inf_L_pct", "Transverse_temporal_L_pct", "Fusiform_L_pct", "Temporal_Pole_L_pct"), mean_long = mean_long),
    Temp_Neo_zcombined_R = weighted_zcombined(df, c("Temporal_Sup_R_pct", "Temporal_Sup_Banks_R_pct", "Temporal_Mid_R_pct", "Temporal_Inf_R_pct", "Transverse_temporal_R_pct", "Fusiform_R_pct", "Temporal_Pole_R_pct"), mean_long = mean_long),
    
    OFC_zcombined_L = weighted_zcombined(df, c("Orbitofrontal_Lat_L_pct", "Orbitofrontal_Med_L_pct"), mean_long = mean_long),
    OFC_zcombined_R = weighted_zcombined(df, c("Orbitofrontal_Lat_R_pct", "Orbitofrontal_Med_R_pct"), mean_long = mean_long),
    PFC_zcombined_L = weighted_zcombined(df, c("Frontal_Sup_L_pct", "Frontal_Mid_Rostral_L_pct", "Frontal_Mid_Caudal_L_pct", "Pars_opercularis_L_pct", "Pars_orbitalis_L_pct", "Pars_triangularis_L_pct", "Frontal_pole_L_pct"), mean_long = mean_long),
    PFC_zcombined_R = weighted_zcombined(df, c("Frontal_Sup_R_pct", "Frontal_Mid_Rostral_R_pct", "Frontal_Mid_Caudal_R_pct", "Pars_opercularis_R_pct", "Pars_orbitalis_R_pct", "Pars_triangularis_R_pct", "Frontal_pole_R_pct"), mean_long = mean_long),
    
    IPL_zcombined_L = weighted_zcombined(df, c("Parietal_Inf_L_pct", "Supramarginal_L_pct"), mean_long = mean_long),
    IPL_zcombined_R = weighted_zcombined(df, c("Parietal_Inf_R_pct", "Supramarginal_R_pct"), mean_long = mean_long),
    
    Med_Occi_zcombined_L = weighted_zcombined(df, c("Cuneus_L_pct", "Lingual_L_pct", "Peri_calcarine_L_pct"), mean_long = mean_long),
    Med_Occi_zcombined_R = weighted_zcombined(df, c("Cuneus_R_pct", "Lingual_R_pct", "Peri_calcarine_R_pct"), mean_long = mean_long)
  )

# 计算新列
df <- df %>%
  mutate(
    Amygdala_zcombined = (Amygdala_L_pct_z + Amygdala_R_pct_z) / 2,
    BG_zcombined = (BG_zcombined_L + BG_zcombined_R) / 2,
    ChP_zcombined = (Choroid_Plexus_L_pct_z + Choroid_Plexus_R_pct_z) / 2,
    Lat_Occi_zcombined = (Occipital_Lat_L_pct_z + Occipital_Lat_R_pct_z) / 2,
    Med_Occi_zcombined = (Med_Occi_zcombined_L + Med_Occi_zcombined_R) / 2,
    OFC_zcombined = (OFC_zcombined_L + OFC_zcombined_R) / 2,
    SPL_zcombined = (Parietal_Sup_L_pct_z + Parietal_Sup_R_pct_z) / 2,
    IPL_zcombined = (IPL_zcombined_L + IPL_zcombined_R) / 2,
    PFC_zcombined = (PFC_zcombined_L + PFC_zcombined_R) / 2,
    PCC_zcombined = (PCC_zcombined_L + PCC_zcombined_R) / 2,
    Precuneus_zcombined = (Precuneus_L_pct_z + Precuneus_R_pct_z) / 2,
    Temp_Neo_zcombined = (Temp_Neo_zcombined_L + Temp_Neo_zcombined_R) / 2,
    Insula_zcombined = (Insula_L_pct_z + Insula_R_pct_z) / 2,
    Hippo_zcombined = (Hippocampus_L_pct_z + Hippocampus_R_pct_z) / 2,
    Parahippo_zcombined = (Parahippocampal_L_pct_z + Parahippocampal_R_pct_z) / 2,
    Entorhinal_zcombined = (Entorhinal_L_pct_z + Entorhinal_R_pct_z) / 2
  )

library(readxl)
library(ggplot2)
zcombined_cols <- grep("_zcombined$", names(df), value = TRUE)
length(zcombined_cols)
print(zcombined_cols)
# 顺序重排
zcombined_cols <- c(
  "Amygdala_zcombined", "Hippo_zcombined", "Parahippo_zcombined", "Entorhinal_zcombined", 
  "Temp_Neo_zcombined",
  "PCC_zcombined","SPL_zcombined","IPL_zcombined", "Precuneus_zcombined", 
  "OFC_zcombined", "PFC_zcombined", 
  "Lat_Occi_zcombined", "Med_Occi_zcombined", 
  "Insula_zcombined", 
  "BG_zcombined", "CC_zcombined", 
  "ChP_zcombined"
)

# 获取数据框中其余列（非 zcombined 列）
non_zcombined_cols <- setdiff(names(df), zcombined_cols)

# 重新排序：非 zcombined 列 + 按解剖顺序排序的 zcombined 列
df <- df[, c(non_zcombined_cols, zcombined_cols)]
nrow(df)
df <- df %>% select(-ends_with("_z"))
df <- df %>% select(-ends_with("_L"))
df <- df %>% select(-ends_with("_R"))

# 可选择保存新表
library(openxlsx)
write_xlsx(df, "/Users/qiuyuyue/Documents/MRI/data/SuStaIn/mri_from4500_renamed_with_zscores_17zcomined.xlsx")


# 逐列检查每个 _zcombined 变量的最小值、最大值和NA个数
summary_stats <- sapply(df[zcombined_cols], function(x) {
  c(min = min(x, na.rm = TRUE),
    max = max(x, na.rm = TRUE),
    mean = mean(x, na.rm = TRUE),
    sd = sd(x, na.rm = TRUE),
    NAs = sum(is.na(x)))
})
# 转置结果以便查看
summary_stats <- t(summary_stats)
# 打印结果
print(summary_stats)

# 将数据转长格式，便于 ggplot 画图
library(tidyr)
library(dplyr)
df_long <- df %>%
  select(all_of(zcombined_cols)) %>%
  pivot_longer(cols = everything(), names_to = "region", values_to = "value")
# 画出密度图
ggplot(df_long, aes(x = value)) +
  geom_density(fill = "steelblue", alpha = 0.5) +
  facet_wrap(~ region, scales = "free") +
  theme_minimal() +
  labs(title = "Distribution of _zcombined Variables", x = "Z value", y = "Density")



# 筛查异常值
outliers_list <- list()
# 遍历每一列，检查 z 值是否超过阈值（绝对值 > 5）
for (col in zcombined_cols) {
  outlier_rows <- which(abs(df[[col]]) > 5 & !is.na(df[[col]]))  # 非NA且|z|>3
  if (length(outlier_rows) > 0) {
    outliers_list[[col]] <- data.frame(
      Row = outlier_rows,
      Value = df[[col]][outlier_rows]
    )
  }
}
# 打印异常值结果
if (length(outliers_list) == 0) {
  cat("没有发现异常值。\n")
} else {
  for (col in names(outliers_list)) {
    cat("异常值列:", col, "\n")
    print(outliers_list[[col]])
    cat("\n")
  }
}


