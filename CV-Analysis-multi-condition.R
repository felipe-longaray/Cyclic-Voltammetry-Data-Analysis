# ==============================================================================
# Cyclic Voltammetry (CV) Surface Modification Analysis 
# Automated peak extraction, multi-condition array mapping, and ANOVA
# ==============================================================================

library(readxl)
library(dplyr)
library(ggplot2)
library(ggpubr)
library(openxlsx)

# ------------------------------------------------------------------------------
# 1. Data Import and Standardization
# ------------------------------------------------------------------------------
data_path <- "./data/CV_Array_Output.xlsx"
raw_data <- read_excel(data_path)

orig_names <- colnames(raw_data)
new_names <- character(length(orig_names))

for (i in seq_along(orig_names)) {
  if (i %% 2 == 1) {  
    new_names[i] <- paste0("POTENTIAL_", orig_names[i])
  } else {            
    prev_name <- new_names[i-1]  
    new_names[i] <- sub("POTENTIAL_", "CURRENT_", prev_name)  
  }
}

colnames(raw_data) <- new_names
raw_data <- raw_data[-1, ] # Remove metadata row

col_current <- seq(2, ncol(raw_data), by = 2)
raw_data[col_current] <- lapply(raw_data[col_current], as.numeric)

col_potential <- seq(1, ncol(raw_data), by = 2)
raw_data[col_potential] <- lapply(raw_data[col_potential], as.numeric)

# ------------------------------------------------------------------------------
# 2. Extract Electrochemical Parameters
# ------------------------------------------------------------------------------
potential_cols <- grep("POTENTIAL_", names(raw_data), value = TRUE)
current_cols <- grep("CURRENT_", names(raw_data), value = TRUE)

if (length(potential_cols) != length(current_cols)) {
  stop("Data structure error: Mismatch between Potential and Current columns.")
}

cv_results <- data.frame()

for (i in seq_along(potential_cols)) {
  potential_vec <- raw_data[[potential_cols[i]]]
  current_vec <- raw_data[[current_cols[i]]]
  
  i_pa <- max(current_vec, na.rm = TRUE)
  i_pc <- min(current_vec, na.rm = TRUE)
  
  e_pa <- potential_vec[which.max(current_vec)]
  e_pc <- potential_vec[which.min(current_vec)]
  
  delta_ep <- abs(e_pa - e_pc)
  
  # Condition mapping (Obfuscated for repository publishing)
  # NOTE: Replace 'BioMod' with your actual raw data string if executing locally
  condition <- case_when(
    grepl("SPE_", current_cols[i]) & !grepl("60|150", current_cols[i]) ~ "Bare SPE",
    grepl("SPE60", current_cols[i]) ~ "Bare SPE (60s)",
    grepl("SPE150", current_cols[i]) ~ "Bare SPE (150s)",
    grepl("BioMod_", current_cols[i]) & !grepl("60|150", current_cols[i]) ~ "Functionalized SPE",
    grepl("BioMod60", current_cols[i]) ~ "Functionalized SPE (60s)",
    grepl("BioMod150", current_cols[i]) ~ "Functionalized SPE (150s)",
    TRUE ~ "Other"
  )
  
  cv_results <- rbind(cv_results, data.frame(
    Measurement = i,
    I_pa = i_pa,
    E_pa = e_pa,
    I_pc = i_pc,
    E_pc = e_pc,
    Delta_Ep = delta_ep,
    Condition = condition
  ))
}

# Define ordinal levels for logical plotting (Control group first, then modified)
condition_levels <- c("Bare SPE", "Bare SPE (60s)", "Bare SPE (150s)", 
                      "Functionalized SPE", "Functionalized SPE (60s)", "Functionalized SPE (150s)")
cv_results$Condition <- factor(cv_results$Condition, levels = condition_levels)

# ------------------------------------------------------------------------------
# 3. Statistical Analysis & ANOVA
# ------------------------------------------------------------------------------
calc_stats <- function(df, var_name) {
  df %>%
    group_by(Condition) %>%
    summarise(
      Mean = mean(.data[[var_name]], na.rm = TRUE),
      SD = sd(.data[[var_name]], na.rm = TRUE)
    ) %>%
    rename_with(~paste0(var_name, "_", .), c(Mean, SD))
}

stats_oxidation <- calc_stats(cv_results, "I_pa")
stats_reduction <- calc_stats(cv_results, "I_pc")
stats_delta_ep <- calc_stats(cv_results, "Delta_Ep")

cat("\n--- ANOVA Summaries ---\n")
print(summary(aov(I_pa ~ Condition, data = cv_results)))
print(summary(aov(I_pc ~ Condition, data = cv_results)))
print(summary(aov(Delta_Ep ~ Condition, data = cv_results)))

# ------------------------------------------------------------------------------
# 4. Advanced Visualizations
# ------------------------------------------------------------------------------
# Palette mapping: Greys for Bare SPEs, Blues for Functionalized SPEs
palette_map <- c(
  "Bare SPE" = "#BDBDBD", "Bare SPE (60s)" = "#737373", "Bare SPE (150s)" = "#252525",
  "Functionalized SPE" = "#9ECAE1", "Functionalized SPE (60s)" = "#4292C6", "Functionalized SPE (150s)" = "#084594"
)

plot_box <- function(data, y_var, title, y_label) {
  ggplot(data, aes_string(x = "Condition", y = y_var, fill = "Condition")) +
    geom_boxplot(alpha = 0.85, outlier.shape = 21, outlier.size = 2) +
    labs(title = title, x = NULL, y = y_label) +
    scale_fill_manual(values = palette_map) +
    stat_compare_means(method = "anova", label.y = max(data[[y_var]]) * 1.1, size = 3.5) + 
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, hjust = 1, face = "bold"),
          legend.position = "none")
}

plot_ipa <- plot_box(cv_results, "I_pa", "Anodic Peak Current", expression(I[pa]~"(µA)"))
plot_ipc <- plot_box(cv_results, "I_pc", "Cathodic Peak Current", expression(I[pc]~"(µA)"))
plot_delta <- plot_box(cv_results, "Delta_Ep", "Peak Potential Separation", expression(Delta*E[p]~"(V)"))

# Arrange plots
final_grid <- ggarrange(plot_ipa, plot_ipc, plot_delta, ncol = 3, nrow = 1)
print(final_grid)

# ------------------------------------------------------------------------------
# 5. Export Consolidated Data
# ------------------------------------------------------------------------------
wb <- createWorkbook()

addWorksheet(wb, "Raw_Metrics")
writeData(wb, sheet = "Raw_Metrics", cv_results, rowNames = FALSE)

addWorksheet(wb, "Oxidation_Stats")
writeData(wb, sheet = "Oxidation_Stats", stats_oxidation, rowNames = FALSE)

addWorksheet(wb, "Reduction_Stats")
writeData(wb, sheet = "Reduction_Stats", stats_reduction, rowNames = FALSE)

addWorksheet(wb, "Delta_E_Stats")
writeData(wb, sheet = "Delta_E_Stats", stats_delta_ep, rowNames = FALSE)

export_path <- "./data/Surface_Modification_Results.xlsx"
saveWorkbook(wb, export_path, overwrite = TRUE)
message("Pipeline executed successfully. Results mapped to: ", export_path)