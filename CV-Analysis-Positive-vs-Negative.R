# ==============================================================================
# Cyclic Voltammetry (CV) Peak Analysis 
# Automated extraction of redox peaks, delta E, and statistical analysis
# ==============================================================================

library(readxl)
library(dplyr)
library(ggplot2)
library(ggpubr)
library(openxlsx)

# ------------------------------------------------------------------------------
# 1. Data Import and Cleaning
# ------------------------------------------------------------------------------
data_path <- "./data/CV_Output.xlsx"
raw_data <- read_excel(data_path)

orig_names <- colnames(raw_data)
new_names <- character(length(orig_names))

# Standardize columns to POTENTIAL_ and CURRENT_
for (i in seq_along(orig_names)) {
  if (i %% 2 == 1) {  
    new_names[i] <- paste0("POTENTIAL_", orig_names[i])
  } else {            
    prev_name <- new_names[i-1]  
    new_names[i] <- sub("POTENTIAL_", "CURRENT_", prev_name)  
  }
}

colnames(raw_data) <- new_names
raw_data <- raw_data[-1, ] # Remove unit/metadata row

col_current <- seq(2, ncol(raw_data), by = 2)
raw_data[col_current] <- lapply(raw_data[col_current], as.numeric)

col_potential <- seq(1, ncol(raw_data), by = 2)
raw_data[col_potential] <- lapply(raw_data[col_potential], as.numeric)

# ------------------------------------------------------------------------------
# 2. Extract Electrochemical Parameters (I_pa, I_pc, Delta E)
# ------------------------------------------------------------------------------
potential_cols <- grep("POTENTIAL_", names(raw_data), value = TRUE)
current_cols <- grep("CURRENT_", names(raw_data), value = TRUE)
cv_results <- data.frame()

for (i in seq_along(potential_cols)) {
  potential_vec <- raw_data[[potential_cols[i]]]
  current_vec <- raw_data[[current_cols[i]]]
  
  # Extract Peak Currents (Anodic/Oxidation and Cathodic/Reduction)
  i_pa <- max(current_vec, na.rm = TRUE) # Anodic peak current
  i_pc <- min(current_vec, na.rm = TRUE) # Cathodic peak current
  
  # Extract corresponding Peak Potentials
  e_pa <- potential_vec[which.max(current_vec)]
  e_pc <- potential_vec[which.min(current_vec)]
  
  # Calculate Peak Potential Separation (Delta E_p)
  delta_ep <- abs(e_pa - e_pc)
  
  # Map Experimental Conditions
  condition <- if(grepl("Positivo", current_cols[i], ignore.case = TRUE)) {
    "Positive"
  } else if(grepl("Negativo", current_cols[i], ignore.case = TRUE)) {
    "Negative"
  } else {
    "Ultrapure Water"
  }
  
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

cv_results$Condition <- as.factor(cv_results$Condition)

# ------------------------------------------------------------------------------
# 3. Statistical Summaries (ANOVA)
# ------------------------------------------------------------------------------
stats_oxidation <- cv_results %>%
  group_by(Condition) %>%
  summarise(
    Mean_I_pa = mean(I_pa, na.rm = TRUE),
    SD_I_pa = sd(I_pa, na.rm = TRUE),
    Mean_E_pa = mean(E_pa, na.rm = TRUE),
    SD_E_pa = sd(E_pa, na.rm = TRUE)
  )

stats_reduction <- cv_results %>%
  group_by(Condition) %>%
  summarise(
    Mean_I_pc = mean(I_pc, na.rm = TRUE),
    SD_I_pc = sd(I_pc, na.rm = TRUE),
    Mean_E_pc = mean(E_pc, na.rm = TRUE),
    SD_E_pc = sd(E_pc, na.rm = TRUE)
  )

stats_delta_ep <- cv_results %>%
  group_by(Condition) %>%
  summarise(
    Mean_Delta_Ep = mean(Delta_Ep, na.rm = TRUE),
    SD_Delta_Ep = sd(Delta_Ep, na.rm = TRUE)
  )

# Display ANOVA summaries to the console
cat("\n--- ANOVA: Anodic Peak Current (Oxidation) ---\n")
print(summary(aov(I_pa ~ Condition, data = cv_results)))

cat("\n--- ANOVA: Cathodic Peak Current (Reduction) ---\n")
print(summary(aov(I_pc ~ Condition, data = cv_results)))

cat("\n--- ANOVA: Peak Separation (Delta Ep) ---\n")
print(summary(aov(Delta_Ep ~ Condition, data = cv_results)))

# ------------------------------------------------------------------------------
# 4. Visualizations
# ------------------------------------------------------------------------------
# Define consistent color palette
palette_colors <- c("Positive" = "#E41A1C", "Negative" = "#377EB8", "Ultrapure Water" = "#4DAF4A")

# Boxplot: Oxidation Current (I_pa)
plot_ipa <- ggplot(cv_results, aes(x = Condition, y = I_pa, fill = Condition)) +
  geom_boxplot(alpha = 0.8) +
  labs(title = "Anodic Peak Current (Oxidation)", x = "Condition", y = expression(I[pa]~"(µA)")) +
  scale_fill_manual(values = palette_colors) +
  stat_compare_means(method = "anova", label.y = max(cv_results$I_pa) * 1.1) + 
  theme_minimal() + theme(legend.position = "none")

# Boxplot: Reduction Current (I_pc)
plot_ipc <- ggplot(cv_results, aes(x = Condition, y = I_pc, fill = Condition)) +
  geom_boxplot(alpha = 0.8) +
  labs(title = "Cathodic Peak Current (Reduction)", x = "Condition", y = expression(I[pc]~"(µA)")) +
  scale_fill_manual(values = palette_colors) +
  stat_compare_means(method = "anova", label.y = max(cv_results$I_pc) * 1.1) + 
  theme_minimal() + theme(legend.position = "none")

# Boxplot: Peak Separation (Delta E_p)
plot_delta <- ggplot(cv_results, aes(x = Condition, y = Delta_Ep, fill = Condition)) +
  geom_boxplot(alpha = 0.8) +
  labs(title = "Peak Potential Separation", x = "Condition", y = expression(Delta*E[p]~"(V)")) +
  scale_fill_manual(values = palette_colors) +
  stat_compare_means(method = "anova", label.y = max(cv_results$Delta_Ep) * 1.1) + 
  theme_minimal() + theme(legend.position = "none")

# Combine plots for a professional overview
combined_plots <- ggarrange(plot_ipa, plot_ipc, plot_delta, ncol = 3, nrow = 1)
print(combined_plots)

# ------------------------------------------------------------------------------
# 5. Export Results
# ------------------------------------------------------------------------------
wb <- createWorkbook()

addWorksheet(wb, "CV_Results")
writeData(wb, sheet = "CV_Results", cv_results, rowNames = FALSE)

addWorksheet(wb, "Oxidation_Stats")
writeData(wb, sheet = "Oxidation_Stats", stats_oxidation, rowNames = FALSE)

addWorksheet(wb, "Reduction_Stats")
writeData(wb, sheet = "Reduction_Stats", stats_reduction, rowNames = FALSE)

addWorksheet(wb, "Delta_E_Stats")
writeData(wb, sheet = "Delta_E_Stats", stats_delta_ep, rowNames = FALSE)

export_path <- "./data/CV_Consolidated_Results.xlsx"
saveWorkbook(wb, export_path, overwrite = TRUE)
message("Analysis complete. Results saved to: ", export_path)