library(tidyverse)
library(rstatix)
library(officer)
library(flextable)
library(stringr)
library(purrr)
library(rio)

# Required packages
library(tidyverse)
library(rstatix)
library(writexl)

# Raw data
raw <- wound_raw_prism

# Long format
long <- raw %>%
  pivot_longer(r1:r5, names_to = "Rat", values_to = "Strength") %>%
  mutate(
    Ointment = factor(Ointment,
                      levels = c("a","b","c","d"))
  )

# Mean ± SEM
summary_stats <- long %>%
  group_by(Ointment) %>%
  summarise(
    Mean = mean(Strength),
    SEM  = sd(Strength) / sqrt(n()),
    .groups = "drop"
  )

# Mean of simple ointment (reference)
mean_simple <- summary_stats %>%
  filter(Ointment == "a") %>%
  pull(Mean)

# % decrease in epithelialization period
summary_stats <- summary_stats %>%
  mutate(
    Percent_Reduction = ifelse(
      Ointment == "a",
      0,
      ((mean_simple - Mean) / mean_simple) * 100
    )
  )

# One-way ANOVA
anova_model <- aov(Strength ~ Ointment, data = long)

# Tukey post hoc
tukey_res <- long %>%
  tukey_hsd(Strength ~ Ointment)

# p-value → superscript mapping
p_to_symbol <- function(p){
  if (p < 0.001) return("³")
  if (p < 0.01)  return("²")
  if (p < 0.05)  return("¹")
  return("")
}

# Extract significant comparisons
tukey_sig <- tukey_res %>%
  mutate(symbol = map_chr(p.adj, p_to_symbol)) %>%
  filter(symbol != "") %>%
  transmute(
    Ointment = group2,
    sup_part = paste0(group1, symbol)
  )

# Combine superscripts
tukey_supers <- tukey_sig %>%
  group_by(Ointment) %>%
  summarise(
    sup = paste0(sup_part, collapse = ""),
    .groups = "drop"
  )

# Final stats with superscripts
final_stats <- summary_stats %>%
  left_join(tukey_supers, by = "Ointment") %>%
  mutate(
    sup = replace_na(sup, ""),
    Cell = sprintf("%.2f ± %.2f%s", Mean, SEM, sup)
  )

# Final publication-ready table
plos_table <- final_stats %>%
  mutate(
    `% decrease in the period of epithelialization` =
      ifelse(Ointment == "a", "-", sprintf("%.2f", Percent_Reduction))
  ) %>%
  select(
    Group = Ointment,
    `Mean period of epithelialization (days) ± SEM` = Cell,
    `% decrease in the period of epithelialization`
  )

# Recode group names
plos_table$Group <- recode(
  plos_table$Group,
  a = "Simple ointment",
  b = "5% CF ointment",
  c = "10% CF ointment",
  d = "5% NHF ointment",
  e = "10% NHF ointment",
  f = "5% AQF ointment",
  g = "10% AQF ointment",
  h = "0.2% Nitrofurazone ointment"
)

# Export to Excel


