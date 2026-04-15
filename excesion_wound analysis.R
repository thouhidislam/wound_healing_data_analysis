# ================================
# 0. Packages
# ================================
library(tidyverse)
library(rstatix)
library(officer)
library(flextable)
library(stringr)
library(purrr)
library(rio)

# ================================
# 1. Input data
# ================================
raw <- wound_raw_prism   # your raw table

# ================================
# 2. Long format (rats = replicates)
# ================================
long <- raw %>%
  pivot_longer(R1:R5, names_to = "Rat", values_to = "Area") %>%
  mutate(
    Ointment = factor(Ointment, levels = c("Simple Ointment","5% CF(w/w) Ointment","10% CF(w/w) Ointment",
                                           "5% NHF(w/w) Ointment", "10% NHF(w/w) Ointment","5% AQF(w/w) Ointment",
                                           "10% AQF(w/w) Ointment","0.5% NS+B(w/w) Ointment" ))
  )
list(raw$Ointment)
export(long,'long.xlsx')
# ================================
# 3. Mean ± SEM
# ================================
summary_stats <- long %>%
  group_by(Day, Ointment) %>%
  summarise(
    Mean = mean(Area, na.rm = TRUE),
    SEM  = sd(Area, na.rm = TRUE) / sqrt(n()),
    .groups = "drop"
  )


# ================================
# 4. % wound contraction (baseline Day 2)
# ================================
baseline_wrong <- summary_stats %>%
  filter(Day == 2) %>%
  select(Ointment, Baseline = Mean)
baseline <- data.frame(Ointment=c("Simple Ointment","5% CF(w/w) Ointment","10% CF(w/w) Ointment",
                                  "5% NHF(w/w) Ointment", "10% NHF(w/w) Ointment","5% AQF(w/w) Ointment",
                                  "10% AQF(w/w) Ointment","0.5% NS+B(w/w) Ointment"),
                       Baseline=300)

summary_stats <- summary_stats %>%
  left_join(baseline, by = "Ointment") %>%
  mutate(
    Contraction = ((Baseline - Mean) / Baseline) * 100
  )
# ================================
# 5. One-way ANOVA per day
# ================================
anova_res <- long %>%
  group_by(Day) %>%
  anova_test(Area ~ Ointment)

# ================================
# 6. Tukey HSD per day
# ================================
tukey_res <- long %>%
  group_by(Day) %>%
  tukey_hsd(Area ~ Ointment)
summary(tukey_res)
export(tukey_res,'tukey_anova.xlsx')
a <- tukey_res
# ================================
# 7. Prism superscript rules
# ================================
p_to_symbol <- function(p){
  if(p < 0.001) return("³")
  if(p < 0.01)  return("²")
  if(p < 0.05)  return("¹")
  return("")
}

tukey_clean <- tukey_res %>%
  mutate(
    comp = paste(group1, group2, sep = "_"),
    symbol = map_chr(p.adj, p_to_symbol),
    prefix = case_when(
      str_detect(comp, fixed("Simple Ointment"))            ~ "a",
      str_detect(comp, fixed("5% CF(w/w) Ointment"))        ~ "b",
      str_detect(comp, fixed("10% CF(w/w) Ointment"))       ~ "c",
      str_detect(comp, fixed("5% NHF(w/w) Ointment"))       ~ "d",
      str_detect(comp, fixed("10% NHF(w/w) Ointment"))      ~ "e",
      str_detect(comp, fixed("5% AQF(w/w) Ointment"))       ~ "f",
      str_detect(comp,fixed("10% AQF(w/w) Ointment"))       ~"g",
      TRUE ~ ""
    ),
    sup = ifelse(symbol == "" | prefix == "", "", paste0(prefix, symbol))
  )

tukey_long_1 <- tukey_clean %>%
  pivot_longer(c(group1, group2), names_to = "side", values_to = "Ointment") %>% 
  filter(side != "group1") %>% 
  select(Day, Ointment, sup)

# ================================
# 8. Final formatted cell:
# Mean ± SEM (% contraction) + superscript
# ================================
final_stats_1 <- summary_stats %>%
  mutate(
    Cell = sprintf(
      "%.2f ± %.2f (%.2f)",
      Mean, SEM, Contraction
    )
  ) %>%
  left_join(tukey_long_1, by = c("Day", "Ointment")) %>%
  mutate(
    Cell = ifelse(is.na(sup), Cell, paste0(Cell, sup))
  )
export(final_stats,'final_stats.xlsx')

# Load required packages
library(tidyverse)
library(openxlsx)

# Assuming your raw data is called 'df' with columns:
# Day, Ointment, Mean, SEM, Contraction, sup

plos_table <- final_stats_1 %>%
  # Step 1: Summarise replicates per Day & Ointment
  group_by(Day, Ointment) %>%
  summarise(
    Mean = mean(Mean),
    SEM = mean(SEM),
    Contraction = mean(Contraction),
    sup = paste(unique(sup), collapse = ""),
    .groups = "drop"
  ) %>%
  # Step 1b: Replace NA in 'sup' with empty string
  mutate(sup = ifelse(is.na(sup) | sup == "NA", "", sup)) %>%
  # Step 2: Format as PLOS-style string
  mutate(
    formatted = paste0(
      round(Mean, 2), " ± ", round(SEM, 2),
      " (", round(Contraction, 2), ")", sup
    )
  ) %>%
  select(Day, Ointment, formatted) %>%
  # Step 3: Pivot to wide
  pivot_wider(names_from = Ointment, values_from = formatted) %>%
  # Step 4: Rename (optional if names are already correct)
  rename(
    `Simple Ointment` = `Simple Ointment`,
    `5% CF(w/w) Ointment` = `5% CF(w/w) Ointment`,
    `10% CF(w/w) Ointment` = `10% CF(w/w) Ointment`,
    `5% NHF(w/w) Ointment` = `5% NHF(w/w) Ointment`,
    `10% NHF(w/w) Ointment` = `10% NHF(w/w) Ointment`,
    `5% AQF(w/w) Ointment` = `5% AQF(w/w) Ointment`,
    `10% AQF(w/w) Ointment` = `10% AQF(w/w) Ointment`,
    `0.5% NS+B(w/w) Ointment` = `0.5% NS+B(w/w) Ointment`
  ) %>%
  # Step 5: Reorder columns
  select(
    Day,
    `Simple Ointment`,
    `5% CF(w/w) Ointment`,
    `10% CF(w/w) Ointment`,
    `5% NHF(w/w) Ointment`,
    `10% NHF(w/w) Ointment`,
    `5% AQF(w/w) Ointment`,
    `10% AQF(w/w) Ointment`,
    `0.5% NS+B(w/w) Ointment`
  ) %>%
  arrange(Day)


# Step 5: Export to Excel (optional)
export(plos_table, "WoundHealing_excision_PLOS_Table.xlsx")

# View final table
plos_table


library(flextable)
library(officer)

# Create Word document
doc <- read_docx()

ft <- flextable(plos_table) %>%
  theme_booktabs() %>%
  align(align = "center", part = "all") %>%
  autofit()

doc <- doc %>%
  body_add_par(
    "Table 2. Wound healing activity of extract of the leaves of Clerodendrum myricoides excision wound model in mice.",
    style = "heading 2"
  ) %>%
  body_add_flextable(ft) %>%
  body_add_par(
    "Values are expressed as mean ± SEM (n = 6) with percentage wound contraction in parentheses.
One-way ANOVA followed by Tukey’s multiple comparison test.
a compared to simple ointment; b compared to 5% (w/w) extract ointment.
P < 0.05, P < 0.01, P < 0.001. Initial wound area was 300 mm².",
    style = "Normal"
  )




# Save the Word document
print(doc, target = "Table1_PLOS_Style.docx")
