library(tidyverse)
library(rstatix)
library(officer)
library(flextable)
library(stringr)
library(purrr)
library(rio)

raw <- wound_raw_prism

long <- raw %>%
  pivot_longer(R1:R5, names_to = "Rat", values_to = "Strength") %>%
  mutate(Ointment = factor(Ointment,
                           levels = c('a','b','c','d','e','f','g','h','i','j','k')))

summary_stats <- long %>%
  group_by(Ointment) %>%
  summarise(
    Mean = mean(Strength),
    SEM  = sd(Strength)/sqrt(n()),
    .groups = "drop"
  )
untreated_mean <- summary_stats$Mean[summary_stats$Ointment == "a"]

summary_stats <- summary_stats %>%
  mutate(
    Percent_Increase = ifelse(
      Ointment == "a",
      NA,
      ((Mean - untreated_mean)/untreated_mean) * 100
    )
  )
anova_model <- aov(Strength ~ Ointment, data = long)
tukey_res   <- tukey_hsd(long, Strength ~ Ointment)

p_to_symbol <- function(p){
  if(p < 0.001) return("³")
  if(p < 0.01)  return("²")
  if(p < 0.05)  return("¹")
  return("")
}

tukey_sig <- tukey_res %>%
  mutate(symbol = map_chr(p.adj, p_to_symbol)) %>%
  filter(symbol != "") %>%
  transmute(
    Ointment = group2,
    sup_part = paste0(group1, symbol)   # a³, b² etc
  )

tukey_supers <- tukey_sig %>%
  group_by(Ointment) %>%
  summarise(
    sup = paste0(sup_part, collapse = ""),
    .groups = "drop"
  )


final_stats <- summary_stats %>%
  left_join(tukey_supers, by = "Ointment") %>%
  mutate(
    sup = ifelse(is.na(sup), "", sup),
    Cell = sprintf("%.1f ± (%.2f)%s", Mean, SEM, sup)
  )


plos_table <- final_stats %>%
  mutate(
    Percent_Increase = ifelse(is.na(Percent_Increase), "-",
                              sprintf("%.2f", Percent_Increase))
  ) %>%
  select(
    Group = Ointment,
    `Mean tensile strength ± SEM` = Cell,
    `Percent increase of tensile strength` = Percent_Increase
  )

plos_table$Group <- recode(plos_table$Group,
                           a = "Untreated",
                           b = "Simple ointment",
                           c = "5% (w/w) extract ointment",
                           d = "10% (w/w) extract ointment",
                           e = "5% CF ointment",
                           f = "10% CF ointment",
                           g = "5% NHF ointment",
                           h = "10% NHF ointment",
                           i = "5% AQF ointment",
                           j = "10% AQF ointment",
                           k = "0.5% NS+B ointment"
)
export(plos_table,'tensile_analysis.xlsx')
