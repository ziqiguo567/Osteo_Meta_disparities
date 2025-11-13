########################################################################
# Osteoporosis Meta-Analysis Pipeline
# Author: Jiamin Zhao & Ziqi Guo
# Outputs: Forest plots, 2×2 combined figure, summary tables
########################################################################

install.packages(c("readxl","dplyr","meta","writexl"))
installed.packages(c("ggplot2","cowplot"))
library(readxl)
library(dplyr)
library(meta)
library(writexl)
library(here)
library(ggplot2)
library(cowplot)

# ---------------------------------------------------------------
# 0. Load Excel
# ---------------------------------------------------------------
wb <- here("Osteo_updated_11072025.xlsx")
sheet_meds_sex  <- "Meds_Sex"
sheet_dexa_sex  <- "Dexa_Sex"
sheet_meds_race <- "Meds_Race"
sheet_dexa_race <- "Dexa_Race"

d_meds_sex  <- read_excel(wb, sheet = sheet_meds_sex)
d_dexa_sex  <- read_excel(wb, sheet = sheet_dexa_sex)
d_meds_race <- read_excel(wb, sheet = sheet_meds_race)
d_dexa_race <- read_excel(wb, sheet = sheet_dexa_race)

if (!dir.exists("output")) dir.create("output")

# ---------------------------------------------------------------
# 1. Function: compute logOR + SE from 2×2 table
# ---------------------------------------------------------------
calc_from_2x2 <- function(df, a, b, c, d) {
  df %>%
    mutate(
      logOR = log((!!sym(a) / !!sym(b)) / (!!sym(c) / !!sym(d))),
      SE = sqrt(1/!!sym(a) + 1/!!sym(b) + 1/!!sym(c) + 1/!!sym(d))
    )
}

# ---------------------------------------------------------------
# 2. Apply to each sheet (following YOUR Excel structure)
# ---------------------------------------------------------------

# Sex — Meds
m_meds_sex <- calc_from_2x2(
  d_meds_sex,
  a = "male_meds",
  b = "male_nonmeds",
  c = "female_meds",
  d = "female_nonmeds"
)

# Sex — Dexa
m_dexa_sex <- calc_from_2x2(
  d_dexa_sex,
  a = "male_dexa",
  b = "male_nondexa",
  c = "female_dexa",
  d = "female_nondexa"
)

# Race — Meds
m_meds_race <- calc_from_2x2(
  d_meds_race,
  a = "black_meds",
  b = "black_nonmeds",
  c = "white_meds",
  d = "white_nonmeds"
)

# Race — Dexa
m_dexa_race <- calc_from_2x2(
  d_dexa_race,
  a = "black_dexa",
  b = "black_nondexa",
  c = "white_dexa",
  d = "white_nondexa"
)

# ---------------------------------------------------------------
# 3. Meta-analysis
# ---------------------------------------------------------------
run_meta <- function(df) {
  metagen(
    TE = df$logOR,
    seTE = df$SE,
    studlab = df$Authors,
    method.tau = "REML",
    sm = "OR"
  )
}

fit_meds_sex  <- run_meta(m_meds_sex)
fit_dexa_sex  <- run_meta(m_dexa_sex)
fit_meds_race <- run_meta(m_meds_race)
fit_dexa_race <- run_meta(m_dexa_race)

# ---------------------------------------------------------------
# 4. Professional ggplot forest plot (blue theme)
# ---------------------------------------------------------------
make_forest_gg <- function(model, title){
  data_plot <- data.frame(
    Study = model$studlab,
    OR = exp(model$TE),
    lower = exp(model$lower),
    upper = exp(model$upper),
    Weight = model$w.random
  )
  
  ggplot(data_plot, aes(x = OR, y = reorder(Study, OR))) +
    geom_point(color = "#1f77b4", size = 3) +
    geom_errorbarh(aes(xmin = lower, xmax = upper),
                   height = 0.18, color = "#1f77b4") +
    geom_vline(xintercept = 1, linetype = "dashed", color = "gray40") +
    scale_x_log10() +
    labs(
      title = title,
      x = "Odds Ratio (log scale)",
      y = "Study"
    ) +
    theme_minimal(base_size = 13) +
    theme(
      plot.title = element_text(size = 16, face = "bold"),
      panel.grid.minor = element_blank()
    )
}

p1 <- make_forest_gg(fit_meds_sex,  "Sex — Meds")
p2 <- make_forest_gg(fit_dexa_sex,  "Sex — Dexa")
p3 <- make_forest_gg(fit_meds_race, "Race — Meds")
p4 <- make_forest_gg(fit_dexa_race, "Race — Dexa")

# ---------------------------------------------------------------
# 5. Save individual HD plots
# ---------------------------------------------------------------
ggsave(here("output","forest_meds_sex.png"),  p1, width=7, height=6, dpi=320)
ggsave(here("output","forest_dexa_sex.png"),  p2, width=7, height=6, dpi=320)
ggsave(here("output","forest_meds_race.png"), p3, width=7, height=6, dpi=320)
ggsave(here("output","forest_dexa_race.png"), p4, width=7, height=6, dpi=320)

# ---------------------------------------------------------------
# 6. Create 2x2 combined figure
# ---------------------------------------------------------------
final_plot <- plot_grid(
  p1, p2,
  p3, p4,
  labels = NULL,
  ncol = 2,
  align = "hv"
)

ggsave(
  filename = here("output","meta_2x2_forest.png"),
  plot = final_plot,
  width = 14, height = 12, dpi = 320
)

# ---------------------------------------------------------------
# 7. Summary table (pooled OR + CI + I2 + Tau2)
# ---------------------------------------------------------------
extract_summary <- function(model, label){
  tibble(
    Group = label,
    Pooled_OR = exp(model$TE.random),
    CI_low   = exp(model$lower.random),
    CI_high  = exp(model$upper.random),
    Tau2     = model$tau^2,
    I2       = model$I2
  )
}

summary_table <- bind_rows(
  extract_summary(fit_meds_sex,  "Sex — Meds"),
  extract_summary(fit_dexa_sex,  "Sex — Dexa"),
  extract_summary(fit_meds_race, "Race — Meds"),
  extract_summary(fit_dexa_race, "Race — Dexa")
)

write.csv(summary_table, here("output","meta_summary_table.csv"), row.names = FALSE)

########################################################################
# End of Pipeline
########################################################################