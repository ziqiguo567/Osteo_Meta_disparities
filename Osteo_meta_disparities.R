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
# Load Excel
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


# ---------------------------------------------------------------
# Meta + Forest
# ---------------------------------------------------------------

meta_meds_sex=metabin(male_meds, male_all, female_meds, female_all, studlab = Authors, sm= "OR", data=d_meds_sex)
forest(meta_meds_sex)

meta_dexa_sex=metabin(male_dexa, male_all, female_dexa, female_all, studlab = Authors, sm= "OR", data=d_dexa_sex)
forest(meta_dexa_sex)

meta_meds_race=metabin(black_meds, black_all, white_meds, white_all, studlab = Authors, sm="OR", data=d_meds_race)
forest(meta_meds_race)

meta_dexa_race=metabin(black_dexa, black_all, white_dexa, white_all, studlab = Authors, sm="OR", data=d_dexa_race)
forest(meta_dexa_race)


# -------------------------------
# output
# -------------------------------

if (!dir.exists("output")) dir.create("output")

# 1️⃣ Meds × Sex
pdf("output/Meds_Sex_forest.pdf", width = 14, height = 8)  # landscape
par(mar = c(5, 10, 5, 5))
forest(meta_meds_sex,
       leftlabs = c("Study", "Events", "Total"),
       xlim = c(0.01, 10),
       main = "Meds × Sex (Male vs Female)")
dev.off()

# 2️⃣ DEXA × Sex
pdf("output/Dexa_Sex_forest.pdf", width = 14, height = 8)
par(mar = c(5, 10, 5, 5))
forest(meta_dexa_sex,
       leftlabs = c("Study", "Events", "Total"),
       xlim = c(0.01, 10),
       main = "DEXA × Sex (Male vs Female)")
dev.off()

# 3️⃣ Meds × Race
pdf("output/Meds_Race_forest.pdf", width = 14, height = 8)
par(mar = c(5, 10, 5, 5))
forest(meta_meds_race,
       leftlabs = c("Study", "Events", "Total"),
       xlim = c(0.01, 10),
       main = "Meds × Race (Black vs White)")
dev.off()

# 4️⃣ DEXA × Race
pdf("output/Dexa_Race_forest.pdf", width = 14, height = 8)
par(mar = c(5, 10, 5, 5))
forest(meta_dexa_race,
       leftlabs = c("Study", "Events", "Total"),
       xlim = c(0.01, 10),
       main = "DEXA × Race (Black vs White)")
dev.off()

