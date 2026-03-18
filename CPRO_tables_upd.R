# =============================================================================
# CPRO Snapshot 2019-2024: Reproducing Tables 1-11
# Data sources: 2019-2024.13.March.2026.xlsx (full dataset)
#               2019-2024 50+.xlsx (high-engagement subset)
# =============================================================================

library(tidyverse)
library(readxl)
library(gt)

# --- Common table style: Times New Roman, 12 pt, with source note -----------
style_table <- function(gt_obj, source_text) {
  gt_obj %>%
    tab_source_note(source_note = source_text) %>%
    opt_table_font(font = list("Times New Roman")) %>%
    tab_options(table.font.size = px(12))
}

# --- File paths (adjust if needed) -----------------------------------------
# FIX: rstudioapi::getSourceEditorContext() throws an error outside RStudio,
#      so we wrap in tryCatch instead of checking NULL after the fact.
dir_path <- tryCatch(
  dirname(rstudioapi::getSourceEditorContext()$path),
  error = function(e) getwd()
)
if (is.null(dir_path) || dir_path == "") {
  dir_path <- getwd()
}

full_file <- file.path(dir_path, "2019-2024.13.March.2026.xlsx")
high_file <- file.path(dir_path, "2019-2024 50+.xlsx")

# --- Load data --------------------------------------------------------------
full <- read_excel(full_file, sheet = "Draft Legal Acts") %>%
  rename(
    title       = `Regulation Title`,
    id          = ID,
    developer   = Developer,
    type        = `Type of Regulation`,
    date_post   = `Date of Posting`,
    date_end    = `End of Posting`,
    comments    = `Number of Comments`,
    sphere      = Sphere
  ) %>%
  mutate(
    year = year(date_post),
    comments = as.numeric(comments)
  ) %>%
  filter(year != 2025)

high <- read_excel(high_file, sheet = "Draft Legal Acts") %>%
  rename(
    title        = `Regulation Title`,
    id           = ID,
    developer    = Developer,
    type         = `Type of Regulation`,
    date_post    = `Date of Posting`,
    date_end     = `End of Posting`,
    comments     = `Number of Comments`,
    ria          = `RIA Report (1/0)`,
    response     = `Response from the Developer (1/0)`,
    adopted      = `Adopted (1/0)`,
    adopted_link = `Adopted Regulation Link`,
    grounds      = `Grounds for Developing Regulation (Self-Initiated / Regulation)`,
    ground_link  = `Ground Regulation Link`,
    sphere       = Sphere
  ) %>%
  mutate(
    year = year(date_post),
    comments = as.numeric(comments)
  ) %>%
  filter(year != 2025)

# =============================================================================
# TABLE 1: Annual Overview of Public Consultation Activity
# =============================================================================
table1 <- full %>%
  group_by(year) %>%
  summarise(
    posted         = n(),
    total_comments = sum(comments, na.rm = TRUE),
    mean_comments  = round(total_comments / posted, 1),
    zero_comment   = sum(comments == 0, na.rm = TRUE),
    with_comments  = sum(comments > 0, na.rm = TRUE),
    engage_rate    = with_comments / posted,
    .groups = "drop"
  )

table1_total <- table1 %>%
  summarise(
    year           = "Total",
    posted         = sum(posted),
    total_comments = sum(total_comments),
    mean_comments  = round(total_comments / posted, 1),
    zero_comment   = sum(zero_comment),
    with_comments  = sum(with_comments),
    engage_rate    = with_comments / posted
  )

table1_display <- table1 %>%
  mutate(year = as.character(year)) %>%
  bind_rows(table1_total) %>%
  mutate(
    posted         = format(posted, big.mark = ","),
    total_comments = format(total_comments, big.mark = ","),
    engage_rate    = paste0(round(engage_rate * 100, 1), "%")
  )

gt_table1 <- table1_display %>%
  gt() %>%
  cols_label(
    year           = "Year",
    posted         = "Posted",
    total_comments = "Total Comments",
    mean_comments  = "Mean",
    zero_comment   = "Zero Comment",
    with_comments  = "With Comments",
    engage_rate    = "Engage. Rate"
  ) %>%
  tab_header(
    title = "Table 1. Annual Overview of Public Consultation Activity on regulation.gov.uz, 2019-2024")

print(gt_table1)

# =============================================================================
# TABLE 2: Distribution of Public Comments by Engagement Bracket
# =============================================================================
brackets <- c(-Inf, 0, 5, 10, 50, 100, 500, 1000, Inf)
labels   <- c("0", "1-5", "6-10", "11-50", "51-100", "101-500", "501-1,000", "1,000+")

table2 <- full %>%
  # FIX: drop NAs before cut() so bracket counts match the full dataset total
  filter(!is.na(comments)) %>%
  mutate(bracket = cut(comments, breaks = brackets, labels = labels, right = TRUE)) %>%
  group_by(bracket) %>%
  summarise(
    n_regs         = n(),
    total_comments = sum(comments, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(
    pct_regs     = n_regs / sum(n_regs),
    pct_comments = total_comments / sum(total_comments)
  )

gt_table2 <- table2 %>%
  mutate(
    pct_regs     = paste0(round(pct_regs * 100, 1), "%"),
    pct_comments = paste0(round(pct_comments * 100, 1), "%"),
    n_regs       = format(n_regs, big.mark = ","),
    total_comments = format(total_comments, big.mark = ",")
  ) %>%
  gt() %>%
  cols_label(
    bracket        = "Comment Bracket",
    n_regs         = "No. of Regs",
    pct_regs       = "% of All Regs",
    total_comments = "Total Comments",
    pct_comments   = "% of Comments"
  ) %>%
  tab_header(
    title = "Table 2. Distribution of Public Comments by Engagement Bracket, 2019-2024")

print(gt_table2)

# =============================================================================
# TABLE 3: Public Engagement by Type of Regulation
# =============================================================================
# FIX: removed unused `other_types` vector -- the major types are listed
#      directly in the if_else() below; everything else maps to "Other".

table3 <- full %>%
  mutate(
    type_grouped = if_else(
      type %in% c("Government Resolution", "Ministry / Agency Resolution",
                   "Presidential Resolution", "Law", "Presidential Decree"),
      type,
      "Other"
    )
  ) %>%
  group_by(type_grouped) %>%
  summarise(
    count       = n(),
    total_cmts  = sum(comments, na.rm = TRUE),
    mean_cmts   = round(total_cmts / count, 1),
    with_cmts   = sum(comments > 0, na.rm = TRUE),
    zero_cmts   = sum(comments == 0, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(
    engage_rate = with_cmts / count,
    pct_zero    = zero_cmts / count,
    pct_all     = count / sum(count)
  )

# Order to match the Word document
type_order <- c("Government Resolution", "Ministry / Agency Resolution",
                "Presidential Resolution", "Law", "Presidential Decree", "Other")

gt_table3 <- table3 %>%
  mutate(type_grouped = factor(type_grouped, levels = type_order)) %>%
  arrange(type_grouped) %>%
  mutate(
    count       = format(count, big.mark = ","),
    total_cmts  = format(total_cmts, big.mark = ","),
    engage_rate = paste0(round(engage_rate * 100, 1), "%"),
    pct_zero    = paste0(round(pct_zero * 100, 1), "%"),
    pct_all     = paste0(round(pct_all * 100, 1), "%")
  ) %>%
  select(type_grouped, count, total_cmts, mean_cmts, engage_rate, pct_zero, pct_all) %>%
  gt() %>%
  cols_label(
    type_grouped = "Type of Regulation",
    count        = "Count",
    total_cmts   = "Total Cmts",
    mean_cmts    = "Mean",
    engage_rate  = "Engage. Rate",
    pct_zero     = "% Zero",
    pct_all      = "% of All Regs"
  ) %>%
  tab_header(
    title = "Table 3. Public Engagement by Type of Regulation, 2019-2024")

print(gt_table3)

# =============================================================================
# TABLE 4: Top Five Developing Agencies by Year
# =============================================================================
table4 <- full %>%
  group_by(year, developer) %>%
  summarise(n = n(), .groups = "drop") %>%
  group_by(year) %>%
  slice_max(order_by = n, n = 5, with_ties = FALSE) %>%
  mutate(
    rank  = row_number(),
    label = paste0(developer, " (", n, ")")
  ) %>%
  ungroup() %>%
  select(year, rank, label) %>%
  pivot_wider(names_from = year, values_from = label) %>%
  select(-rank)

gt_table4 <- table4 %>%
  gt() %>%
  tab_header(
    title = "Table 4. Top Five Developing Agencies by Year, 2019-2024")

print(gt_table4)

# =============================================================================
# TABLE 5: Developer Participation Dynamics
# =============================================================================
devs_by_year <- full %>%
  group_by(year) %>%
  summarise(devs = list(unique(developer)), .groups = "drop")

table5_rows <- list()
for (i in seq_len(nrow(devs_by_year))) {
  y    <- devs_by_year$year[i]
  curr <- devs_by_year$devs[[i]]

  # HHI
  shares <- full %>%
    filter(year == y) %>%
    count(developer) %>%
    mutate(share = n / sum(n))
  hhi <- round(sum(shares$share^2), 3)

  if (i == 1) {
    table5_rows[[i]] <- tibble(
      year = y, active = length(curr),
      new_entrants = "---", exited = "---", continuing = "---",
      hhi = hhi
    )
  } else {
    prev <- devs_by_year$devs[[i - 1]]
    new_e <- setdiff(curr, prev)
    exit  <- setdiff(prev, curr)
    cont  <- intersect(curr, prev)
    table5_rows[[i]] <- tibble(
      year = y, active = length(curr),
      new_entrants = as.character(length(new_e)),
      exited       = as.character(length(exit)),
      continuing   = as.character(length(cont)),
      hhi = hhi
    )
  }
}

table5 <- bind_rows(table5_rows)

gt_table5 <- table5 %>%
  gt() %>%
  cols_label(
    year         = "Year",
    active       = "Active Developers",
    new_entrants = "New Entrants",
    exited       = "Exited",
    continuing   = "Continuing",
    hhi          = "HHI"
  ) %>%
  tab_header(
    title = "Table 5. Developer Participation Dynamics, 2019-2024")

print(gt_table5)

# =============================================================================
# TABLE 6: Comment Intensity by Leading Regulatory Sphere
# =============================================================================
leading_spheres <- c("Education, Science and Culture", "Business", "Banking")

table6 <- full %>%
  filter(sphere %in% leading_spheres) %>%
  group_by(sphere, year) %>%
  summarise(
    n_regs = n(),
    mean_c = round(mean(comments, na.rm = TRUE), 1),
    .groups = "drop"
  )

# Pivot to match the Word table format (Sphere | Metric | 2019 | 2020 | ...)
table6_n <- table6 %>%
  select(sphere, year, n_regs) %>%
  pivot_wider(names_from = year, values_from = n_regs) %>%
  mutate(metric = "N regs") %>%
  relocate(sphere, metric)

table6_m <- table6 %>%
  select(sphere, year, mean_c) %>%
  pivot_wider(names_from = year, values_from = mean_c) %>%
  mutate(metric = "Mean") %>%
  relocate(sphere, metric)

table6_display <- bind_rows(table6_n, table6_m) %>%
  mutate(sphere = factor(sphere, levels = leading_spheres)) %>%
  arrange(sphere, desc(metric == "N regs"))

gt_table6 <- table6_display %>%
  gt() %>%
  cols_label(
    sphere = "Sphere",
    metric = "Metric"
  ) %>%
  tab_header(
    title = "Table 6. Comment Intensity by Leading Regulatory Sphere, 2019-2024")

print(gt_table6)

# =============================================================================
# TABLE 7: Temporal Profile of the High-Engagement Subset (50+ Comments)
# =============================================================================
# Use main dataset filtered to 50+ for consistency
high_from_full <- full %>% filter(comments >= 50)

table7 <- high_from_full %>%
  group_by(year) %>%
  summarise(
    count_50plus   = n(),
    total_comments = sum(comments, na.rm = TRUE),
    mean_comments  = round(mean(comments, na.rm = TRUE)),
    .groups = "drop"
  ) %>%
  # Median quintile: based on brackets 51-100=4, 101-500=5, 501-1000=6, 1000+=7
  left_join(
    high_from_full %>%
      mutate(
        quintile = case_when(
          comments <= 100  ~ 4,
          comments <= 500  ~ 5,
          comments <= 1000 ~ 6,
          TRUE             ~ 7
        )
      ) %>%
      group_by(year) %>%
      summarise(median_quintile = median(quintile), .groups = "drop"),
    by = "year"
  ) %>%
  left_join(
    full %>% group_by(year) %>% summarise(total_regs = n(), .groups = "drop"),
    by = "year"
  ) %>%
  mutate(
    pct_year = count_50plus / total_regs
  )

gt_table7 <- table7 %>%
  mutate(
    total_comments = format(total_comments, big.mark = ","),
    pct_year       = paste0(round(pct_year * 100, 1), "%"),
    # FIX: use round() before as.integer() so e.g. median 4.5 becomes 5, not 4
    median_quintile = as.integer(round(median_quintile))
  ) %>%
  select(year, count_50plus, total_comments, mean_comments, median_quintile, pct_year) %>%
  gt() %>%
  cols_label(
    year            = "Year",
    count_50plus    = "50+ Count",
    total_comments  = "Total Comments",
    mean_comments   = "Mean Comments",
    median_quintile = "Median Quintile",
    pct_year        = "% of Year's Regs"
  ) %>%
  tab_header(
    title = "Table 7. Temporal Profile of the High-Engagement Subset (50+ Comments), 2019-2024")

print(gt_table7)

# =============================================================================
# TABLE 8: Sectoral Composition: Full Dataset vs. High-Engagement Subset
# =============================================================================
full_sphere <- full %>%
  count(sphere, name = "full_n") %>%
  mutate(full_pct = full_n / sum(full_n))

high_sphere <- high_from_full %>%
  group_by(sphere) %>%
  summarise(
    high_n    = n(),
    high_cmts = sum(comments, na.rm = TRUE),
    high_mean = round(mean(comments, na.rm = TRUE), 1),
    .groups = "drop"
  ) %>%
  mutate(high_pct = high_n / sum(high_n))

table8 <- full_sphere %>%
  full_join(high_sphere, by = "sphere") %>%
  replace_na(list(high_n = 0, high_cmts = 0, high_mean = 0, high_pct = 0)) %>%
  mutate(
    over_rep = round(high_pct / full_pct, 2)
  ) %>%
  arrange(desc(high_n))

# Show top spheres + "All Other" row
top_spheres <- table8 %>% slice_head(n = 9)
other_row <- table8 %>%
  slice_tail(n = nrow(table8) - 9) %>%
  summarise(
    sphere    = "All Other Spheres",
    full_n    = sum(full_n, na.rm = TRUE),
    full_pct  = sum(full_pct, na.rm = TRUE),
    high_n    = sum(high_n, na.rm = TRUE),
    high_pct  = sum(high_pct, na.rm = TRUE),
    high_cmts = sum(high_cmts, na.rm = TRUE),
    high_mean = NA_real_,
    over_rep  = round(high_pct / full_pct, 2)
  )

table8_display <- bind_rows(top_spheres, other_row) %>%
  mutate(
    full_n   = format(full_n, big.mark = ","),
    full_pct = paste0(round(full_pct * 100, 1), "%"),
    high_pct = paste0(round(high_pct * 100, 1), "%"),
    high_cmts = ifelse(is.na(high_cmts) | high_cmts == 0, "---",
                       format(high_cmts, big.mark = ",")),
    high_mean = ifelse(is.na(high_mean) | high_mean == 0, "---",
                       as.character(high_mean))
  ) %>%
  select(sphere, full_n, full_pct, high_n, high_pct, high_cmts, high_mean, over_rep)

gt_table8 <- table8_display %>%
  gt() %>%
  cols_label(
    sphere    = "Regulatory Sphere",
    full_n    = "Full N",
    full_pct  = "Full %",
    high_n    = "50+ N",
    high_pct  = "50+ %",
    high_cmts = "50+ Cmts",
    high_mean = "50+ Mean",
    over_rep  = "Over-rep."
  ) %>%
  tab_header(
    title = "Table 8. Sectoral Composition: Full Dataset vs. High-Engagement Subset (50+ Comments)")

print(gt_table8)

# =============================================================================
# TABLE 9: Leading Developing Agencies in the High-Engagement Subset
# =============================================================================
# Use the 50+ Excel file for agency names (more detailed naming)
high_agency <- high %>%
  group_by(developer) %>%
  summarise(
    high_n    = n(),
    total_cmts = sum(comments, na.rm = TRUE),
    mean_cmts  = round(mean(comments, na.rm = TRUE)),
    .groups = "drop"
  ) %>%
  arrange(desc(high_n)) %>%
  slice_head(n = 10)

# Try to match with full dataset for full_n and 50+ rate
# Note: high file uses full formal names (e.g. "Ministry of Justice of the
# Republic of Uzbekistan") while full file uses short names ("Ministry of
# Justice"), so we use fuzzy substring matching instead of exact join.
full_agency_counts <- full %>%
  count(developer, name = "full_n")

# Build a lookup: for each high-file developer, find the best matching
# short name from the full file via substring containment.
# Manual overrides for names that differ by more than abbreviation
dev_name_map <- c(
  "Ministry of Higher Education, Science and Innovation of the Republic of Uzbekistan"
    = "Ministry of Higher Education, Science and Innovations",
  "Ministry of Poverty Reduction and Employment of the Republic of Uzbekistan"
    = "Ministry of Employment and Poverty Reduction",
  "State Inspection for Education Quality Control under the Cabinet of Ministers of the Republic of Uzbekistan"
    = "State Inspection for Educational Quality Control",
  "Ministry of Higher and Secondary Specialized Education of the Republic of Uzbekistan"
    = "Ministry of Public Education"
)

match_developer <- function(high_dev, full_devs) {
  # Check manual overrides first
  if (high_dev %in% names(dev_name_map)) return(dev_name_map[[high_dev]])
  # Fall back to substring containment
  hits <- full_devs[sapply(full_devs, function(fd) grepl(fd, high_dev, fixed = TRUE))]
  if (length(hits) == 0) return(NA_character_)
  hits[which.max(nchar(hits))]
}

high_agency <- high_agency %>%
  rowwise() %>%
  mutate(dev_match = match_developer(developer, full_agency_counts$developer)) %>%
  ungroup()

table9 <- high_agency %>%
  left_join(full_agency_counts, by = c("dev_match" = "developer")) %>%
  mutate(
    rate_50 = ifelse(is.na(full_n), "---",
                     paste0(round(high_n / full_n * 100, 1), "%"))
  ) %>%
  mutate(
    full_n = ifelse(is.na(full_n), "---", as.character(full_n)),
    total_cmts = format(total_cmts, big.mark = ",")
  ) %>%
  select(developer, high_n, total_cmts, mean_cmts, full_n, rate_50)

gt_table9 <- table9 %>%
  gt() %>%
  cols_label(
    developer  = "Developing Agency",
    high_n     = "50+ N",
    total_cmts = "Total Cmts",
    mean_cmts  = "Mean",
    full_n     = "Full N",
    rate_50    = "50+ Rate"
  ) %>%
  tab_header(
    title = "Table 9. Leading Developing Agencies in the High-Engagement Subset (50+ Comments)")

print(gt_table9)

# =============================================================================
# TABLE 10: Ten Most-Commented Draft Regulations
# =============================================================================
# Use the full dataset (sorted by comments descending) for top 10
table10 <- full %>%
  arrange(desc(comments)) %>%
  slice_head(n = 10) %>%
  mutate(
    # Abbreviate type names to match Word document
    type_abbr = case_when(
      type == "Ministry / Agency Resolution" ~ "Agency Res.",
      type == "Government Resolution"        ~ "Govt Res.",
      type == "Presidential Decree"          ~ "Pres. Decree",
      type == "Presidential Resolution"      ~ "Pres. Res.",
      TRUE                                   ~ type
    ),
    comments = format(comments, big.mark = ",")
  ) %>%
  select(year, title, developer, comments, type_abbr)

gt_table10 <- table10 %>%
  gt() %>%
  cols_label(
    year      = "Year",
    title     = "Regulation (abbreviated)",
    developer = "Developer",
    comments  = "Comments",
    type_abbr = "Type"
  ) %>%
  tab_header(
    title = "Table 10. Ten Most-Commented Draft Regulations, 2019-2024")

print(gt_table10)

# =============================================================================
# TABLE 11: Regulatory Type: Full Dataset vs. High-Engagement Subset
# =============================================================================
full_type <- full %>%
  mutate(
    type_grouped = if_else(
      type %in% c("Government Resolution", "Ministry / Agency Resolution",
                   "Presidential Resolution", "Law", "Presidential Decree"),
      type, "Other"
    )
  ) %>%
  count(type_grouped, name = "full_n") %>%
  mutate(full_pct = full_n / sum(full_n))

high_type <- high_from_full %>%
  mutate(
    type_grouped = if_else(
      type %in% c("Government Resolution", "Ministry / Agency Resolution",
                   "Presidential Resolution", "Law", "Presidential Decree"),
      type, "Other"
    )
  ) %>%
  count(type_grouped, name = "high_n") %>%
  mutate(high_pct = high_n / sum(high_n))

table11 <- full_type %>%
  left_join(high_type, by = "type_grouped") %>%
  replace_na(list(high_n = 0, high_pct = 0)) %>%
  mutate(
    shift    = round((high_pct - full_pct) * 100, 1),
    over_rep = round(high_pct / full_pct, 2)
  ) %>%
  mutate(type_grouped = factor(type_grouped, levels = type_order)) %>%
  arrange(type_grouped)

gt_table11 <- table11 %>%
  mutate(
    full_n   = format(full_n, big.mark = ","),
    full_pct = paste0(round(full_pct * 100, 1), "%"),
    high_pct = paste0(round(high_pct * 100, 1), "%"),
    shift    = ifelse(shift > 0, paste0("+", shift), as.character(shift))
  ) %>%
  select(type_grouped, full_n, full_pct, high_n, high_pct, shift, over_rep) %>%
  gt() %>%
  cols_label(
    type_grouped = "Type of Regulation",
    full_n       = "Full N",
    full_pct     = "Full %",
    high_n       = "50+ N",
    high_pct     = "50+ %",
    shift        = "Shift (p.p.)",
    over_rep     = "Over-rep."
  ) %>%
  tab_header(
    title = "Table 11. Regulatory Type: Full Dataset vs. High-Engagement Subset")

print(gt_table11)
