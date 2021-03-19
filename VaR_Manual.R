library(PortfolioAnalytics)
library(PerformanceAnalytics)
library(plyr)
library(stringr)
library(ggplot2)
library(plotly)
library(car)
library(quantmod)
library(quadprog)
library(dplyr)
library(openxlsx)
options(scipen = 999)
options(digits = 10)

# set the working directory
x <- dirname(rstudioapi::getSourceEditorContext()$path)
setwd(x)

# set the date range
start_date <- as.Date("2020-07-01")
end_date <- as.Date("2020-12-31")

# set data source
data_source <- "Web"

# Set the confidence level
conf_level <- 0.95

# get the porfolio data
portfolio <- read.xlsx("port.xlsx")
portfolio <- filter(portfolio , MKT_VAL > 0) %>% arrange(Instrument)
portfolio_d <- select(portfolio, 1, 2) %>% filter(MKT_VAL > 0)

# extract the instrument names
inst <- portfolio$Instrument

# import historical data
if (data_source == "Web") {
  # import the python code
  reticulate::source_python("dse.py")
  # import the imstruments data
  d <- dse_hist(start_date, end_date, inst)
} else {
  # read the data from file
  d <- read.xlsx("Book3.xlsx", detectDates = T)
  d <-
    filter(
      d,
      TRADING.CODE %in% inst &
        DATE < base::as.Date(end_date, format = "%Y-%m-%d") &
        DATE > base::as.Date(start_date, format = "%Y-%m-%d")
    )
}

# format data
d_split <- split(d, f = d$TRADING.CODE)
d_split <- lapply(d_split, function(x)
  as.data.frame(x[, c(1, 7)]))

portfolio_table <- join_all(d_split, by = "DATE", type = "left")
names(portfolio_table) <- c("Date", names(d_split))
names(portfolio_table) <-
  gsub("^\\s+|\\s+$", "", names(portfolio_table))
portfolio_table[is.na(portfolio_table)] <- 0

# calculate the return series
return_series <-
  xts(portfolio_table[, -1], order.by = as.Date(portfolio_table$Date))
return_series <-
  do.call(cbind, lapply(return_series, function(x)
    dailyReturn(x, type = "arithmetic")))[-1,]
return_series[is.na(return_series)] <- 0
return_series[!is.finite(return_series)] <- 0

fund_names <- colnames(portfolio_table)[-1]
names(return_series) <- fund_names

# get the weight of the instruments in portfolio
weight <- portfolio$MKT_VAL / sum(portfolio$MKT_VAL)
weight_data <- data.frame(Instrument = fund_names, Weight = weight)

# average of return series
avg_return <- data.frame(Avg_Return = apply(return_series, 2, mean))
avg_return$Instrument <- row.names(avg_return)
avg_return <- select(avg_return, 2, 1)
row.names(avg_return) <- 1:nrow(avg_return)

# standard deviation of return series
stdvp <- data.frame(Std_Dev = apply(return_series, 2, sd))
stdvp$Instrument <- row.names(stdvp)
stdvp <- stdvp[, c(2, 1)]
row.names(stdvp) <- 1:nrow(stdvp)

# portfolio details
portfolio_table <-
  join_all(list(portfolio_d, weight_data, avg_return, stdvp),
           by = "Instrument",
           type = "full")
portfolio_table$Gain <-
  (portfolio$MKT_VAL - portfolio$Cost) / portfolio$Cost

# portfolio summary
portfolio_summary <-
  data.frame(
    Total = nrow(portfolio_table),
    MKT_VAL = sum(portfolio_table$MKT_VAL),
    WGT_AVG_RET = (sum(portfolio_table$MKT_VAL) - sum(portfolio$Cost)) / sum(portfolio$Cost)
  )

# varcov and correlation matrix
varcov_matrix <- cov(return_series)
corr_matrix <- cor(return_series)

# calculate the VaR
# portfolio standard dev = [weight array * varcov matrix * T(weight array)]^(1/2) = [ W VARCOV W']^(1/2)
# portfolio mean = weight array * T(average return array) = [W AVG_RETURN']
# var percentage = - portfolio mean + portfolio standard dev * Z
# var = var percentage * portfolio size
port_std <-
  (matrix(weight, nrow = 1) %*% varcov_matrix %*% t(matrix(weight, nrow = 1))) ^ (1 /
                                                                                    2)
VaR_pct <- qnorm(conf_level) * port_std
VaR_pct_M <-
  -matrix(weight, nrow = 1) %*% matrix(avg_return$Avg_Return, ncol = 1) + VaR_pct
VaR_manual <- VaR_pct_M * sum(portfolio_table$MKT_VAL)

# Var for other time lengths
days <- c(5, 22, 22 * 2, 22 * 3)
VaR_pct_days <- as.vector(VaR_pct_M) * sqrt(days)
VaR_manual_days <- VaR_pct_days * sum(portfolio_table$MKT_VAL)

# generate output table
VaR_output <-
  data.frame(
    Days = c(1, days),
    VaR_Pct = c(as.vector(VaR_pct_M), VaR_pct_days),
    VaR = c(as.vector(VaR_manual), VaR_manual_days)
  )

write.xlsx(VaR_output, "VaR_Manual.xlsx", row.names = F)