# VWIUX Functions

ggplotRegression <- function(fit, title="", xlab="Historical Volatility", 
                             ylab="RAFI Forecast") {
  
  require(ggplot2)
  
  ggplot(fit$model, aes_string(x = names(fit$model)[2], y = names(fit$model)[1])) + 
    geom_point() + 
    geom_abline(intercept = 0, slope = 1, color="gray") +
    geom_vline(xintercept = 0, color="gray") + 
    geom_hline(yintercept = 0, color="gray") +
    xlim(-0.5,1) + ylim(-0.5,1) +
    stat_smooth(method = "lm", col = "red") +
    labs(title=title, subtitle = paste("Adj R2 = ",round(summary(fit)$adj.r.squared, 3),
                                       "Intercept =",round(fit$coef[[1]],3 ),
                                       " Slope =",round(fit$coef[[2]], 3),
                                       " P =",round(summary(fit)$coef[2,4], 5)),
         x=xlab, y=ylab)
}

getxlranges <- function(){
  if (file.exists("xlranges.yaml")) {
    xl <- yaml::yaml.load_file("xlranges.yaml")
  } else {
    xl <- list()
    xl$rng.return <- "B4:L48"
    xl$rng.corr <- "C4:AC31"
    xl$rng.cov <- "C35:AC62"
    xl$rng.date <- "C50"
  }
  return(xl)
}

rafi.read.xls.v3 <- function(rafi.data.loc, xls.file.name = "Asset-Allocation-Interactive-Data.xlsx", xlInfo) {
  out <-list()
  out$as_of_date <-
    names(read_xlsx(paste0(rafi.data.loc, xls.file.name),
                    sheet = xlInfo$return.sheet.name,
                    range = xlInfo$rng.date))
  out$ret <-
    as.data.frame(read_xlsx(
      paste0(rafi.data.loc, xls.file.name),
      sheet = xlInfo$return.sheet.name,
      range = xlInfo$rng.return
    ))
  out$corr <-
    read_xlsx(paste0(rafi.data.loc, xls.file.name),
              sheet = xlInfo$risk.sheet.name,
              range = xlInfo$rng.corr)
  out$corr <- out$corr %>% mutate(Segment = names(out$corr)) %>% relocate(Segment)
  out$cov <-
    read_xlsx(paste0(rafi.data.loc, xls.file.name),
              sheet = xlInfo$risk.sheet.name,
              range = xlInfo$rng.cov)
  out$cov <- out$cov %>% mutate(Segment = names(out$cov)) %>% relocate(Segment)
  return(out)
}

rafi.read.xls.v2 <- function(rafi.data.loc, xls.file.name = "Asset-Allocation-Interactive-Data.xlsx") {
  xl <- yaml.load_file("xlranges.yaml")
  out <- list()
  out$as_of_date <-
    names(read_xlsx(paste0(rafi.data.loc, xls.file.name),
                    sheet = "Expected.Returns",
                    range = xl$rng.date))
  out$ret <-
    as.data.frame(read_xlsx(
      paste0(rafi.data.loc, xls.file.name),
      sheet = "Expected.Returns",
      range = xl$rng.return
    ))
  out$corr <-
    read_xlsx(paste0(rafi.data.loc, xls.file.name),
              sheet = "Expected.Risk.Matrix",
              range = xl$rng.corr)
  out$corr <- out$corr %>% mutate(Segment = names(out$corr)) %>% relocate(Segment)
  out$cov <-
    read_xlsx(paste0(rafi.data.loc, xls.file.name),
              sheet = "Expected.Risk.Matrix",
              range = xl$rng.cov)
  out$cov <- out$cov %>% mutate(Segment = names(out$cov)) %>% relocate(Segment)
  return(out)
}

# Function to get a series of returns on a list of tickers using adjusted series
# tickers is list of tickers
GetReturns <- function(tickers, freq = "D") {
  Sys.getenv("UTC")
  data <- do.call(cbind, lapply(tickers, getSymbols, auto.assign = FALSE))
  AdjustedCols <- grep(".Adjusted", colnames(data))
  data <- data[, AdjustedCols]
  data <- switch(
    toupper(freq),
    D = do.call(cbind, lapply(1:length(tickers), function(x)
      dailyReturn(data[, x]))),
    W = do.call(cbind, lapply(1:length(tickers), function(x)
      weeklyReturn(data[, x]))),
    M = do.call(cbind, lapply(1:length(tickers), function(x)
      monthlyReturn(data[, x])))
  )
  colnames(data) <- tickers
  data <- data[2:nrow(data), ]
  return(data)
}

# number of rows that are not na
nna <- function(x) {
  return(length(x) - sum(is.na(x)))
}