# This file takes the results of adding Munis and adds classes for 
# balanced asset classes (combinations of other classes)

# Read in the Excel file, probably named something like 'RAFI YYYYMM w Muni.xlsx'.
# Conservative	24% R3000 + 16% ACWIxUS + 60% Agg
# Moderate	30% R3000 + 20% ACWIxUS + 50% Agg
# Aggressive	36% R3000 + 24% ACWIxUS + 40% Agg

# The arithmetic returns will be weighted averages.
# The geometric return will be the arithmetic average - cov/2
# Portfolio Cov = wi^2 + 2*wi*wj*covij
# Portfolio Sigma = Portfolio Cov ^ 0.5

# Covij = Corij*sigmi*sigmaj

# Covip = covariance of asset i with portfolio p.   https://imgur.com/a/Tdis0ES
# Covip = wj * covij where wj is wt of jth asset in the portfolio
source("Read_RAFI_Utilities.r")
library(dplyr)
xls.file <- "RAFI 202403 w Muni.xlsx"
rafi.data.loc <- paste0(Sys.getenv("RT_PATH"), "Research/R/Development/Muni/")
erSheetNames <- c("Expected.Returns", "Expected Returns") # sheet names containing returns
riskSheetNames <- c("Expected.Risk.Matrix", "Expected Correlations") # sheet names containing corr and cov matrices
fnRAFI <- paste0(rafi.data.loc, xls.file)

rafi_xlranges <- list(date  = "C50",
                      return = "B4:K23",
                      corr = "C4:U23",
                      cov = "C35:U54")
rafi <- RedTor::readRAFIFile(fnRAFI, rafi_xlranges)

portfolioDefs <- RedTor::readExcelFile("BalancedPortfolioDefs.xlsx") %>%
  mutate(Asset.Class = make.names(Asset.Class))
portfolioNames <- unique(portfolioDefs$Name)

computePortStats <- function(portName, portfolioDefs, rafi) {
  portDef <- portfolioDefs %>% filter(Name == portName)
  if (sum(portDef$Wt) != 1) {
    warning("Sum of weights for ", portName, "portfolio do not equal 1. Portfolio skipped.")
  }
  tmp <- !(portDef$Asset.Class %in% rafi$ret$Asset.Class)
  if (sum(tmp) > 0) {
    warning("The asset classes ", portDef$Asset.Class[tmp], "not found in RAFI file. Portfolio skipped")
  }
  
  pRet <- rafi$ret %>% mutate(Arith = Expected.Return..Nominal. + Volatility^2 / 2, wt = 0) %>% replace(is.na(.), 0)
  idx <- match(portDef$Asset.Class, rafi$ret$Asset.Class)
  pRet[idx, "wt"] <- portDef$Wt
  portRisk <- as.numeric(sqrt(t(as.matrix(pRet$wt)) %*% as.matrix(rafi$cov[, -1]) %*% as.matrix(pRet$wt)))
  portCov <- apply(rafi$cov[, -1], 2, function(x) sum(x * pRet$wt))
  portCorr <- portCov / (diag(as.matrix(rafi$cov[, -1])))
  portRet <- apply(pRet[, c(-1, -ncol(pRet))], 2, function(x) sum(x * pRet$wt))
  portRet["Volatility"] <- portRisk
  portRet["Expected.Return..Real."] <- portRet["Arith"] - portRisk^2 / 2
  portRet["Expected.Return..Nominal."] <- portRet["Expected.Return..Real."] + 
    (pRet[1, 2] - pRet[1, 3])
  return(list(Asset.Class = portName, Wts = pRet$wt, Ret = portRet, Corr = portCorr, Cov = portCov))
}

portfolioStats <- lapply(unique(portfolioDefs$Name), computePortStats, portfolioDefs, rafi)
portCov <- matrix(0, ncol = length(portfolioNames), nrow = length(portfolioNames),
                  dimnames = list(portfolioNames, portfolioNames))
portCorr <- matrix(0, ncol = length(portfolioNames), nrow = length(portfolioNames),
                   dimnames = list(portfolioNames, portfolioNames))
for(i in 1:length(portfolioNames)) {
  for(j in 1:length(portfolioNames)) {
    portCov[i, j] <- portfolioStats[[i]]$Wts %*% as.matrix(rafi$cov[,-1]) %*% portfolioStats[[j]]$Wts
  }
}
for(i in 1:length(portfolioNames)) {
  for(j in 1:length(portfolioNames)) {
    portCorr[i, j] <- portCov[i, j] / (portCov[i, i]^0.5 * portCov[j,j]^0.5)
  }
}

newCov <- rafi$cov
newCorr <- rafi$corr
for (i in 1:length(portfolioStats)) {
  newCorr[nrow(newCorr) + 1, 1] <- portfolioStats[[i]]$Asset.Class
  newCorr[nrow(newCorr), 2:ncol(newCorr)] <- portfolioStats[[i]]$Corr
  newCov[nrow(newCov) + 1, 1] <- portfolioStats[[i]]$Asset.Class
  newCov[nrow(newCov), 2:ncol(newCov)] <- portfolioStats[[i]]$Cov
}

for (i in 1:length(portfolioStats)) {
  newCov <- cbind(newCov, c(portfolioStats[[i]]$Cov, portCov[, i]))
  newCorr <- cbind(newCorr, c(portfolioStats[[i]]$Corr, portCorr[, i]))
}
colnames(newCorr)[(ncol(newCorr) - 2):ncol(newCorr)] <- portfolioNames 
colnames(newCov)[(ncol(newCov) - 2):ncol(newCov)] <- portfolioNames 

