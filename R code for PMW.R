rm(list = ls())
gc()


library(haven)
data <- read_dta("xxx.dta")
data_7824 <- read_dta("xxx.dta")

eid_to_keep <- setdiff(data$eid_ageing, data_7824$eid_ageing)
data <- data[data$eid_ageing %in% eid_to_keep, ]

mean(data$age_baseline)
sd(data$age_baseline)
table(data$gender)[1]/(table(data$gender)[1]+table(data$gender)[2]) #女性比例
sum(data$followup_death) #person-days

outcomes <- c("death", "cvddeath", "cancerdeath", "otherdeath", #死亡
              "mi", "stroke", "hf", "chd", "aaa", "pad", #心血管病
              "t2d", "hyperlipid", "hbp", #代谢病
              "renal_disease", "esrd", #肾
              "liver_disease", #肝
              "copd", "asthma", #肺
              "lungca", "breastca") #癌症
incident <- data.frame()
for (outcome in outcomes) {
  incident <- rbind(incident, cbind(outcome, sum(table(data[[paste("incident_",outcome,sep="")]])), table(data[[paste("incident_",outcome,sep="")]])[2], table(data[[paste("incident_",outcome,sep="")]])[2] / sum(data[[paste("followup_",outcome,sep="")]],na.rm = T) *1000*365, sum(data[[paste("followup_",outcome,sep="")]],na.rm = T)/365))
}
colnames(incident) <- c("outcome","participants","events","incident","followup_year")
columns_to_convert <- c("participants","events","incident","followup_year")
incident[columns_to_convert] <- lapply(incident[columns_to_convert], as.numeric)
library(openxlsx)
write.xlsx(incident, "xxx.xlsx")

sd(data$x5) #x5——ELM-ISOS
data$x5_sd <- data$x5/sd(data$x5)
sd(data$r32) #r32 - ELM-ISOS thickness of central subfield (right)
data$r32_sd <- data$r32/sd(data$r32)
sd(data$r33) #r33 - ELM-ISOS thickness of inner subfield (right)
data$r33_sd <- data$r33/sd(data$r33)
sd(data$r34) #r34 - ELM-ISOS thickness of outer subfield (right)
data$r34_sd <- data$r34/sd(data$r34)

library(readxl)
height_weight <- read_excel("xxx.xlsx", sheet = "Sheet1")
data <- merge(data, height_weight, by = "eid_ageing", all.x = T)

#ELM-ISOS
#x5(r7) - average ELM-ISOS (right)
#r32 - ELM-ISOS thickness of central subfield (right)
#r33 - ELM-ISOS thickness of inner subfield (right)
#r34 - ELM-ISOS thickness of outer subfield (right)

# 遍历的不同亚场的ELM-ISOS
ISOS_subfields <- c("x5", "r32", "r33", "r34")

outcomes <- c("death", "cvddeath", "cancerdeath", "otherdeath", #死亡
              "chd", "mi", "stroke", "hf", "pad", "aaa", #心血管病
              "t2d", "hyperlipid", "hbp", #代谢病
              "copd", "asthma", #肺
              "renal_disease", "esrd", #肾
              "liver_disease", #肝
              "lungca", "breastca") #癌症

results_list <- list()

library(survival)

for (ISOS_subfield in ISOS_subfields) {
  cox_result <- data.frame()
  
  for (outcome in outcomes) {
    data_filtered <- data[data[[paste0("followup_", outcome)]] >= 6 * 30, ]
    formula <- as.formula(paste0("Surv(followup_", outcome, ", incident_", outcome, ") ~ ", ISOS_subfield, "_sd + age_baseline + as.factor(a2) + height + weight + as.factor(a8) + o1"))
    cox_model <- coxph(formula, data = data_filtered)
    
    coef <- as.numeric(summary(cox_model)$coefficients[, 1])[1]
    HR <- as.numeric(summary(cox_model)$coefficients[, 2])[1]
    HR_low <- as.numeric(summary(cox_model)$conf.int[, 3])[1]
    HR_up <- as.numeric(summary(cox_model)$conf.int[, 4])[1]
    se <- as.numeric(summary(cox_model)$coefficients[, 3])[1]
    zvalue <- as.numeric(summary(cox_model)$coefficients[, 4])[1]
    pvalue <- as.numeric(summary(cox_model)$coefficients[, 5])[1]
    
    cox_result <- rbind(cox_result, c(outcome, coef, HR, HR_low, HR_up, se, zvalue, pvalue))
  }
  
  colnames(cox_result) <- c("outcome", "coef", "HR", "HR_low", "HR_up", "se", "zvalue", "pvalue")
  cox_result$coef <- as.numeric(cox_result$coef)
  cox_result$HR <- as.numeric(cox_result$HR)
  cox_result$HR_low <- as.numeric(cox_result$HR_low)
  cox_result$HR_up <- as.numeric(cox_result$HR_up)
  cox_result$se <- as.numeric(cox_result$se)
  cox_result$zvalue <- as.numeric(cox_result$zvalue)
  cox_result$pvalue <- as.numeric(cox_result$pvalue)
  
  cox_result$BH <- p.adjust(cox_result$pvalue, "BH")
  cox_result$BH_sig <- ifelse(cox_result$BH <= 0.05, "significant", "non-significant")
  
  results_list[[ISOS_subfield]] <- cox_result
}

results_list

library(openxlsx)
wb <- createWorkbook()
addWorksheet(wb, "average ELM-ISOS")
addWorksheet(wb, "average ELM-ISOS显著")
addWorksheet(wb, "central")
addWorksheet(wb, "central显著")
addWorksheet(wb, "inner")
addWorksheet(wb, "inner显著")
addWorksheet(wb, "outer")
addWorksheet(wb, "outer显著")

writeData(wb, "average ELM-ISOS", results_list$x5)
writeData(wb, "average ELM-ISOS显著", results_list$x5[results_list$x5$BH < 0.05, ])
writeData(wb, "central", results_list$r32)
writeData(wb, "central显著", results_list$r32[results_list$r32$BH < 0.05, ])
writeData(wb, "inner", results_list$r33)
writeData(wb, "inner显著", results_list$r33[results_list$r33$BH < 0.05, ])
writeData(wb, "outer", results_list$r34)
writeData(wb, "outer显著", results_list$r34[results_list$r34$BH < 0.05, ])
saveWorkbook(wb, "~/000_2 文章/036 文章11-ISOS代谢组/返修数据/UKB_ISOS——结局 排除6个月.xlsx", overwrite = TRUE)




rm(list = ls())
library(haven)
data_7824 <- read_dta("xxx.dta")

mean(data_7824$baselineage)
sd(data_7824$baselineage)
table(data_7824$a2)[1]/(table(data_7824$a2)[1]+table(data_7824$a2)[2]) #女性比例
sum(data_7824$followup_death) #person-days

sd(data_7824$x5) #x5——ELM-ISOS
data_7824$x5_sd <- data_7824$x5 * sd(data_7824$x5)
sd(data_7824$r32) #r32 - ELM-ISOS thickness of central subfield (right)
data_7824$r32_sd <- data_7824$r32 * sd(data_7824$r32)
sd(data_7824$r33) #r33 - ELM-ISOS thickness of inner subfield (right)
data_7824$r33_sd <- data_7824$r33 * sd(data_7824$r33)
sd(data_7824$r34) #r34 - ELM-ISOS thickness of outer subfield (right)
data_7824$r34_sd <- data_7824$r34 * sd(data_7824$r34)

library(readxl)
height_weight <- read_excel("xxx.xlsx", sheet = "Sheet1")
data_7824 <- merge(data_7824, height_weight, by = "eid_ageing", all.x = T)

lm_results_list <- list()
m_vars <- paste0("m", 1:168)
variables <- c("x5", "r32", "r33", "r34")

for (var in variables) {
  lm_result <- data.frame(variable = character(),
                          coef = numeric(),
                          tvalue = numeric(),
                          std = numeric(),
                          pvalue = numeric(),
                          stringsAsFactors = FALSE)
  
  for (var_m in m_vars) {
    formula <- paste0(var, "_sd ~", var_m, "+ baselineage + as.factor(a2) + height + weight + as.factor(a8) + o1")
    lm_model <- lm(formula, data = data_7824)
    
    coef <- as.numeric(summary(lm_model)$coefficients[, 1][2])
    std <- as.numeric(summary(lm_model)$coefficients[, 2][2])
    tvalue <- as.numeric(summary(lm_model)$coefficients[, 3][2])
    pvalue <- as.numeric(summary(lm_model)$coefficients[, 4][2])
    
    lm_result <- rbind(lm_result, c(var_m, coef, tvalue, std, pvalue))
  }
  
  colnames(lm_result) <- c("met", "coef", "tvalue", "std", "pvalue")
  lm_result$coef <- as.numeric(lm_result$coef)
  lm_result$std <- as.numeric(lm_result$std)
  lm_result$tvalue <- as.numeric(lm_result$tvalue)
  lm_result$pvalue <- as.numeric(lm_result$pvalue)
  lm_result$low_limit <- as.numeric(lm_result$coef) - 1.96 * as.numeric(lm_result$std)
  lm_result$up_limit <- as.numeric(lm_result$coef) + 1.96 * as.numeric(lm_result$std)
  lm_result <- lm_result[c("met", "coef", "low_limit", "up_limit", "std", "tvalue", "pvalue")]
  
  library(openxlsx)
  metname<- read.xlsx("xxx.xlsx", sheet = "Sheet1")
  lm_result <- cbind(metname, lm_result[,2:7])
  
  lm_result<- lm_result[1:168,]
  lm_result$BH <- p.adjust(lm_result$pvalue, "BH")
  
  lm_results_list[[var]] <- lm_result
}


met_filtered_x5 <- lm_results_list$x5[lm_results_list$x5$BH < 0.05, ]
met_filtered_r32 <- lm_results_list$r32[lm_results_list$r32$BH < 0.05, ]
met_filtered_r33 <- lm_results_list$r33[lm_results_list$r33$BH < 0.05, ]
met_filtered_r34 <- lm_results_list$r34[lm_results_list$r34$BH < 0.05, ]

library(openxlsx)
wb <- createWorkbook()
addWorksheet(wb, "average ELM-ISOS all 168")
addWorksheet(wb, "average ELM-ISOS selected 72")
addWorksheet(wb, "central all 168")
addWorksheet(wb, "central selected 20")
addWorksheet(wb, "inner all 168")
addWorksheet(wb, "inner selected 7")
addWorksheet(wb, "outer all 168")
addWorksheet(wb, "outer selected 25")

writeData(wb, "average ELM-ISOS all 168", lm_results_list$x5)
writeData(wb, "average ELM-ISOS selected 72", met_filtered_x5)
writeData(wb, "central all 168", lm_results_list$r32)
writeData(wb, "central selected 20", met_filtered_r32)
writeData(wb, "inner all 168", lm_results_list$r33)
writeData(wb, "inner selected 7", met_filtered_r33)
writeData(wb, "outer all 168", lm_results_list$r34)
writeData(wb, "outer selected 25", met_filtered_r34)

saveWorkbook(wb, "xxx.xlsx", overwrite = TRUE)


rm(list = ls())
library(haven)
data_86014 <- read_dta("xxx.dta")
mean(data_86014$baselineage)
sd(data_86014$baselineage)
table(data_86014$a2)[1]/(table(data_86014$a2)[1]+table(data_86014$a2)[2]) #女性比例

metabolite <- data_86014[, which(names(data_86014) == "eid_ageing"):which(names(data_86014) == "m249")]
covariate <- data_86014[, c("baselineage","a1", "a2", "a3", "a5", "a4", "a6", "a7", "a8", "a9", "a17", "a18", "a10", "a11", "a12", "a13", "a14", "a15", "a16")]
data_86014 <- cbind(metabolite, covariate)

data_86014$age_square <- (data_86014$baselineage)^2
library(readxl)
height_weight <- read_excel("xxx.xlsx", sheet = "Sheet1")
data_86014 <- merge(data_86014, height_weight, by = "eid_ageing", all.x = T)

data_502495 <- read_dta("xxx.dta")
names(data_502495)
data <- merge(data_86014, data_502495, by = "eid_ageing", all = FALSE)
names(data)

sum(data$followup_death)/365

outcomes <- c("death", "cancerdeath", "otherdeath", #死亡
              "t2d", "mi", "stroke", "hf", "chd", "aaa", "pad", #心血管代谢病
              "renal_disease", "esrd", #肾
              "liver_disease", #肝
              "copd", "asthma", #肺
              "lungca") #癌症
incident <- data.frame()
for (outcome in outcomes) {
  incident <- rbind(incident, cbind(outcome, sum(table(data[[paste("incident_",outcome,sep="")]])), table(data[[paste("incident_",outcome,sep="")]])[2], table(data[[paste("incident_",outcome,sep="")]])[2] / sum(data[[paste("followup_",outcome,sep="")]],na.rm = T) *1000*365, sum(data[[paste("followup_",outcome,sep="")]],na.rm = T)/365))
}
colnames(incident) <- c("outcome","participants","events","incident","followup_year")
columns_to_convert <- c("participants","events","incident","followup_year")
incident[columns_to_convert] <- lapply(incident[columns_to_convert], as.numeric)
library(openxlsx)
write.xlsx(incident, "xxx.xlsx")


# * average ####
library(openxlsx)
library(dplyr)
sigmet <- read.xlsx("xxx.xlsx", sheet = "average ELM-ISOS selected 72")
siglist <- sigmet[,1]

sig_out <- read.xlsx("xxx.xlsx", sheet = "average ELM-ISOS显著")
outcomes <- sig_out[,1]
outcomes <- append(outcomes, "chd", after = 3)
outcomes <- append(outcomes, "hf", after = 4)

results_list <- list()

library(survival)

for (outcome in outcomes) {
  cox_result <- data.frame()
  
  for (var in siglist) {
    formula <- as.formula(paste0("Surv(followup_", outcome, ",incident_", outcome, ") ~", var, " + age_baseline + as.factor(a2) + height + weight + as.factor(a8) + o1"))
    cox_model <- coxph(formula, data = data)
    
    coef <- as.numeric(summary(cox_model)$coefficients[,1][1])
    HR <- as.numeric(summary(cox_model)$coefficients[,2][1])
    HR_low <- as.numeric(summary(cox_model)$conf.int[,3][1])
    HR_up <- as.numeric(summary(cox_model)$conf.int[,4][1])
    se <- as.numeric(summary(cox_model)$coefficients[,3][1])
    zvalue <- as.numeric(summary(cox_model)$coefficients[,4][1])
    pvalue <- as.numeric(summary(cox_model)$coefficients[,5][1])
    
    cox_result <- rbind(cox_result, c(var, coef, HR, HR_low, HR_up, se, zvalue, pvalue))
  }
  
  colnames(cox_result) <- c("met", "coef", "HR", "HR_low", "HR_up", "se", "zvalue", "pvalue")
  cox_result$coef <- as.numeric(cox_result$coef)
  cox_result$HR <- as.numeric(cox_result$HR)
  cox_result$HR_low <- as.numeric(cox_result$HR_low)
  cox_result$HR_up <- as.numeric(cox_result$HR_up)
  cox_result$se <- as.numeric(cox_result$se)
  cox_result$zvalue <- as.numeric(cox_result$zvalue)
  cox_result$pvalue <- as.numeric(cox_result$pvalue)
  
  cox_result$BH <- p.adjust(cox_result$pvalue, "BH")
  cox_result$BH_sig <- ifelse(cox_result$BH <= 0.05, "significant", "non-significant")
  
  library(openxlsx)
  metname<- read.xlsx("xxx.xlsx", sheet = "Sheet1")
  cox_result <- left_join(cox_result, metname, by = "met")
  
  results_list[[outcome]] <- cox_result
}

results_list


library(openxlsx)
wb <- createWorkbook()

for (outcome in outcomes) {
  addWorksheet(wb, outcome)
  writeData(wb, outcome, results_list[[outcome]])
}

saveWorkbook(wb, "xxx.xlsx", overwrite = TRUE)




--------------------------------------------------------------------------------------------------------------------
# * central ####
library(openxlsx)
sigmet <- read.xlsx("xxx.xlsx", sheet = "central selected 20")
siglist <- sigmet[,1]

sig_out <- read.xlsx("xxx.xlsx", sheet = "central显著")
outcomes <- sig_out[,1]

results_list <- list()

library(survival)

for (outcome in outcomes) {
  cox_result <- data.frame(variable = character(),
                           coef = numeric(),
                           HR = numeric(),
                           HR_low = numeric(),
                           HR_up = numeric(),
                           se = numeric(),
                           zvalue = numeric(),
                           pvalue = numeric(),
                           stringsAsFactors = FALSE)
  
  for (var in siglist) {
    formula <- as.formula(paste0("Surv(followup_", outcome, ",incident_", outcome, ") ~", var, " + age_baseline + as.factor(a2) + height + weight + as.factor(a8)"))
    cox_model <- coxph(formula, data = data)
    
    coef <- as.numeric(summary(cox_model)$coefficients[,1][1])
    HR <- as.numeric(summary(cox_model)$coefficients[,2][1])
    HR_low <- as.numeric(summary(cox_model)$conf.int[,3][1])
    HR_up <- as.numeric(summary(cox_model)$conf.int[,4][1])
    se <- as.numeric(summary(cox_model)$coefficients[,3][1])
    zvalue <- as.numeric(summary(cox_model)$coefficients[,4][1])
    pvalue <- as.numeric(summary(cox_model)$coefficients[,5][1])
    
    cox_result <- rbind(cox_result, c(var, coef, HR, HR_low, HR_up, se, zvalue, pvalue))
  }
  
  colnames(cox_result) <- c("met", "coef", "HR", "HR_low", "HR_up", "se", "zvalue", "pvalue")
  cox_result$coef <- as.numeric(cox_result$coef)
  cox_result$HR <- as.numeric(cox_result$HR)
  cox_result$HR_low <- as.numeric(cox_result$HR_low)
  cox_result$HR_up <- as.numeric(cox_result$HR_up)
  cox_result$se <- as.numeric(cox_result$se)
  cox_result$zvalue <- as.numeric(cox_result$zvalue)
  cox_result$pvalue <- as.numeric(cox_result$pvalue)
  
  cox_result$BH <- p.adjust(cox_result$pvalue, "BH")
  cox_result$BH_sig <- ifelse(cox_result$BH <= 0.05, "significant", "non-significant")
  
  library(openxlsx)
  metname<- read.xlsx("xxx.xlsx", sheet = "Sheet1")
  cox_result <- left_join(cox_result, metname, by = "met")
  
  results_list[[outcome]] <- cox_result
}

results_list


library(openxlsx)
wb <- createWorkbook()

for (outcome in outcomes) {
  addWorksheet(wb, outcome)
  writeData(wb, outcome, results_list[[outcome]])
}

saveWorkbook(wb, "~/000_2 文章/036 文章11-ISOS代谢组/UKB_ISOS代谢物——结局_central.xlsx", overwrite = TRUE)





--------------------------------------------------------------------------------------------------------------------
# * inner ####
library(openxlsx)
sigmet <- read.xlsx("xxx.xlsx", sheet = "inner selected 7")
siglist <- sigmet[,1]

sig_out <- read.xlsx("xxx.xlsx", sheet = "inner显著")
outcomes <- sig_out[,1]

results_list <- list()

library(survival)

for (outcome in outcomes) {
  cox_result <- data.frame(variable = character(),
                           coef = numeric(),
                           HR = numeric(),
                           HR_low = numeric(),
                           HR_up = numeric(),
                           se = numeric(),
                           zvalue = numeric(),
                           pvalue = numeric(),
                           stringsAsFactors = FALSE)
  
  for (var in siglist) {
    formula <- as.formula(paste0("Surv(followup_", outcome, ",incident_", outcome, ") ~", var, " + age_baseline + as.factor(a2) + height + weight + as.factor(a8)"))
    cox_model <- coxph(formula, data = data)
    
    coef <- as.numeric(summary(cox_model)$coefficients[,1][1])
    HR <- as.numeric(summary(cox_model)$coefficients[,2][1])
    HR_low <- as.numeric(summary(cox_model)$conf.int[,3][1])
    HR_up <- as.numeric(summary(cox_model)$conf.int[,4][1])
    se <- as.numeric(summary(cox_model)$coefficients[,3][1])
    zvalue <- as.numeric(summary(cox_model)$coefficients[,4][1])
    pvalue <- as.numeric(summary(cox_model)$coefficients[,5][1])
    
    cox_result <- rbind(cox_result, c(var, coef, HR, HR_low, HR_up, se, zvalue, pvalue))
  }
  
  colnames(cox_result) <- c("met", "coef", "HR", "HR_low", "HR_up", "se", "zvalue", "pvalue")
  cox_result$coef <- as.numeric(cox_result$coef)
  cox_result$HR <- as.numeric(cox_result$HR)
  cox_result$HR_low <- as.numeric(cox_result$HR_low)
  cox_result$HR_up <- as.numeric(cox_result$HR_up)
  cox_result$se <- as.numeric(cox_result$se)
  cox_result$zvalue <- as.numeric(cox_result$zvalue)
  cox_result$pvalue <- as.numeric(cox_result$pvalue)
  
  cox_result$BH <- p.adjust(cox_result$pvalue, "BH")
  cox_result$BH_sig <- ifelse(cox_result$BH <= 0.05, "significant", "non-significant")
  
  library(openxlsx)
  metname<- read.xlsx("xxx.xlsx", sheet = "Sheet1")
  cox_result <- left_join(cox_result, metname, by = "met")
  
  results_list[[outcome]] <- cox_result
}

results_list


library(openxlsx)
wb <- createWorkbook()

for (outcome in outcomes) {
  addWorksheet(wb, outcome)
  writeData(wb, outcome, results_list[[outcome]])
}

saveWorkbook(wb, "~/000_2 文章/036 文章11-ISOS代谢组/UKB_ISOS代谢物——结局_inner.xlsx", overwrite = TRUE)






--------------------------------------------------------------------------------------------------------------------
# *3.4 outer ####
library(openxlsx)
sigmet <- read.xlsx("xxx.xlsx", sheet = "outer selected 25")
siglist <- sigmet[,1]

sig_out <- read.xlsx("xxx.xlsx", sheet = "outer显著")
outcomes <- sig_out[,1]

results_list <- list()

library(survival)

for (outcome in outcomes) {
  cox_result <- data.frame(variable = character(),
                           coef = numeric(),
                           HR = numeric(),
                           HR_low = numeric(),
                           HR_up = numeric(),
                           se = numeric(),
                           zvalue = numeric(),
                           pvalue = numeric(),
                           stringsAsFactors = FALSE)
  
  for (var in siglist) {
    formula <- as.formula(paste0("Surv(followup_", outcome, ",incident_", outcome, ") ~", var, " + age_baseline + as.factor(a2) + height + weight + as.factor(a8)"))
    cox_model <- coxph(formula, data = data)
    
    coef <- as.numeric(summary(cox_model)$coefficients[,1][1])
    HR <- as.numeric(summary(cox_model)$coefficients[,2][1])
    HR_low <- as.numeric(summary(cox_model)$conf.int[,3][1])
    HR_up <- as.numeric(summary(cox_model)$conf.int[,4][1])
    se <- as.numeric(summary(cox_model)$coefficients[,3][1])
    zvalue <- as.numeric(summary(cox_model)$coefficients[,4][1])
    pvalue <- as.numeric(summary(cox_model)$coefficients[,5][1])
    
    cox_result <- rbind(cox_result, c(var, coef, HR, HR_low, HR_up, se, zvalue, pvalue))
  }
  
  colnames(cox_result) <- c("met", "coef", "HR", "HR_low", "HR_up", "se", "zvalue", "pvalue")
  cox_result$coef <- as.numeric(cox_result$coef)
  cox_result$HR <- as.numeric(cox_result$HR)
  cox_result$HR_low <- as.numeric(cox_result$HR_low)
  cox_result$HR_up <- as.numeric(cox_result$HR_up)
  cox_result$se <- as.numeric(cox_result$se)
  cox_result$zvalue <- as.numeric(cox_result$zvalue)
  cox_result$pvalue <- as.numeric(cox_result$pvalue)
  
  cox_result$BH <- p.adjust(cox_result$pvalue, "BH")
  cox_result$BH_sig <- ifelse(cox_result$BH <= 0.05, "significant", "non-significant")
  
  library(openxlsx)
  metname<- read.xlsx("xxx.xlsx", sheet = "Sheet1")
  cox_result <- left_join(cox_result, metname, by = "met")
  
  results_list[[outcome]] <- cox_result
}

results_list


library(openxlsx)
wb <- createWorkbook()

for (outcome in outcomes) {
  addWorksheet(wb, outcome)
  writeData(wb, outcome, results_list[[outcome]])
}

saveWorkbook(wb, "xxx.xlsx", overwrite = TRUE)



library(openxlsx)
sigmet1 <- read.xlsx("xxx.xlsx", sheet = "average ELM-ISOS selected 72")
siglist1 <- sigmet1[,1]
sigmet2 <- read.xlsx("xxx.xlsx", sheet = "central selected 20")
siglist2 <- sigmet2[,1]
sigmet3 <- read.xlsx("xxx.xlsx", sheet = "inner selected 7")
siglist3 <- sigmet3[,1]
sigmet4 <- read.xlsx("xxx.xlsx", sheet = "outer selected 25")
siglist4 <- sigmet4[,1]
siglist <- union(union(union(siglist1, siglist2), siglist3), siglist4)

sig_out <- read.xlsx("xxx.xlsx", sheet = "average ELM-ISOS显著")
outcomes <- sig_out[,1]
outcomes <- append(outcomes, "chd", after = 3) #加上CHD
outcomes <- append(outcomes, "hf", after = 4) #加上hf

results_list <- list()

library(survival)

for (outcome in outcomes) {
  cox_result <- data.frame()
  
  for (var in siglist) {
    formula <- as.formula(paste0("Surv(followup_", outcome, ",incident_", outcome, ") ~", var, " + age_baseline + as.factor(a2) + height + weight + as.factor(a8)"))
    cox_model <- coxph(formula, data = data)
    
    coef <- as.numeric(summary(cox_model)$coefficients[,1][1])
    HR <- as.numeric(summary(cox_model)$coefficients[,2][1])
    HR_low <- as.numeric(summary(cox_model)$conf.int[,3][1])
    HR_up <- as.numeric(summary(cox_model)$conf.int[,4][1])
    se <- as.numeric(summary(cox_model)$coefficients[,3][1])
    zvalue <- as.numeric(summary(cox_model)$coefficients[,4][1])
    pvalue <- as.numeric(summary(cox_model)$coefficients[,5][1])
    
    cox_result <- rbind(cox_result, c(var, coef, HR, HR_low, HR_up, se, zvalue, pvalue))
  }
  
  colnames(cox_result) <- c("met", "coef", "HR", "HR_low", "HR_up", "se", "zvalue", "pvalue")
  cox_result$coef <- as.numeric(cox_result$coef)
  cox_result$HR <- as.numeric(cox_result$HR)
  cox_result$HR_low <- as.numeric(cox_result$HR_low)
  cox_result$HR_up <- as.numeric(cox_result$HR_up)
  cox_result$se <- as.numeric(cox_result$se)
  cox_result$zvalue <- as.numeric(cox_result$zvalue)
  cox_result$pvalue <- as.numeric(cox_result$pvalue)
  
  cox_result$BH <- p.adjust(cox_result$pvalue, "BH")
  cox_result$BH_sig <- ifelse(cox_result$BH <= 0.05, "significant", "non-significant")
  
  results_list[[outcome]] <- cox_result
}

results_list


library(openxlsx)
wb <- createWorkbook()

for (outcome in outcomes) {
  addWorksheet(wb, outcome)
  writeData(wb, outcome, results_list[[outcome]])
}

saveWorkbook(wb, "xxx.xlsx", overwrite = TRUE)



rm(list = ls())
library(haven)
data_86014 <- read_dta("xxx.dta")
metabolite <- data_86014[, which(names(data_86014) == "eid_ageing"):which(names(data_86014) == "m249")]
covariate <- data_86014[, c("baselineage","a1", "a2", "a3", "a5", "a4", "a6", "a7", "a8", "a9", "a17", "a18", "a10", "a11", "a12", "a13", "a14", "a15", "a16")]
data_86014 <- cbind(metabolite, covariate)

data_86014$age_square <- (data_86014$baselineage)^2
library(readxl)
height_weight <- read_excel("xxx.xlsx",
                            sheet = "Sheet1")
data_86014 <- merge(data_86014, height_weight, by = "eid_ageing", all.x = T)

data_502495 <- read_dta("xxx.dta")
names(data_502495)
data <- merge(data_86014, data_502495, by = "eid_ageing", all = FALSE)
names(data)

library(haven)
score_train <- read_dta("ypred_train新.dta")
score_train$dataset <- "train"
score_validate <- read_dta("ypred_val新.dta")
score_validate$dataset <- "validate"
score_test <- read_dta("ypred_test新.dta")
score_test$dataset <- "test"
score <- rbind(score_train, score_validate, score_test)
numeric_cols <- which(sapply(score, is.numeric) & !(names(score) %in% c("eid_ageing", "index"))) # 先找到除id列以外的所有数值列的索引
score[, numeric_cols] <- scale(score[, numeric_cols]) # 标准化除了id列以外的所有数值列


score$dataset_index <- score$index
score <- score[, -which(names(score) == "index")]
outcomes <- c('death','cancerdeath','otherdeath','chd','t2d','mi','stroke','pad','aaa','renal_disease','esrd','copd','asthma','liver_disease','lungca','hf') #给列命名
score_outcomes <- paste0("score_", outcomes)
colnames(score)[2:(ncol(score)-2)] <- score_outcomes

data <- merge(data, score, by = "eid_ageing", all.x = T)


# 仅使用需要的变量，以减少运算
data <- data[c("eid_ageing","eid_heart","dataset","a11","a22","a88","a55","a99","a1010","a1","a2","a3","a4","a5","a6","a7","a8","a9","a10","a12","a13","a14","a15","a16",
               "height","weight","waist","age_baseline","townsend","bmi","whr","pa_met","SBP","DBP","diet","ipaqc","sleep_score",
               "TC","HDL","triglycerides","fbg","screa_umol","hemoglobin","plt","egfr_scr","mau","uacr","ggt","ast","alt","fbg","sua","hba1c","AST_divide_ALT","TG_divide_HDL","LDL_divide_HDL","Albumin_divide_Scr",
               "prev_dm","prev_asthma","prev_hbp","prev_cvd_chd","prev_cvd","prev_cvd_af","prev_cvd_hf","prev_ckd","prev_dementia","prev_stroke","prev_copd","prev_vte","prev_cancer",
               "followup_death", "followup_cancerdeath", "followup_otherdeath", "followup_chd", "followup_mi", "followup_stroke", "followup_hf", "followup_pad", "followup_aaa", "followup_t2d", "followup_copd", "followup_asthma", "followup_renal_disease", "followup_esrd", "followup_liver_disease", "followup_lungca",
               "incident_death", "incident_cancerdeath", "incident_otherdeath", "incident_chd", "incident_mi", "incident_stroke", "incident_hf", "incident_pad", "incident_aaa", "incident_t2d", "incident_copd", "incident_asthma", "incident_renal_disease", "incident_esrd", "incident_liver_disease", "incident_lungca",
               "score_death", "score_cancerdeath", "score_otherdeath", "score_chd", "score_mi", "score_stroke", "score_hf", "score_pad", "score_aaa", "score_t2d", "score_copd", "score_asthma", "score_renal_disease", "score_esrd", "score_liver_disease", "score_lungca"
               )]


# 拆分数据集
table(data$dataset)
traindata <- subset(data, dataset == "train" | dataset == "validate")
testdata <- subset(data, dataset == "test")


# 正式分析
library(survival)
library(openxlsx)
outcomes <- c("death", "cancerdeath", "otherdeath", "chd", "mi", "stroke", "hf", "pad", "aaa", "t2d", "copd", "asthma", "renal_disease", "esrd", "liver_disease", "lungca")


# * 关联 ####
result_list_simplecox <- list()
for (outcome in outcomes) {
  formula <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ score_", outcome, " + as.factor(a1) + as.factor(a2) + height + weight + as.factor(a8)", sep = ""))
  cox_result <- coxph(formula, data = testdata)
  result_list_simplecox[[outcome]] <- cox_result
}

wb <- createWorkbook()
for (outcome in outcomes) {
  addWorksheet(wb, outcome)
  writeData(wb, outcome, summary(result_list_simplecox[[outcome]]))
}

saveWorkbook(wb, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_关联_简单校正协变量_新.xlsx", overwrite = TRUE)

library(readxl)
Summary <- data.frame()
for (outcome in names(result_list_simplecox)) {
  result <- read_excel("~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_关联_简单校正协变量_新.xlsx",sheet = outcome)
  result$Sig <- ifelse(result$`Pr(>|z|)` >= 0.05, "",
                       ifelse(result$`Pr(>|z|)` < 0.001, "***",
                              ifelse(result$`Pr(>|z|)` >= 0.001 & result$`Pr(>|z|)` < 0.01, "**", "*")))
  first_row <- as.data.frame(result[1,])
  Summary <- rbind(Summary, first_row)
}
names(Summary)[1] <- "outcome"
Summary$BH <- p.adjust(Summary$`Pr(>|z|)`, method = "BH")
addWorksheet(wb, "Summary")
writeData(wb, "Summary", Summary)
saveWorkbook(wb, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_关联_简单校正协变量_新.xlsx", overwrite = TRUE)

--------------------------------------------------------------------------------------------------------------------------------------------------------
result_list_completecox <- list()
for (outcome in outcomes) {
  formula <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ score_", outcome, " + as.factor(a1) + as.factor(a2) + height + weight + as.factor(a8) + as.factor(a5) + as.factor(a6) + as.factor(a7) + as.factor(a9) + as.factor(a10)", sep = ""))
  cox_result <- coxph(formula, data = testdata)
  result_list_completecox[[outcome]] <- cox_result
}

wb <- createWorkbook()
for (outcome in outcomes) {
  addWorksheet(wb, outcome)
  writeData(wb, outcome, summary(result_list_completecox[[outcome]]))
}

saveWorkbook(wb, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_关联_校正更多协变量.xlsx", overwrite = TRUE)

library(readxl)
Summary <- data.frame()
for (outcome in names(result_list_completecox)) {
  result <- read_excel("~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_关联_校正更多协变量.xlsx",sheet = outcome)
  result$Sig <- ifelse(result$`Pr(>|z|)` >= 0.05, "",
                       ifelse(result$`Pr(>|z|)` < 0.001, "***",
                              ifelse(result$`Pr(>|z|)` >= 0.001 & result$`Pr(>|z|)` < 0.01, "**", "*")))
  first_row <- as.data.frame(result[1,])
  Summary <- rbind(Summary, first_row)
}
names(Summary)[1] <- "outcome"
Summary$BH <- p.adjust(Summary$`Pr(>|z|)`, method = "BH")
addWorksheet(wb, "Summary")
writeData(wb, "Summary", Summary)
saveWorkbook(wb, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_关联_校正更多协变量.xlsx", overwrite = TRUE)


--------------------------------------------------------------------------------------------------------------------------------------------------------
# ** 亚组分析 ####
#拆分亚组数据集
data_young <- subset(data, a11 == 0)
data_old <- subset(data, a11 == 1)
data_male <- subset(data, a22 == 0)
data_female <- subset(data, a22 == 1)
data_white <- subset(data, a88 == 0)
data_other <- subset(data, a88 == 1)
data_ndep <- subset(data, a55 == 0)
data_dep <- subset(data, a55 == 1)
data_edu <- subset(data, a99 == 0)
data_nedu <- subset(data, a99 == 1)
data_fit <- subset(data, a1010 == 0)
data_obese <- subset(data, a1010 == 1)

covariate_list <- list("SimpleAdj" = "as.factor(a1) + as.factor(a2) + height + weight + as.factor(a8)",
                       "FullyAdj" = "as.factor(a1) + as.factor(a2) + height + weight + as.factor(a8) + as.factor(a5) + as.factor(a6) + as.factor(a7) + as.factor(a9) + as.factor(a10)"
)
subset_list <- list(
  "AGE" = list(data_subset1 = "data_young", data_subset2 = "data_old"),
  "SEX" = list(data_subset1 = "data_male", data_subset2 = "data_female"),
  "ETHNICITY" = list(data_subset1 = "data_white", data_subset2 = "data_other"),
  "TDI" = list(data_subset1 = "data_ndep", data_subset2 = "data_dep"),
  "EDUCATION" = list(data_subset1 = "data_edu", data_subset2 = "data_nedu"),
  "BMI" = list(data_subset1 = "data_fit", data_subset2 = "data_obese")
)

outcomes <- c("death", "cancerdeath", "otherdeath", "chd", "mi", "stroke", "hf", "pad", "aaa", "t2d", "copd", "asthma", "renal_disease", "esrd", "liver_disease", "lungca")

cox_result1 <- data.frame()
cox_result2 <- data.frame()

result_list <- list()

wb <- createWorkbook()


for (model in names(covariate_list)) {
  covariates <- covariate_list[[model]]
  for (subset_name in names(subset_list)) {
    data_subset1 <- get(subset_list[[subset_name]]$data_subset1)
    data_subset2 <- get(subset_list[[subset_name]]$data_subset2)
    data_subset1_name <- subset_list[[subset_name]]$data_subset1
    data_subset2_name <- subset_list[[subset_name]]$data_subset2
    
    if (subset_name == "SEX") {
      covariates <- gsub(" \\+ as.factor\\(a2\\)", "", covariates)
    }
    if (subset_name == "ETHNICITY") {
      covariates <- gsub(" \\+ as.factor\\(a8\\)", "", covariates)
    }
    if (subset_name == "TDI") {
      covariates <- gsub(" \\+ as.factor\\(a5\\)", "", covariates)
    }
    if (subset_name == "EDUCATION") {
      covariates <- gsub(" \\+ as.factor\\(a9\\)", "", covariates)
    }
    if (subset_name == "BMI") {
      covariates <- gsub(" \\+ as.factor\\(a10\\)", "", covariates)
    }
    
    result_list <- list()
    cox_result1 <- data.frame()
    cox_result2 <- data.frame()
    for (outcome in outcomes) {
      formula <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ score_", outcome, " + ", covariates, sep = ""))
      cox_model1 <- coxph(formula, data = data_subset1)
      cox_model2 <- coxph(formula, data = data_subset2)
      
      outcome_subset1 <- paste(outcome, gsub("data", "", data_subset1_name), sep = "")
      coef1 <- as.numeric(summary(cox_model1)$coefficients[,1][1])
      HR1 <- as.numeric(summary(cox_model1)$coefficients[,2][1])
      HR_low1 <- as.numeric(summary(cox_model1)$conf.int[,3][1])
      HR_up1 <- as.numeric(summary(cox_model1)$conf.int[,4][1])
      se1 <- as.numeric(summary(cox_model1)$coefficients[,3][1])
      zvalue1 <- as.numeric(summary(cox_model1)$coefficients[,4][1])
      pvalue1 <- as.numeric(summary(cox_model1)$coefficients[,5][1])
      sig1 <- ifelse(pvalue1 >= 0.05, "",
                     ifelse(pvalue1 < 0.001, "***",
                            ifelse(pvalue1 >= 0.001 & pvalue1 < 0.01, "**", "*")))
      
      outcome_subset2 <- paste(outcome, gsub("data", "", data_subset2_name), sep = "")
      coef2 <- as.numeric(summary(cox_model2)$coefficients[,1][1])
      HR2 <- as.numeric(summary(cox_model2)$coefficients[,2][1])
      HR_low2 <- as.numeric(summary(cox_model2)$conf.int[,3][1])
      HR_up2 <- as.numeric(summary(cox_model2)$conf.int[,4][1])
      se2 <- as.numeric(summary(cox_model2)$coefficients[,3][1])
      zvalue2 <- as.numeric(summary(cox_model2)$coefficients[,4][1])
      pvalue2 <- as.numeric(summary(cox_model2)$coefficients[,5][1])
      sig2 <- ifelse(pvalue2 >= 0.05, "",
                     ifelse(pvalue2 < 0.001, "***",
                            ifelse(pvalue2 >= 0.001 & pvalue2 < 0.01, "**", "*")))
      
      cox_result1 <- rbind(cox_result1, c(outcome_subset1, coef1, HR1, HR_low1, HR_up1, se1, zvalue1, pvalue1, sig1))
      cox_result2 <- rbind(cox_result2, c(outcome_subset2, coef2, HR2, HR_low2, HR_up2, se2, zvalue2, pvalue2, sig2))
    }
    cox_result <- cbind(cox_result1, cox_result2)
    colnames(cox_result) <- c("outcome_superior", "coef_superior", "HR_superior", "HR_low_superior", "HR_up_superior", "se_superior", "zvalue_superior", "pvalue_superior", "sig_superior",
                              "outcome_inferior", "coef_inferior", "HR_inferior", "HR_low_inferior", "HR_up_inferior", "se_inferior", "zvalue_inferior", "pvalue_inferior", "sig_inferior")
    columns_to_convert <- c("coef_superior", "HR_superior", "HR_low_superior", "HR_up_superior", "se_superior", "zvalue_superior", "pvalue_superior",
                            "coef_inferior", "HR_inferior", "HR_low_inferior", "HR_up_inferior", "se_inferior", "zvalue_inferior", "pvalue_inferior")
    cox_result[columns_to_convert] <- lapply(cox_result[columns_to_convert], as.numeric)
    
    sheetname <- paste(subset_name, "_", model, sep = "")
    addWorksheet(wb, sheetname)
    writeData(wb, sheetname, cox_result)
    
  }
  
}

saveWorkbook(wb, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_关联_亚组分析.xlsx", overwrite = TRUE)




--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# * K-M曲线 RR值 ####

outcomes <- c("death", "cancerdeath", "otherdeath", "chd", "mi", "stroke", "hf", "pad", "aaa", "t2d", "copd", "asthma", "renal_disease", "esrd", "liver_disease", "lungca")
results <- data.frame()

for (outcome in outcomes) {
  riskScore <- paste("score_", outcome, sep = "")
  risk_decile <- ifelse(!is.na(testdata[[riskScore]]) & (testdata[[riskScore]] > quantile(testdata[[riskScore]], 0.9, na.rm = T)), "HIGH",
                        ifelse(!is.na(testdata[[riskScore]]) & (testdata[[riskScore]] < quantile(testdata[[riskScore]], 0.1, na.rm = T)), "LOW",
                               ifelse(!is.na(testdata[[riskScore]]) & (testdata[[riskScore]] > quantile(testdata[[riskScore]], 0.5, na.rm = T) & testdata[[riskScore]] < quantile(testdata[[riskScore]], 0.6, na.rm = T)), "MIDDLE",
                                      NA)))
  
  event <- paste("incident_", outcome, sep = "")
  time <- paste("followup_", outcome, sep = "")
  event <- testdata[[event]]
  time <- testdata[[time]]
  
  calculate <- as.data.frame(cbind(event, time, risk_decile))
  calculate$event <- as.factor(calculate$event)
  calculate$time <- as.numeric(calculate$time)
  calculate$risk_decile <- as.factor(calculate$risk_decile)
  
  calculate <- na.omit(calculate)
  
  top <- subset(calculate,calculate$risk_decile=="HIGH") # top10%数据
  bottom <- subset(calculate,calculate$risk_decile=="LOW") # bottom10%数据
  a <- summary(top$event)[2] # top10%发病数
  b <- sum(top$time)/365 # top10%人年
  c <- summary(bottom$event)[2] #bottom10%发病数
  d <- sum(bottom$time)/365 #bottom10%人年
  e1 <- a/b*1000 # c为千人年发病率
  e2 <- c/d*1000 # c为千人年发病率
  RR <- e1/e2 ###相对发病率
  lower = exp(log(RR) - 1.96*sqrt((1/a)+(1/c))) # 上限
  upper = exp(log(RR) + 1.96*sqrt((1/a)+(1/c))) # 下限
  diff <- survdiff(Surv(time,as.numeric(event)) ~ risk_decile, data = calculate)
  pValue = 1-pchisq(diff$chisq, df=1)
  pValue <- signif(pValue,3)
  pValue <- format(pValue, scientific = TRUE)
  
  result <- cbind(outcome, RR, lower, upper, pValue)
  results <- rbind(results, result)
}

results$RR <- as.numeric(results$RR)
results$lower <- as.numeric(results$lower)
results$upper <- as.numeric(results$upper)
results$pValue <- as.numeric(results$pValue)
results
write.xlsx(results, "~/000_2 文章/036 文章11-ISOS代谢组/K-M曲线.xlsx")

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------



--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# * 预测 ####

# ** 单独predictor比较 ####
library(pROC)
library(survival)

auc_results <- data.frame()

for (outcome in outcomes) {
  # 单独代谢物
  formula_met <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ score_", outcome, sep = ""))
  model_met <- coxph(formula_met, data = traindata)
  testdata[[paste(outcome, "_riskscore_met", sep = "")]] <- predict(model_met, type = "risk", newdata = testdata)
  roc_result_met <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste("score_", outcome, sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_met <- roc_result_met$auc
  ci_lower_met <- roc_result_met$ci[1]
  ci_upper_met <- roc_result_met$ci[3]
  
  # 年龄
  formula_age <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ age_baseline", sep = ""))
  model_age <- coxph(formula_age, data = traindata)
  testdata[[paste(outcome, "_riskscore_age", sep = "")]] <- predict(model_age, type = "risk", newdata = testdata)
  roc_result_age <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste(outcome, "_riskscore_age", sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_age <- roc_result_age$auc
  ci_lower_age <- roc_result_age$ci[1]
  ci_upper_age <- roc_result_age$ci[3]
  test_age <- roc.test(roc_result_met,roc_result_age) #Delong's test
  p_age <- test_age[["p.value"]]
  compare_age <- ifelse(p_age >= 0.05, "Comparable", ifelse(auc_met > auc_age, "Superior", "-"))
  
  # 性别
  formula_sex <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ as.factor(a2)", sep = ""))
  model_sex <- coxph(formula_sex, data = traindata)
  testdata[[paste(outcome, "_riskscore_sex", sep = "")]] <- predict(model_sex, type = "risk", newdata = testdata)
  roc_result_sex <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste(outcome, "_riskscore_sex", sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_sex <- roc_result_sex$auc
  ci_lower_sex <- roc_result_sex$ci[1]
  ci_upper_sex <- roc_result_sex$ci[3]
  test_sex <- roc.test(roc_result_met,roc_result_sex) #Delong's test
  p_sex <- test_sex[["p.value"]]
  compare_sex <- ifelse(p_sex >= 0.05, "Comparable", ifelse(auc_met > auc_sex, "Superior", "-"))
  
  # 收入
  formula_income <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ as.factor(a4)", sep = ""))
  model_income <- coxph(formula_income, data = traindata)
  testdata[[paste(outcome, "_riskscore_income", sep = "")]] <- predict(model_income, type = "risk", newdata = testdata)
  roc_result_income <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste(outcome, "_riskscore_income", sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_income <- roc_result_income$auc
  ci_lower_income <- roc_result_income$ci[1]
  ci_upper_income <- roc_result_income$ci[3]
  test_income <- roc.test(roc_result_met,roc_result_income) #Delong's test
  p_income <- test_income[["p.value"]]
  compare_income <- ifelse(p_income >= 0.05, "Comparable", ifelse(auc_met > auc_income, "Superior", "-"))
  
  # TDI
  formula_tdi <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ as.factor(a5)", sep = ""))
  model_tdi <- coxph(formula_tdi, data = traindata)
  testdata[[paste(outcome, "_riskscore_tdi", sep = "")]] <- predict(model_tdi, type = "risk", newdata = testdata)
  roc_result_tdi <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste(outcome, "_riskscore_tdi", sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_tdi <- roc_result_tdi$auc
  ci_lower_tdi <- roc_result_tdi$ci[1]
  ci_upper_tdi <- roc_result_tdi$ci[3]
  test_tdi <- roc.test(roc_result_met,roc_result_tdi) #Delong's test
  p_tdi <- test_tdi[["p.value"]]
  compare_tdi <- ifelse(p_tdi >= 0.05, "Comparable", ifelse(auc_met > auc_tdi, "Superior", "-"))
  
  # 吸烟
  formula_smoking <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ as.factor(a6)", sep = ""))
  model_smoking <- coxph(formula_smoking, data = traindata)
  testdata[[paste(outcome, "_riskscore_smoking", sep = "")]] <- predict(model_smoking, type = "risk", newdata = testdata)
  roc_result_smoking <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste(outcome, "_riskscore_smoking", sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_smoking <- roc_result_smoking$auc
  ci_lower_smoking <- roc_result_smoking$ci[1]
  ci_upper_smoking <- roc_result_smoking$ci[3]
  test_smoking <- roc.test(roc_result_met,roc_result_smoking) #Delong's test
  p_smoking <- test_smoking[["p.value"]]
  compare_smoking <- ifelse(p_smoking >= 0.05, "Comparable", ifelse(auc_met > auc_smoking, "Superior", "-"))
  
  # 喝酒
  formula_drinking <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ as.factor(a7)", sep = ""))
  model_drinking <- coxph(formula_drinking, data = traindata)
  testdata[[paste(outcome, "_riskscore_drinking", sep = "")]] <- predict(model_drinking, type = "risk", newdata = testdata)
  roc_result_drinking <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste(outcome, "_riskscore_drinking", sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_drinking <- roc_result_drinking$auc
  ci_lower_drinking <- roc_result_drinking$ci[1]
  ci_upper_drinking <- roc_result_drinking$ci[3]
  test_drinking <- roc.test(roc_result_met,roc_result_drinking) #Delong's test
  p_drinking <- test_drinking[["p.value"]]
  compare_drinking <- ifelse(p_drinking >= 0.05, "Comparable", ifelse(auc_met > auc_drinking, "Superior", "-"))
  
  # 种族
  formula_ethnicity <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ as.factor(a8)", sep = ""))
  model_ethnicity <- coxph(formula_ethnicity, data = traindata)
  testdata[[paste(outcome, "_riskscore_ethnicity", sep = "")]] <- predict(model_ethnicity, type = "risk", newdata = testdata)
  roc_result_ethnicity <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste(outcome, "_riskscore_ethnicity", sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_ethnicity <- roc_result_ethnicity$auc
  ci_lower_ethnicity <- roc_result_ethnicity$ci[1]
  ci_upper_ethnicity <- roc_result_ethnicity$ci[3]
  test_ethnicity <- roc.test(roc_result_met,roc_result_ethnicity) #Delong's test
  p_ethnicity <- test_ethnicity[["p.value"]]
  compare_ethnicity <- ifelse(p_ethnicity >= 0.05, "Comparable", ifelse(auc_met > auc_ethnicity, "Superior", "-"))
  
  # 教育
  formula_education <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ as.factor(a9)", sep = ""))
  model_education <- coxph(formula_education, data = traindata)
  testdata[[paste(outcome, "_riskscore_education", sep = "")]] <- predict(model_education, type = "risk", newdata = testdata)
  roc_result_education <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste(outcome, "_riskscore_education", sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_education <- roc_result_education$auc
  ci_lower_education <- roc_result_education$ci[1]
  ci_upper_education <- roc_result_education$ci[3]
  test_education <- roc.test(roc_result_met,roc_result_education) #Delong's test
  p_education <- test_education[["p.value"]]
  compare_education <- ifelse(p_education >= 0.05, "Comparable", ifelse(auc_met > auc_education, "Superior", "-"))
  
  # BMI
  formula_bmi <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ as.factor(a10)", sep = ""))
  model_bmi <- coxph(formula_bmi, data = traindata)
  testdata[[paste(outcome, "_riskscore_bmi", sep = "")]] <- predict(model_bmi, type = "risk", newdata = testdata)
  roc_result_bmi <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste(outcome, "_riskscore_bmi", sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_bmi <- roc_result_bmi$auc
  ci_lower_bmi <- roc_result_bmi$ci[1]
  ci_upper_bmi <- roc_result_bmi$ci[3]
  test_bmi <- roc.test(roc_result_met,roc_result_bmi) #Delong's test
  p_bmi <- test_bmi[["p.value"]]
  compare_bmi <- ifelse(p_bmi >= 0.05, "Comparable", ifelse(auc_met > auc_bmi, "Superior", "-"))
  
  # 降脂药
  formula_lipid_lowering <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ as.factor(a12)", sep = ""))
  model_lipid_lowering <- coxph(formula_lipid_lowering, data = traindata)
  testdata[[paste(outcome, "_riskscore_lipid_lowering", sep = "")]] <- predict(model_lipid_lowering, type = "risk", newdata = testdata)
  roc_result_lipid_lowering <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste(outcome, "_riskscore_lipid_lowering", sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_lipid_lowering <- roc_result_lipid_lowering$auc
  ci_lower_lipid_lowering <- roc_result_lipid_lowering$ci[1]
  ci_upper_lipid_lowering <- roc_result_lipid_lowering$ci[3]
  test_lipid_lowering <- roc.test(roc_result_met,roc_result_lipid_lowering) #Delong's test
  p_lipid_lowering <- test_lipid_lowering[["p.value"]]
  compare_lipid_lowering <- ifelse(p_lipid_lowering >= 0.05, "Comparable", ifelse(auc_met > auc_lipid_lowering, "Superior", "-"))
  
  # 降压药
  formula_bp_lowering <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ as.factor(a13)", sep = ""))
  model_bp_lowering <- coxph(formula_bp_lowering, data = traindata)
  testdata[[paste(outcome, "_riskscore_bp_lowering", sep = "")]] <- predict(model_bp_lowering, type = "risk", newdata = testdata)
  roc_result_bp_lowering <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste(outcome, "_riskscore_bp_lowering", sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_bp_lowering <- roc_result_bp_lowering$auc
  ci_lower_bp_lowering <- roc_result_bp_lowering$ci[1]
  ci_upper_bp_lowering <- roc_result_bp_lowering$ci[3]
  test_bp_lowering <- roc.test(roc_result_met,roc_result_bp_lowering) #Delong's test
  p_bp_lowering <- test_bp_lowering[["p.value"]]
  compare_bp_lowering <- ifelse(p_bp_lowering >= 0.05, "Comparable", ifelse(auc_met > auc_bp_lowering, "Superior", "-"))
  
  
  
  # 将结果添加到 auc_results 数据框中，并改名
  auc_results <- rbind(auc_results, c(outcome, auc_met, ci_lower_met, ci_upper_met,
                                      auc_age, ci_lower_age, ci_upper_age, p_age, compare_age,
                                      auc_sex, ci_lower_sex, ci_upper_sex, p_sex, compare_sex,
                                      auc_income, ci_lower_income, ci_upper_income, p_income, compare_income,
                                      auc_tdi, ci_lower_tdi, ci_upper_tdi, p_tdi, compare_tdi,
                                      auc_smoking, ci_lower_smoking, ci_upper_smoking, p_smoking, compare_smoking,
                                      auc_drinking, ci_lower_drinking, ci_upper_drinking, p_drinking, compare_drinking,
                                      auc_ethnicity, ci_lower_ethnicity, ci_upper_ethnicity, p_ethnicity, compare_ethnicity,
                                      auc_education, ci_lower_education, ci_upper_education, p_education, compare_education,
                                      auc_bmi, ci_lower_bmi, ci_upper_bmi, p_bmi, compare_bmi,
                                      auc_lipid_lowering, ci_lower_lipid_lowering, ci_upper_lipid_lowering, p_lipid_lowering, compare_lipid_lowering,
                                      auc_bp_lowering, ci_lower_bp_lowering, ci_upper_bp_lowering, p_bp_lowering, compare_bp_lowering
                                      ))
  colnames(auc_results) <- c("outcome", "MET", "met_lower", "met_upper",
                             "AGE", "age_lower", "age_upper", "Pvalue_AGE", "MET_vs._AGE",
                             "SEX", "sex_lower", "sex_upper", "Pvalue_SEX", "MET_vs._SEX",
                             "INCOME", "income_lower", "income_upper", "Pvalue_INCOME", "MET_vs._INCOME",
                             "TDI", "tdi_lower", "tdi_upper", "Pvalue_TDI", "MET_vs._TDI",
                             "SMOKING", "smoking_lower", "smoking_upper", "Pvalue_SMOKING", "MET_vs._SMOKING",
                             "DRINKING", "drinking_lower", "drinking_upper", "Pvalue_DRINKING", "MET_vs._DRINKING",
                             "ETHNICITY", "ethinicty_lower", "ethinicty_upper", "Pvalue_ETHNICITY", "MET_vs._ETHNICITY",
                             "EDUCATION", "education_lower", "education_upper", "Pvalue_EDUCATION", "MET_vs._EDUCATION",
                             "BMI", "bmi_lower", "bmi_upper", "Pvalue_BMI", "MET_vs._BMI",
                             "ANTI_LIPID", "anti_lipid_lower", "anti_lipid_upper", "Pvalue_ANTI_LIPID", "MET_vs._ANTI_LIPID",
                             "ANTI_BP", "anti_bp_lower", "anti_bp_upper", "Pvalue_ANTI_BP", "MET_vs._ANTI_BP")
}

#转换数值变量
columns_to_convert <- c("MET", "met_lower", "met_upper",
                        "AGE", "age_lower", "age_upper", "Pvalue_AGE",
                        "SEX", "sex_lower", "sex_upper", "Pvalue_SEX",
                        "INCOME", "income_lower", "income_upper", "Pvalue_INCOME",
                        "TDI", "tdi_lower", "tdi_upper", "Pvalue_TDI",
                        "SMOKING", "smoking_lower", "smoking_upper", "Pvalue_SMOKING",
                        "DRINKING", "drinking_lower", "drinking_upper", "Pvalue_DRINKING",
                        "ETHNICITY", "ethinicty_lower", "ethinicty_upper", "Pvalue_ETHNICITY",
                        "EDUCATION", "education_lower", "education_upper", "Pvalue_EDUCATION",
                        "BMI", "bmi_lower", "bmi_upper", "Pvalue_BMI",
                        "ANTI_LIPID", "anti_lipid_lower", "anti_lipid_upper", "Pvalue_ANTI_LIPID",
                        "ANTI_BP", "anti_bp_lower", "anti_bp_upper", "Pvalue_ANTI_BP")
auc_results[columns_to_convert] <- lapply(auc_results[columns_to_convert], as.numeric)

write.xlsx(auc_results, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_预测_Transformer新结果/UKB关联交互预测_预测_单独predictor比较.xlsx")


--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

models <- create_models_and_outcomes()$models
outcomes <- create_models_and_outcomes()$outcomes
Category <- create_models_and_outcomes()$category


# ** 总人群预测 ####
library(pROC)
library(survival)
library(openxlsx)
library(dplyr)

auc_results_list <- list()

wb <- createWorkbook()


for (model_name in names(models)) {
  auc_results <- data.frame()
  
  for (outcome in outcomes) {
    
    #如果是hf结局，模型中删除prev_hf
    if (outcome == "hf") {
      for (model in names(models)) {
        # 使用gsub删除模型字符串中的+ as.factor(prev_cvd_hf)
        models[[model]] <- gsub("\\+ as\\.factor\\(prev_cvd_hf\\)", "", models[[model]])
      }
    }
    # #如果是t2d结局，模型中删除prev_dm
    # if (outcome == "t2d") {
    #   for (model in names(models)) {
    #     # 使用gsub删除模型字符串中的+ as.factor(prev_dm)
    #     models[[model]] <- gsub("\\+ as\\.factor\\(prev_dm\\)", "", models[[model]])
    #   }
    # }
    #如果是asthma结局，模型中删除prev_asthma
    if (outcome == "asthma") {
      for (model in names(models)) {
        # 使用gsub删除模型字符串中的+ as.factor(prev_asthma)
        models[[model]] <- gsub("\\+ as\\.factor\\(prev_asthma\\)", "", models[[model]])
      }
    }
    #如果是stroke结局，模型中删除prev_stroke
    if (outcome == "stroke") {
      for (model in names(models)) {
        # 使用gsub删除模型字符串中的+ as.factor(prev_stroke)
        models[[model]] <- gsub("\\+ as\\.factor\\(prev_stroke\\)", "", models[[model]])
      }
    }
    #如果是renal_disease结局，模型中删除prev_ckd
    if (outcome == "renal_disease") {
      for (model in names(models)) {
        # 使用gsub删除模型字符串中的+ as.factor(prev_ckd)
        models[[model]] <- gsub("\\+ as\\.factor\\(prev_ckd\\)", "", models[[model]])
      }
    }
    #如果是copd结局，模型中删除prev_copd
    if (outcome == "copd") {
      for (model in names(models)) {
        # 使用gsub删除模型字符串中的+ as.factor(prev_copd)
        models[[model]] <- gsub("\\+ as\\.factor\\(prev_copd\\)", "", models[[model]])
      }
    }
    
    covariate <- models[[model_name]]
    
    # 传统模型
    formula1 <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ ", covariate, sep = ""))
    model1 <- coxph(formula1, data = traindata)
    testdata[[paste(outcome, "_", model_name, "_1", sep = "")]] <- predict(model1, type = "risk", newdata = testdata)
    roc_result1 <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste(outcome, "_", model_name, "_1", sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
    auc_value1 <- roc_result1$auc
    ci_lower1 <- roc_result1$ci[1]
    ci_upper1 <- roc_result1$ci[3]
    
    # 加入代谢物
    formula2 <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ ", covariate, "+ score_", outcome, sep = ""))
    model2 <- coxph(formula2, data = traindata)
    testdata[[paste(outcome, "_", model_name, "_2", sep = "")]] <- predict(model2, type = "risk", newdata = testdata)
    roc_result2 <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste(outcome, "_", model_name, "_2", sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
    auc_value2 <- roc_result2$auc
    ci_lower2 <- roc_result2$ci[1]
    ci_upper2 <- roc_result2$ci[3]
    test <- roc.test(roc_result1,roc_result2) #Delong's test
    p <- test[["p.value"]]
    auc_diff <- auc_value2 - auc_value1
    auc_improve_proportion <- auc_diff / auc_value1
    sig <- ifelse(p >= 0.05, "",
                  ifelse(p < 0.001, "***",
                         ifelse(p >=0.001 & p < 0.01, "**", "*")))
    
    # 单独代谢物
    roc_result3 <- roc(testdata[[paste("incident_", outcome, sep = "")]] ~ testdata[[paste("score_", outcome, sep = "")]], data = testdata, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
    auc_value3 <- roc_result3$auc
    ci_lower3 <- roc_result3$ci[1]
    ci_upper3 <- roc_result3$ci[3]
    test_met <- roc.test(roc_result1,roc_result3) #Delong's test
    p_met <- test_met[["p.value"]]
    compare_met <- ifelse(p_met >= 0.05, "Comparable", ifelse(auc_value3 > auc_value1, "Superior", "-"))
    
    # 将结果添加到 auc_results 数据框中，并改名
    auc_results <- rbind(auc_results, c(outcome, auc_value1, ci_lower1, ci_upper1, auc_value3, ci_lower3, ci_upper3, p_met, compare_met, auc_value2, ci_lower2, ci_upper2, auc_diff, auc_improve_proportion, p, sig))
    colnames(auc_results) <- c("outcome", "CONV", "conv_lower", "conv_upper", "MET", "met_lower", "met_upper", "Pvalue_MET_vs_CONV", "MET_vs._CONV", "COMBINED", "combined_lower", "combined_upper", "Diff", "Diff_proportion" ,"Pvalue_COMBINED_vs_CONV", "Significance")
    
    models <- create_models_and_outcomes()$models
    outcomes <- create_models_and_outcomes()$outcomes
  }
  
  # 增加一列结局分类
  Category <- create_models_and_outcomes()$category
  auc_results <- cbind(auc_results, Category)
  auc_results <- auc_results %>% select(which(names(auc_results) == "outcome"):which(names(auc_results) == "Pvalue_COMBINED_vs_CONV"), "Category", "Significance") #显著标记放在最后一列
  
  # as.numeric
  auc_results$CONV <- as.numeric(auc_results$CONV)
  auc_results$conv_lower <- as.numeric(auc_results$conv_lower)
  auc_results$conv_upper <- as.numeric(auc_results$conv_upper)
  auc_results$MET <- as.numeric(auc_results$MET)
  auc_results$met_lower <- as.numeric(auc_results$met_lower)
  auc_results$met_upper <- as.numeric(auc_results$met_upper)
  auc_results$Pvalue_MET_vs_CONV <- as.numeric(auc_results$Pvalue_MET_vs_CONV)
  auc_results$COMBINED <- as.numeric(auc_results$COMBINED)
  auc_results$combined_lower <- as.numeric(auc_results$combined_lower)
  auc_results$combined_upper <- as.numeric(auc_results$combined_upper)
  auc_results$Diff <- as.numeric(auc_results$Diff)
  auc_results$Diff_proportion <- as.numeric(auc_results$Diff_proportion)
  auc_results$Pvalue_COMBINED_vs_CONV <- as.numeric(auc_results$Pvalue_COMBINED_vs_CONV)
  
  # 将 auc_results 添加到 auc_result_list 相应的数据框中
  auc_results_list[[model_name]] <- auc_results
  
  # 将数据写入工作簿的`model_name'工作表
  addWorksheet(wb, sheetName = model_name)
  writeData(wb, model_name, auc_results_list[[model_name]])
}


# Save the Excel file
# saveWorkbook(wb, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_预测_Transformer新结果/UKB关联交互预测_预测_总人群.xlsx", overwrite = TRUE)
saveWorkbook(wb, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_预测_Transformer新结果/UKB关联交互预测_预测_总人群_返修加模型.xlsx", overwrite = TRUE)

# 保存预测label
testdata_label <- testdata[,68:ncol(testdata)]
testdata_id <- testdata[,1:3]
testdata_predict_label <- cbind(testdata_id,testdata_label)

library(haven)
write_dta(testdata_predict_label, "~/000_2 文章/036 文章11-ISOS代谢组/UKB_testdata_predict_label_ysp20231228.dta")



# ** 亚组预测 ####

# 平衡亚组数据集
library(caret)
# 年龄亚组
table(traindata$a11)
set.seed(888)
traindata_age <- upSample(x = traindata[, 1:ncol(traindata)], y = traindata$a11)                         
table(traindata_age$a11) 

traindata_young <- subset(traindata_age, a11 == 0)
testdata_young <- subset(testdata, a11 == 0)
traindata_old <- subset(traindata_age, a11 == 1)
testdata_old <- subset(testdata, a11 == 1)

#性别亚组
table(traindata$a22)
set.seed(888)
traindata_sex <- upSample(x = traindata[, 1:ncol(traindata)], y = traindata$a22)                         
table(traindata_sex$a22) 

traindata_male <- subset(traindata_sex, a22 == 0)
testdata_male <- subset(testdata, a22 == 0)
traindata_female <- subset(traindata_sex, a22 == 1)
testdata_female <- subset(testdata, a22 == 1)

#种族亚组
table(traindata$a88)
set.seed(888)
traindata_ethnicity <- upSample(x = traindata[, 1:ncol(traindata)], y = traindata$a88)                         
table(traindata_ethnicity$a88) 

traindata_white <- subset(traindata_ethnicity, a88 == 0)
testdata_white <- subset(testdata, a88 == 0)
traindata_other <- subset(traindata_ethnicity, a88 == 1)
testdata_other <- subset(testdata, a88 == 1)

#剥夺亚组
table(traindata$a55)
set.seed(888)
traindata_deprivation <- upSample(x = traindata[, 1:ncol(traindata)], y = traindata$a55)                         
table(traindata_deprivation$a55) 

traindata_ndep <- subset(traindata_deprivation, a55 == 0)
testdata_ndep <- subset(testdata, a55 == 0)
traindata_dep <- subset(traindata_deprivation, a55 == 1)
testdata_dep <- subset(testdata, a55 == 1)

#教育亚组
table(traindata$a99)
set.seed(888)
traindata_education <- upSample(x = traindata[, 1:ncol(traindata)], y = traindata$a99)                         
table(traindata_education$a99) 

traindata_edu <- subset(traindata_education, a99 == 0)
testdata_edu <- subset(testdata, a99 == 0)
traindata_nedu <- subset(traindata_education, a99 == 1)
testdata_nedu <- subset(testdata, a99 == 1)

#BMI亚组
table(traindata$a1010)
set.seed(888)
traindata_bmi <- upSample(x = traindata[, 1:ncol(traindata)], y = traindata$a1010)                         
table(traindata_bmi$a1010) 

traindata_fit <- subset(traindata_bmi, a1010 == 0)
testdata_fit <- subset(testdata, a1010 == 0)
traindata_obes <- subset(traindata_bmi, a1010 == 1)
testdata_obes <- subset(testdata, a1010 == 1)


# 正式开始亚组预测
library(pROC)
library(survival)
library(openxlsx)
library(dplyr)

auc_results_list <- list()

wb <- createWorkbook()

models <- create_models_and_outcomes()$models
outcomes <- create_models_and_outcomes()$outcomes

auc_results_list <- list()

AGE <- createWorkbook()
SEX <- createWorkbook()
ETHNICITY <- createWorkbook()
TDI <- createWorkbook()
EDUCATION <- createWorkbook()
BMI <- createWorkbook()

subset_info <- list(
  "AGE" = list(traindata_subset1 = "traindata_young", testdata_subset1 = "testdata_young", traindata_subset2 = "traindata_old", testdata_subset2 = "testdata_old"),
  "SEX" = list(traindata_subset1 = "traindata_male", testdata_subset1 = "testdata_male", traindata_subset2 = "traindata_female", testdata_subset2 = "testdata_female"),
  "ETHNICITY" = list(traindata_subset1 = "traindata_white", testdata_subset1 = "testdata_white", traindata_subset2 = "traindata_other", testdata_subset2 = "testdata_other"),
  "TDI" = list(traindata_subset1 = "traindata_ndep", testdata_subset1 = "testdata_ndep", traindata_subset2 = "traindata_dep", testdata_subset2 = "testdata_dep"),
  "EDUCATION" = list(traindata_subset1 = "traindata_edu", testdata_subset1 = "testdata_edu", traindata_subset2 = "traindata_nedu", testdata_subset2 = "testdata_nedu"),
  "BMI" = list(traindata_subset1 = "traindata_fit", testdata_subset1 = "testdata_fit", traindata_subset2 = "traindata_obes", testdata_subset2 = "testdata_obes")
)


for (subset_name in names(subset_info)) {
  traindata_subset1 <- get(subset_info[[subset_name]]$traindata_subset1)
  testdata_subset1 <- get(subset_info[[subset_name]]$testdata_subset1)
  traindata_subset2 <- get(subset_info[[subset_name]]$traindata_subset2)
  testdata_subset2 <- get(subset_info[[subset_name]]$testdata_subset2)
  traindata_subset1_name <- subset_info[[subset_name]]$traindata_subset1
  traindata_subset2_name <- subset_info[[subset_name]]$traindata_subset2
  
  models <- create_models_and_outcomes()$models
  outcomes <- create_models_and_outcomes()$outcomes
  
  #如果是性别亚组，模型中删除性别
  if (subset_name == "SEX") {
    for (model_name in names(models)) {
      # 使用gsub删除模型字符串中的+ as.factor(a2)
      models[[model_name]] <- gsub("\\+ as\\.factor\\(a2\\)", "", models[[model_name]])
    }
  }
  
  #如果是种族亚组，模型中删除种族
  if (subset_name == "ETHNICITY") {
    for (model_name in names(models)) {
      # 使用gsub删除模型字符串中的+ as.factor(a8)和+ mau + pa_met  #mau和pa_met在other种族里过多缺失
      models[[model_name]] <- gsub("\\s*\\+\\s*as\\.factor\\(a8\\) | \\s*\\+\\s*mau | \\s*\\+\\s*pa_met", "", models[[model_name]], perl = TRUE)
    }
    # 删除"aaa"结局
    outcomes <- setdiff(outcomes, "aaa")
  }
  
  #如果是教育亚组，模型中删除教育
  if (subset_name == "EDUCATION") {
    for (model_name in names(models)) {
      # 使用gsub删除模型字符串中的+ as.factor(a9)
      models[[model_name]] <- gsub("\\+ as\\.factor\\(a9\\)", "", models[[model_name]])
    }
  }
  
  #如果是BMI亚组，模型中删除BMI
  if (subset_name == "BMI") {
    for (model_name in names(models)) {
      # 使用gsub删除模型字符串中的+ as.factor(a10)
      models[[model_name]] <- gsub("\\+ as\\.factor\\(a10\\)", "", models[[model_name]])
    }
  }
  
  for (model_name in names(models)) {
    auc_results <- data.frame(outcome = character(),
                              ci_lower1 = numeric(),
                              auc_value1 = numeric(),
                              ci_upper1 = numeric(),
                              ci_lower3 = numeric(),
                              auc_value3 = numeric(),
                              ci_upper3 = numeric(),
                              ci_lower2 = numeric(),
                              auc_value2 = numeric(),
                              ci_upper2 = numeric(),
                              p = numeric(),
                              stringsAsFactors = FALSE)
    
    for (outcome in outcomes) {
      covariate <- models[[model_name]]
      
      #如果是hf结局，模型中删除prev_hf
      if (outcome == "hf") {
        covariate <- gsub("\\+ as\\.factor\\(prev_cvd_hf\\)", "", covariate)
      }
      # #如果是t2d结局，模型中删除prev_dm
      # if (outcome == "t2d") {
      #   covariate <- gsub("\\+ as\\.factor\\(prev_dm\\)", "", covariate)
      # }
      #如果是asthma结局，模型中删除asthma
      if (outcome == "asthma") {
        covariate <- gsub("\\+ as\\.factor\\(prev_asthma\\)", "", covariate)
      }
      #如果是stroke结局，模型中删除prev_stroke
      if (outcome == "stroke") {
        covariate <- gsub("\\+ as\\.factor\\(prev_stroke\\)", "", covariate)
      }
      #如果是renal_disease结局，模型中删除prev_ckd
      if (outcome == "renal_disease") {
        covariate <- gsub("\\+ as\\.factor\\(prev_ckd\\)", "", covariate)
      }
      #如果是copd结局，模型中删除prev_copd
      if (outcome == "copd") {
        covariate <- gsub("\\+ as\\.factor\\(prev_copd\\)", "", covariate)
      }
      
      
      # 亚组1
      # 传统模型
      formula1 <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ ", covariate, sep = ""))
      model1 <- coxph(formula1, data = traindata_subset1)
      testdata_subset1[[paste(outcome, "_riskscore_1", sep = "")]] <- predict(model1, type = "risk", newdata = testdata_subset1)
      roc_result1 <- roc(testdata_subset1[[paste("incident_", outcome, sep = "")]] ~ testdata_subset1[[paste(outcome, "_riskscore_1", sep = "")]], data = testdata_subset1, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
      auc_value1 <- roc_result1$auc
      ci_lower1 <- roc_result1$ci[1]
      ci_upper1 <- roc_result1$ci[3]
      
      # 加入代谢物
      formula2 <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ ", covariate, "+ score_", outcome, sep = ""))
      model2 <- coxph(formula2, data = traindata_subset1)
      testdata_subset1[[paste(outcome, "_riskscore_2", sep = "")]] <- predict(model2, type = "risk", newdata = testdata_subset1)
      roc_result2 <- roc(testdata_subset1[[paste("incident_", outcome, sep = "")]] ~ testdata_subset1[[paste(outcome, "_riskscore_2", sep = "")]], data = testdata_subset1, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
      auc_value2 <- roc_result2$auc
      ci_lower2 <- roc_result2$ci[1]
      ci_upper2 <- roc_result2$ci[3]
      test1 <- roc.test(roc_result1,roc_result2) #Delong's test
      p1 <- test1[["p.value"]]
      auc_diff1 <- auc_value2 - auc_value1
      auc_improve_proportion1 <- auc_diff1 / auc_value1
      sig1 <- ifelse(p1 >= 0.05, "",
                     ifelse(p1 < 0.001, "***",
                            ifelse(p1 >= 0.001 & p1 < 0.01, "**", "*")))
      
      # 单独代谢物
      formula3 <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ score_", outcome, sep = ""))
      model3 <- coxph(formula3, data = traindata_subset1)
      testdata_subset1[[paste(outcome, "_riskscore_3", sep = "")]] <- predict(model3, type = "risk", newdata = testdata_subset1)
      roc_result3 <- roc(testdata_subset1[[paste("incident_", outcome, sep = "")]] ~ testdata_subset1[[paste(outcome, "_riskscore_3", sep = "")]], data = testdata_subset1, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
      auc_value3 <- roc_result3$auc
      ci_lower3 <- roc_result3$ci[1]
      ci_upper3 <- roc_result3$ci[3]
      
      # 亚组2
      # 传统模型
      formula4 <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ ", covariate, sep = ""))
      model4 <- coxph(formula4, data = traindata_subset2)
      testdata_subset2[[paste(outcome, "_riskscore_4", sep = "")]] <- predict(model4, type = "risk", newdata = testdata_subset2)
      roc_result4 <- roc(testdata_subset2[[paste("incident_", outcome, sep = "")]] ~ testdata_subset2[[paste(outcome, "_riskscore_4", sep = "")]], data = testdata_subset2, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
      auc_value4 <- roc_result4$auc
      ci_lower4 <- roc_result4$ci[1]
      ci_upper4 <- roc_result4$ci[3]
      
      # 加入代谢物
      formula5 <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ ", covariate, "+ score_", outcome, sep = ""))
      model5 <- coxph(formula5, data = traindata_subset2)
      testdata_subset2[[paste(outcome, "_riskscore_5", sep = "")]] <- predict(model5, type = "risk", newdata = testdata_subset2)
      roc_result5 <- roc(testdata_subset2[[paste("incident_", outcome, sep = "")]] ~ testdata_subset2[[paste(outcome, "_riskscore_5", sep = "")]], data = testdata_subset2, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
      auc_value5 <- roc_result5$auc
      ci_lower5 <- roc_result5$ci[1]
      ci_upper5 <- roc_result5$ci[3]
      test2 <- roc.test(roc_result4,roc_result5) #Delong's test
      p2 <- test2[["p.value"]]
      auc_diff2 <- auc_value5 - auc_value4
      auc_improve_proportion2 <- auc_diff2 / auc_value4
      sig2 <- ifelse(p2 >= 0.05, "",
                     ifelse(p2 < 0.001, "***",
                            ifelse(p2 >= 0.001 & p2 < 0.01, "**", "*")))
      
      # 单独代谢物
      formula6 <- as.formula(paste("Surv(followup_", outcome, ", incident_", outcome, ") ~ score_", outcome, sep = ""))
      model6 <- coxph(formula6, data = traindata_subset2)
      testdata_subset2[[paste(outcome, "_riskscore_6", sep = "")]] <- predict(model6, type = "risk", newdata = testdata_subset2)
      roc_result6 <- roc(testdata_subset2[[paste("incident_", outcome, sep = "")]] ~ testdata_subset2[[paste(outcome, "_riskscore_6", sep = "")]], data = testdata_subset2, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
      auc_value6 <- roc_result6$auc
      ci_lower6 <- roc_result6$ci[1]
      ci_upper6 <- roc_result6$ci[3]
      
      # 将结果添加到 auc_results 数据框中，并初步改名
      auc_results <- rbind(auc_results, c(outcome, auc_value1, ci_lower1, ci_upper1, auc_value3, ci_lower3, ci_upper3, auc_value2, ci_lower2, ci_upper2, auc_diff1, auc_improve_proportion1, p1, sig1,
                                          outcome, auc_value4, ci_lower4, ci_upper4, auc_value6, ci_lower6, ci_upper6, auc_value5, ci_lower5, ci_upper5, auc_diff2, auc_improve_proportion2, p2, sig2))
      colnames(auc_results) <- c("outcome", "auc_value1", "ci_lower1", "ci_upper1", "auc_value3", "ci_lower3", "ci_upper3", "auc_value2", "ci_lower2", "ci_upper2", "auc_diff1", "auc_improve_proportion1", "p1", "sig1",
                                 "outcome", "auc_value4", "ci_lower4", "ci_upper4", "auc_value6", "ci_lower6", "ci_upper6", "auc_value5", "ci_lower5", "ci_upper5", "auc_diff2", "auc_improve_proportion2", "p2", "sig2")
    }
    
    #如果是种族亚组，模型中无aaa，补充aaa空行
    if (subset_name == "ETHNICITY") {
      aaa_na <- c("aaa", NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,
                  "aaa", NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA)
      auc_results <- rbind(auc_results[1:8,], aaa_na, auc_results[9:15,])
    }
    
    # 增加一列结局分类
    Category <- create_models_and_outcomes()$category
    auc_results <- cbind(auc_results, Category)
    
    #重命名列名、输出的表格列排序
    outcomes_colname <- create_models_and_outcomes()$outcomes
    auc_results <- data.frame(outcome_superior = paste(outcomes_colname, strsplit(traindata_subset1_name, "_")[[1]][2], sep="_"), #每个outcome加上“_亚组1”
                              CONV_superior = as.numeric(auc_results$auc_value1),
                              conv_lower_superior = as.numeric(auc_results$ci_lower1),
                              conv_upper_superior = as.numeric(auc_results$ci_upper1),
                              MET_superior = as.numeric(auc_results$auc_value3),
                              met_lower_superior = as.numeric(auc_results$ci_lower3),
                              met_upper_superior = as.numeric(auc_results$ci_upper3),
                              COMBINED_superior = as.numeric(auc_results$auc_value2),
                              combined_lower_superior = as.numeric(auc_results$ci_lower2),
                              combined_upper_superior = as.numeric(auc_results$ci_upper2),
                              Diff_superior = as.numeric(auc_results$auc_diff1),
                              Diff_proportion_superior = as.numeric(auc_results$auc_improve_proportion1),
                              Pvalue_superior = as.numeric(auc_results$p1),
                              Significance_superior = auc_results$sig1,
                              
                              outcome_inferior = paste(outcomes_colname, strsplit(traindata_subset2_name, "_")[[1]][2], sep="_"), #每个outcome加上“_亚组2”
                              CONV_inferior = as.numeric(auc_results$auc_value4),
                              conv_lower_inferior = as.numeric(auc_results$ci_lower4),
                              conv_upper_inferior = as.numeric(auc_results$ci_upper4),
                              MET_inferior = as.numeric(auc_results$auc_value6),
                              met_lower_inferior = as.numeric(auc_results$ci_lower6),
                              met_upper_inferior = as.numeric(auc_results$ci_upper6),
                              COMBINED_inferior = as.numeric(auc_results$auc_value5),
                              combined_lower_inferior = as.numeric(auc_results$ci_lower5),
                              combined_upper_inferior = as.numeric(auc_results$ci_upper5),
                              Diff_inferior = as.numeric(auc_results$auc_diff2),
                              Diff_proportion_inferior = as.numeric(auc_results$auc_improve_proportion2),
                              Pvalue_inferior = as.numeric(auc_results$p2),
                              Significance_inferior = auc_results$sig2,
                              
                              SigCompare = ifelse(as.numeric(auc_results$p1) < 0.05 & as.numeric(auc_results$p2) < 0.05, "BOTH",
                                                  ifelse(as.numeric(auc_results$p1) < 0.05 & as.numeric(auc_results$p2) >= 0.05, "Superior only",
                                                         ifelse(as.numeric(auc_results$p1) >=0.05 & as.numeric(auc_results$p2) <0.05, "Inferior only",
                                                                "-"))),
                              Category = auc_results$Category,
                              DiffCompare = ifelse(as.numeric(auc_results$p1) < 0.05 & as.numeric(auc_results$p2) < 0.05, ifelse(as.numeric(auc_results$auc_diff2) >= as.numeric(auc_results$auc_diff1), "Inferior", "Superior"),
                                                   ifelse(as.numeric(auc_results$p1) < 0.05 & as.numeric(auc_results$p2) >= 0.05, "Superior",
                                                          ifelse(as.numeric(auc_results$p1) >= 0.05 & as.numeric(auc_results$p2) < 0.05, "Inferior",
                                                                 "-"))),
                              Diff_value_compare = ifelse(as.numeric(auc_results$auc_diff2) >= as.numeric(auc_results$auc_diff1), "Inferior", "Superior"))
    
    # 将 auc_results 添加到 auc_result_list 相应的数据框中
    auc_results_list[[model_name]] <- auc_results
    
    # 将数据写入工作簿的`model_name'工作表
    addWorksheet(get(subset_name), sheetName = model_name)
    writeData(get(subset_name), model_name, auc_results_list[[model_name]])
  }
}






# ** 导出结果-亚组预测 ####
# Save the Excel file
saveWorkbook(AGE, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_预测_Transformer新结果/UKB预测_年龄.xlsx", overwrite = TRUE)
saveWorkbook(SEX, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_预测_Transformer新结果/UKB预测_性别.xlsx", overwrite = TRUE)
saveWorkbook(ETHNICITY, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_预测_Transformer新结果/UKB预测_预测_种族.xlsx", overwrite = TRUE)
saveWorkbook(TDI, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_预测_Transformer新结果/UKB预测_剥夺.xlsx", overwrite = TRUE)
saveWorkbook(EDUCATION, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_预测_Transformer新结果/UKB预测_教育.xlsx", overwrite = TRUE)
saveWorkbook(BMI, "~/000_2 文章/036 文章11-ISOS代谢组/UKB关联交互预测_预测_Transformer新结果/UKB预测_身材.xlsx", overwrite = TRUE)





# * DCA ####
# 清除除了testdata以外的所有数据
objects_to_keep <- c("testdata","create_models_and_outcomes")  #要保留的对象名称
objects_to_remove <- setdiff(ls(), objects_to_keep)
rm(list = objects_to_remove)

# 开始
library(rmda)
models <- create_models_and_outcomes()$models
outcomes <- create_models_and_outcomes()$outcomes

for (model_name in names(models)) {
  # 使用gsub删除模型字符串中的+ as.factor(prev_dementia)
  models[[model_name]] <- gsub("\\+ as\\.factor\\(prev_dementia\\)", "", models[[model_name]])
}

# 存储DCA计算结果
results_baseline <- list()
results_full <- list()
for (outcome in outcomes) {
  
  if (outcome == "death" | outcome == "cancerdeath" | outcome == "otherdeath") {
    model_names <- c("Age_Sex","Liao_model","Chiu_model","Mannan_model","Li_model")
  }
  if (outcome == "chd" | outcome == "t2d" | outcome == "mi" | outcome == "stroke" | outcome == "pad" | outcome == "aaa" | outcome =="hf") {
    model_names <- c("Age_Sex","FGCRS","SCORE2","AHA_ASCVD","WHO_CVD")
  }
  if (outcome == "renal_disease" | outcome == "esrd") {
    model_names <- c("Age_Sex","Nelson_model","KFRE","Chien_model","O_Seaghdha_model")
  }
  if (outcome == "copd" | outcome == "asthma") {
    model_names <- c("Age_Sex","Kotz_model","EHS_COPD","BARC_index","Himes_model")
  }
  if (outcome == "liver_disease") {
    model_names <- c("Age_Sex","CLivD_score","Zhu_model","Zhang_model","Xue_model")
  }
  if (outcome == "lungca") {
    model_names <- c("Age_Sex","CanPredict_lung","LLP","PLMO","LCRAT")
  }
  
  for (model_name in model_names) {
    
    
    #baseline
    formula <- formula(paste("incident_", outcome, "~", models[[model_name]], sep = ""))
    result <- decision_curve(formula, data = testdata, thresholds = seq(0, 1, by = .01))
    results_baseline[[outcome]][[model_name]] <- result
    
    #full
    formula <- formula(paste("incident_", outcome, "~", models[[model_name]], "+ score_", outcome, sep = ""))
    result <- decision_curve(formula, data = testdata, thresholds = seq(0, 1, by = .01))
    results_full[[outcome]][[model_name]] <- result
  }
  gc()
}

library(rmda)
pdf("~/000_2 文章/036 文章11-ISOS代谢组/图表/UKB_DCA.pdf", width = 20/1.5, height = 25/1.5)
par(mfrow = c(4,4))  # 将图分成一行两列
plot_decision_curve(list(results_baseline$death$Age_Sex, results_full$death$Age_Sex, results_baseline$death$Liao_model, results_full$death$Liao_model, results_baseline$death$Chiu_model, results_full$death$Chiu_model, results_baseline$death$Mannan_model, results_full$death$Mannan_model, results_baseline$death$Li_model, results_full$death$Li_model),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "Liao model", "Liao model + PMW", "Chiu model", "Chiu model + PMW", "Mannan model", "Mannan model + PMW", "Li model", "Li model + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.4),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("All-cause mortality")

plot_decision_curve(list(results_baseline$cancerdeath$Age_Sex, results_full$cancerdeath$Age_Sex, results_baseline$cancerdeath$Liao_model, results_full$cancerdeath$Liao_model, results_baseline$cancerdeath$Chiu_model, results_full$cancerdeath$Chiu_model, results_baseline$cancerdeath$Mannan_model, results_full$cancerdeath$Mannan_model, results_baseline$cancerdeath$Li_model, results_full$cancerdeath$Li_model),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "Liao model", "Liao model + PMW", "Chiu model", "Chiu model + PMW", "Mannan model", "Mannan model + PMW", "Li model", "Li model + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.25),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("Cancer mortality")

plot_decision_curve(list(results_baseline$otherdeath$Age_Sex, results_full$otherdeath$Age_Sex, results_baseline$otherdeath$Liao_model, results_full$otherdeath$Liao_model, results_baseline$otherdeath$Chiu_model, results_full$otherdeath$Chiu_model, results_baseline$otherdeath$Mannan_model, results_full$otherdeath$Mannan_model, results_baseline$otherdeath$Li_model, results_full$otherdeath$Li_model),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "Liao model", "Liao model + PMW", "Chiu model", "Chiu model + PMW", "Mannan model", "Mannan model + PMW", "Li model", "Li model + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.2),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("Other mortality")


plot_decision_curve(list(results_baseline$mi$Age_Sex, results_full$mi$Age_Sex, results_baseline$mi$FGCRS, results_full$mi$FGCRS, results_baseline$mi$SCORE2, results_full$mi$SCORE2, results_baseline$mi$AHA_ASCVD, results_full$mi$AHA_ASCVD, results_baseline$mi$WHO_CVD, results_full$mi$WHO_CVD),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "FGCRS", "FGCRS + PMW", "SCORE2", "SCORE2 + PMW", "AHA/ASCVD", "AHA/ASCVD + PMW", "WHO-CVD", "WHO-CVD + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.3),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("Myocardial infarction")

plot_decision_curve(list(results_baseline$stroke$Age_Sex, results_full$stroke$Age_Sex, results_baseline$stroke$FGCRS, results_full$stroke$FGCRS, results_baseline$stroke$SCORE2, results_full$stroke$SCORE2, results_baseline$stroke$AHA_ASCVD, results_full$stroke$AHA_ASCVD, results_baseline$stroke$WHO_CVD, results_full$stroke$WHO_CVD),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "FGCRS", "FGCRS + PMW", "SCORE2", "SCORE2 + PMW", "AHA/ASCVD", "AHA/ASCVD + PMW", "WHO-CVD", "WHO-CVD + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.15),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("Stroke")

plot_decision_curve(list(results_baseline$hf$Age_Sex, results_full$hf$Age_Sex, results_baseline$hf$FGCRS, results_full$hf$FGCRS, results_baseline$hf$SCORE2, results_full$hf$SCORE2, results_baseline$hf$AHA_ASCVD, results_full$hf$AHA_ASCVD, results_baseline$hf$WHO_CVD, results_full$hf$WHO_CVD),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "FGCRS", "FGCRS + PMW", "SCORE2", "SCORE2 + PMW", "AHA/ASCVD", "AHA/ASCVD + PMW", "WHO-CVD", "WHO-CVD + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.2),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("Heart failure")

plot_decision_curve(list(results_baseline$chd$Age_Sex, results_full$chd$Age_Sex, results_baseline$chd$FGCRS, results_full$chd$FGCRS, results_baseline$chd$SCORE2, results_full$chd$SCORE2, results_baseline$chd$AHA_ASCVD, results_full$chd$AHA_ASCVD, results_baseline$chd$WHO_CVD, results_full$chd$WHO_CVD),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "FGCRS", "FGCRS + PMW", "SCORE2", "SCORE2 + PMW", "AHA/ASCVD", "AHA/ASCVD + PMW", "WHO-CVD", "WHO-CVD + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.35),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")

title("CHD")

plot_decision_curve(list(results_baseline$aaa$Age_Sex, results_full$aaa$Age_Sex, results_baseline$aaa$FGCRS, results_full$aaa$FGCRS, results_baseline$aaa$SCORE2, results_full$aaa$SCORE2, results_baseline$aaa$AHA_ASCVD, results_full$aaa$AHA_ASCVD, results_baseline$aaa$WHO_CVD, results_full$aaa$WHO_CVD),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "FGCRS", "FGCRS + PMW", "SCORE2", "SCORE2 + PMW", "AHA/ASCVD", "AHA/ASCVD + PMW", "WHO-CVD", "WHO-CVD + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.05),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("AAA")

plot_decision_curve(list(results_baseline$pad$Age_Sex, results_full$pad$Age_Sex, results_baseline$pad$FGCRS, results_full$pad$FGCRS, results_baseline$pad$SCORE2, results_full$pad$SCORE2, results_baseline$pad$AHA_ASCVD, results_full$pad$AHA_ASCVD, results_baseline$pad$WHO_CVD, results_full$pad$WHO_CVD),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "FGCRS", "FGCRS + PMW", "SCORE2", "SCORE2 + PMW", "AHA/ASCVD", "AHA/ASCVD + PMW", "WHO-CVD", "WHO-CVD + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.15),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")

title("PAD")

plot_decision_curve(list(results_baseline$t2d$Age_Sex, results_full$t2d$Age_Sex, results_baseline$t2d$FGCRS, results_full$t2d$FGCRS, results_baseline$t2d$SCORE2, results_full$t2d$SCORE2, results_baseline$t2d$AHA_ASCVD, results_full$t2d$AHA_ASCVD, results_baseline$t2d$WHO_CVD, results_full$t2d$WHO_CVD),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "FGCRS", "FGCRS + PMW", "SCORE2", "SCORE2 + PMW", "AHA/ASCVD", "AHA/ASCVD + PMW", "WHO-CVD", "WHO-CVD + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.835),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("T2D")

plot_decision_curve(list(results_baseline$renal_disease$Age_Sex, results_full$renal_disease$Age_Sex, results_baseline$renal_disease$Nelson_model, results_full$renal_disease$Nelson_model, results_baseline$renal_disease$KFRE, results_full$renal_disease$KFRE, results_baseline$renal_disease$Chien_model, results_full$renal_disease$Chien_model, results_baseline$renal_disease$O_Seaghdha_model, results_full$renal_disease$O_Seaghdha_model),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "Nelson model", "Nelson model + PMW", "KFRE", "KFRE + PMW", "Chien model", "Chien model + PMW", "O'Seaghdha model", "O'Seaghdha model + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.8),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("Renal disease")

plot_decision_curve(list(results_baseline$esrd$Age_Sex, results_full$esrd$Age_Sex, results_baseline$esrd$Nelson_model, results_full$esrd$Nelson_model, results_baseline$esrd$KFRE, results_full$esrd$KFRE, results_baseline$esrd$Chien_model, results_full$esrd$Chien_model, results_baseline$esrd$O_Seaghdha_model, results_full$esrd$O_Seaghdha_model),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "Nelson model", "Nelson model + PMW", "KFRE", "KFRE + PMW", "Chien model", "Chien model + PMW", "O'Seaghdha model", "O'Seaghdha model + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.69),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("ESRD")

plot_decision_curve(list(results_baseline$copd$Age_Sex, results_full$copd$Age_Sex, results_baseline$copd$Kotz_model, results_full$copd$Kotz_model, results_baseline$copd$EHS_COPD, results_full$copd$EHS_COPD, results_baseline$copd$BARC_index, results_full$copd$BARC_index, results_baseline$copd$Himes_model, results_full$copd$Himes_model),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "Kotz_model", "Kotz_model + PMW", "EHS_COPD", "EHS_COPD + PMW", "BARC_index", "BARC_index + PMW", "Himes_model", "Himes_model + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.5),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("COPD")

plot_decision_curve(list(results_baseline$asthma$Age_Sex, results_full$asthma$Age_Sex, results_baseline$asthma$Kotz_model, results_full$asthma$Kotz_model, results_baseline$asthma$EHS_COPD, results_full$asthma$EHS_COPD, results_baseline$asthma$BARC_index, results_full$asthma$BARC_index, results_baseline$asthma$Himes_model, results_full$asthma$Himes_model),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "Kotz_model", "Kotz_model + PMW", "EHS_COPD", "EHS_COPD + PMW", "BARC_index", "BARC_index + PMW", "Himes_model", "Himes_model + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.6),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("Asthma")

plot_decision_curve(list(results_baseline$liver_disease$Age_Sex, results_full$liver_disease$Age_Sex, results_baseline$liver_disease$CLivD_score, results_full$liver_disease$CLivD_score, results_baseline$liver_disease$Zhu_model, results_full$liver_disease$Zhu_model, results_baseline$liver_disease$Zhang_model, results_full$liver_disease$Zhang_model, results_baseline$liver_disease$Xue_model, results_full$liver_disease$Xue_model),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "CLivD score", "CLivD score + PMW", "Zhu model", "Zhu model + PMW", "Zhang model", "Zhang model + PMW", "Xue model", "Xue model + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.2),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("Hepatic disease")

plot_decision_curve(list(results_baseline$lungca$Age_Sex, results_full$lungca$Age_Sex, results_baseline$lungca$CanPredict_lung, results_full$lungca$CanPredict_lung, results_baseline$lungca$LLP, results_full$lungca$LLP, results_baseline$lungca$PLMO, results_full$lungca$PLMO, results_baseline$lungca$LCRAT, results_full$lungca$LCRAT),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "CanPredict (lung)", "CanPredict (lung) + PMW", "LLP", "LLP + PMW", "PLMO", "PLMO + PMW", "LCRAT", "LCRAT + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.25),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("Lung cancer")

dev.off()





rm(list = ls())
library(haven)
data_all <- read_dta("xxx.dta")
data_NMR <- data_all[!is.na(data_all$TPTG), ]
# NMR清单
nmr_met_list <- names(data_NMR)[which(names(data_NMR) == "TPTG"):which(names(data_NMR) == "Dimethylsulfone")]



# 创建一个列表，用于存储每个结局的logit结果
logit_result_list <- list()
library(openxlsx)
wb <- createWorkbook()
subfields <- c("mean", "average", "inner_average", "outer_average", "central")
outcomes <- c("cvd","mortality","dr_progression","vtdr","ckd")


# 外层循环：遍历不同亚场的ELM-ISOS
for (subfield in subfields) {
  # 创建一个空的数据框来存储回归结果
  logit_result <- data.frame()
  
  # 内层循循环：遍历结局
  for (outcome in outcomes) {
    # 构建公式
    formula <- as.formula(paste0(outcome, " ~ age_baseline + as.factor(sex) + height + weight + se + isos_", subfield, "_baseline_sd_x"))
    logit_model <- glm(formula, data = data_all, family = binomial(link = "logit"))
    
    coef <- as.numeric(summary(logit_model)$coefficients[, 1][nrow(summary(logit_model)$coefficients)])
    std <- as.numeric(summary(logit_model)$coefficients[, 2][nrow(summary(logit_model)$coefficients)])
    zvalue <- as.numeric(summary(logit_model)$coefficients[, 3][nrow(summary(logit_model)$coefficients)])
    HR <- exp(coef)
    HR_low <- exp(coef - 2*std)
    HR_up <- exp(coef + 2*std)
    pvalue <- as.numeric(summary(logit_model)$coefficients[, 4][nrow(summary(logit_model)$coefficients)])
    logit_result <- rbind(logit_result, c(outcome, coef, std, zvalue, HR, HR_low, HR_up, pvalue))
    
    colnames(logit_result) <- c("outcome", "coef", "std", "zvalue", "HR", "HR_low", "HR_up", "pvalue")
    logit_result$coef <- as.numeric(logit_result$coef)
    logit_result$std <- as.numeric(logit_result$std)
    logit_result$zvalue <- as.numeric(logit_result$zvalue)
    logit_result$HR <- as.numeric(logit_result$HR)
    logit_result$HR_low <- as.numeric(logit_result$HR_low)
    logit_result$HR_up <- as.numeric(logit_result$HR_up)
    logit_result$pvalue <- as.numeric(logit_result$pvalue)
    
    logit_result <- logit_result[c("outcome", "coef", "std", "zvalue", "HR", "HR_low", "HR_up", "pvalue")]
    logit_result$Sig <- ifelse(logit_result$pvalue >= 0.05, "",
                               ifelse(logit_result$pvalue < 0.001, "***",
                                      ifelse(logit_result$pvalue >= 0.001 & logit_result$pvalue < 0.01, "**", "*")))
    logit_result$BH <- p.adjust(logit_result$pvalue, "BH")
  }
  
  logit_result_list[[subfield]] <- logit_result
  addWorksheet(wb, sheetName = subfield)
  writeData(wb, subfield, logit_result_list[[subfield]])
}
logit_result_list

saveWorkbook(wb, "xxx.xlsx", overwrite = TRUE)



rm(list = ls())
library(haven)
data_all <- read_dta("xxx.dta")
data_NMR <- data_all[!is.na(data_all$TPTG), ]

lm_results_list <- list()
met_filtered_list <- list()

# 亚场清单
variables <- c("mean", "central", "inner_average", "outer_average")

library(openxlsx)
wb <- createWorkbook()

for (var in variables) {
  lm_result <- data.frame()
  
  # 循环每个代谢物并存储结果
  for (var_m in nmr500_met_list) {
    formula <- paste0("isos_", var, "_baseline_sd_y ~", var_m, " + age + as.factor(sex) + height + weight + se")
    lm_model <- lm(formula, data = data_NMR)
    
    coef <- as.numeric(summary(lm_model)$coefficients[, 1][2])
    std <- as.numeric(summary(lm_model)$coefficients[, 2][2])
    tvalue <- as.numeric(summary(lm_model)$coefficients[, 3][2])
    pvalue <- as.numeric(summary(lm_model)$coefficients[, 4][2])
    
    lm_result <- rbind(lm_result, c(var_m, coef, tvalue, std, pvalue))
  }
  
  colnames(lm_result) <- c("met", "coef", "tvalue", "std", "pvalue")
  lm_result$coef <- as.numeric(lm_result$coef)
  lm_result$std <- as.numeric(lm_result$std)
  lm_result$tvalue <- as.numeric(lm_result$tvalue)
  lm_result$pvalue <- as.numeric(lm_result$pvalue)
  lm_result$low_limit <- as.numeric(lm_result$coef) - 1.96 * as.numeric(lm_result$std)
  lm_result$up_limit <- as.numeric(lm_result$coef) + 1.96 * as.numeric(lm_result$std)
  lm_result <- lm_result[c("met", "coef", "low_limit", "up_limit", "std", "tvalue", "pvalue")]
  
  lm_result$BH <- p.adjust(lm_result$pvalue, "BH")
  
  #将每一个lm_result保存到相应的lm_results_list中
  lm_results_list[[var]] <- lm_result
  addWorksheet(wb, var)
  writeData(wb, var, lm_results_list[[var]])
  
  #筛选显著代谢物
  met_filtered_list[[var]] <- lm_results_list[[var]][lm_results_list[[var]]$BH < 0.05, ]
  addWorksheet(wb, paste0(var, "显著 ", nrow(met_filtered_list[[var]])))
  writeData(wb, paste0(var, "显著 ", nrow(met_filtered_list[[var]])), met_filtered_list[[var]])
}


saveWorkbook(wb, "xxx.xlsx", overwrite = TRUE)





rm(list = ls())
library(haven)
data_all <- read_dta("x.dta")
data_NMR <- data_all[!is.na(data_all$TPTG), ]

logit_result_list <- list()
library(openxlsx)
wb <- createWorkbook()
outcomes <- c("cvd","mortality","dr_progression","vtdr","ckd")

# 代谢物列表
result_mean <- read.xlsx("xxx.xlsx", sheet = "mean")
coef_mean <- data.frame(metabolite = result_mean$met, coef_mean = result_mean$coef)
result_filter_mean <- subset(result_mean, sig == "Significant" & direction == "Replicated")
metlist_mean <- unique(result_filter_mean$met)

result_central <- read.xlsx("xxx.xlsx", sheet = "central")
coef_central <- data.frame(metabolite = result_central$met, coef_central = result_central$coef)
result_filter_central <- subset(result_central, sig == "Significant" & direction == "Replicated")
metlist_central <- unique(result_filter_central$met)

result_inner <- read.xlsx("xxx.xlsx", sheet = "inner_average")
coef_inner <- data.frame(metabolite = result_inner$met, coef_inner = result_inner$coef)
result_filter_inner <- subset(result_inner, sig == "Significant" & direction == "Replicated")
metlist_inner <- unique(result_filter_inner$met)

result_outer <- read.xlsx("xxx.xlsx", sheet = "outer_average")
coef_outer <- data.frame(metabolite = result_outer$met, coef_outer = result_outer$coef)
result_filter_outer <- subset(result_outer, sig == "Significant" & direction == "Replicated")
metlist_outer <- unique(result_filter_outer$met)

metlist_all <- unique(unlist(list(metlist_mean, metlist_central, metlist_inner, metlist_outer)))

# 外层循环：遍历不同结局
for (outcome in outcomes) {
  # 创建一个空的数据框来存储回归结果
  logit_result <- data.frame()
  
  # 内层循循环：遍历结局
  for (metabolite in metlist_all) {
    # 构建公式
    formula <- as.formula(paste0(outcome, " ~ age + as.factor(sex) + height + weight + se + as.factor(income) + as.factor(education) + as.factor(smoking) + as.factor(drinking) + as.factor(med_lipid) + as.factor(med_hbp) + ", metabolite))
    logit_model <- glm(formula, data = data_all, family = binomial(link = "logit"))
    
    coef <- as.numeric(summary(logit_model)$coefficients[, 1][nrow(summary(logit_model)$coefficients)])
    std <- as.numeric(summary(logit_model)$coefficients[, 2][nrow(summary(logit_model)$coefficients)])
    zvalue <- as.numeric(summary(logit_model)$coefficients[, 3][nrow(summary(logit_model)$coefficients)])
    HR <- exp(coef)
    HR_low <- exp(coef - 2*std)
    HR_up <- exp(coef + 2*std)
    pvalue <- as.numeric(summary(logit_model)$coefficients[, 4][nrow(summary(logit_model)$coefficients)])
    logit_result <- rbind(logit_result, c(metabolite, coef, std, zvalue, HR, HR_low, HR_up, pvalue))
    
    colnames(logit_result) <- c("metabolite", "coef", "std", "zvalue", "HR", "HR_low", "HR_up", "pvalue")
    logit_result$coef <- as.numeric(logit_result$coef)
    logit_result$std <- as.numeric(logit_result$std)
    logit_result$zvalue <- as.numeric(logit_result$zvalue)
    logit_result$HR <- as.numeric(logit_result$HR)
    logit_result$HR_low <- as.numeric(logit_result$HR_low)
    logit_result$HR_up <- as.numeric(logit_result$HR_up)
    logit_result$pvalue <- as.numeric(logit_result$pvalue)
    
    logit_result <- logit_result[c("metabolite", "coef", "std", "zvalue", "HR", "HR_low", "HR_up", "pvalue")]
    logit_result$Sig <- ifelse(logit_result$pvalue >= 0.05, "",
                               ifelse(logit_result$pvalue < 0.001, "***",
                                      ifelse(logit_result$pvalue >= 0.001 & logit_result$pvalue < 0.01, "**", "*")))
    #给代谢物列表加列名
    logit_result <- left_join(logit_result, metname, by = "metabolite")
    #合并ISOS-代谢物关联
    logit_result <- left_join(logit_result, coef_mean, by = "metabolite")
    logit_result <- left_join(logit_result, coef_central, by = "metabolite")
    logit_result <- left_join(logit_result, coef_inner, by = "metabolite")
    logit_result <- left_join(logit_result, coef_outer, by = "metabolite")
  }
  
  logit_result_list[[outcome]] <- logit_result
  addWorksheet(wb, sheetName = outcome)
  writeData(wb, outcome, logit_result_list[[outcome]])
}
logit_result_list

saveWorkbook(wb, "xxx.xlsx", overwrite = TRUE)



rm(list = ls())
#标准化
numeric_cols <- which(sapply(score, is.numeric) & !(names(score) %in% c("id", "dataset_cvd", "dataset_dr_progression", "dataset_vtdr", "dataset_ckd"))) # 找到所有数值列的索引
score[, numeric_cols] <- scale(score[, numeric_cols]) # 标准化除了id列以外的所有数值列

outcomes <- c("cvd","dr_progression","vtdr","ckd")

logit_result <- data.frame()

for (outcome in outcomes) {
  # 构建公式
  formula <- as.formula(paste0(outcome, " ~ age + as.factor(sex) + height + weight + se + as.factor(income) + as.factor(education) + as.factor(smoking) + as.factor(drinking) + as.factor(med_lipid) + as.factor(med_hbp) +  score_", outcome))
  data <- get(paste0("dataset_", outcome))
  logit_model <- glm(formula, data, family = binomial(link = "logit"))
  
  coef <- as.numeric(summary(logit_model)$coefficients[, 1][nrow(summary(logit_model)$coefficients)])
  std <- as.numeric(summary(logit_model)$coefficients[, 2][nrow(summary(logit_model)$coefficients)])
  zvalue <- as.numeric(summary(logit_model)$coefficients[, 3][nrow(summary(logit_model)$coefficients)])
  HR <- exp(coef)
  HR_low <- exp(coef - 2*std)
  HR_up <- exp(coef + 2*std)
  pvalue <- as.numeric(summary(logit_model)$coefficients[, 4][nrow(summary(logit_model)$coefficients)])
  logit_result <- rbind(logit_result, c(outcome, coef, std, zvalue, HR, HR_low, HR_up, pvalue))
  
  colnames(logit_result) <- c("outcome", "coef", "std", "zvalue", "HR", "HR_low", "HR_up", "pvalue")
  logit_result$coef <- as.numeric(logit_result$coef)
  logit_result$std <- as.numeric(logit_result$std)
  logit_result$zvalue <- as.numeric(logit_result$zvalue)
  logit_result$HR <- as.numeric(logit_result$HR)
  logit_result$HR_low <- as.numeric(logit_result$HR_low)
  logit_result$HR_up <- as.numeric(logit_result$HR_up)
  logit_result$pvalue <- as.numeric(logit_result$pvalue)
  
  logit_result <- logit_result[c("outcome", "coef", "std", "zvalue", "HR", "HR_low", "HR_up", "pvalue")]
  logit_result$Sig <- ifelse(logit_result$pvalue >= 0.05, "",
                             ifelse(logit_result$pvalue < 0.001, "***",
                                    ifelse(logit_result$pvalue >= 0.001 & logit_result$pvalue < 0.01, "**", "*")))
}
logit_result

library(openxlsx)
write.xlsx(logit_result, "xxx.xlsx")






# * 单独predictor比较 ####
library(pROC)

auc_results <- data.frame()

for (outcome in outcomes) {
  # 划分数据集
  data_train <- get(paste0("dataset_", outcome, "_train"))
  data_test <- get(paste0("dataset_", outcome, "_test"))
  
  # 单独代谢物
  formula_met <- as.formula(paste(outcome, "~ score_", outcome, sep = ""))
  model_met <- glm(formula_met, data = data_train, family = binomial)
  data_test[[paste(outcome, "_riskscore_met", sep = "")]] <- predict(model_met, newdata = data_test)
  roc_result_met <- roc(data_test[[outcome]] ~ data_test[[paste(outcome, "_riskscore_met", sep = "")]], data = data_test, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_met <- roc_result_met$auc
  ci_lower_met <- roc_result_met$ci[1]
  ci_upper_met <- roc_result_met$ci[3]
  
  # 年龄
  formula_age <- as.formula(paste(outcome, "~ age", sep = ""))
  model_age <- glm(formula_age, data = data_train, family = binomial)
  data_test[[paste(outcome, "_riskscore_age", sep = "")]] <- predict(model_age, newdata = data_test)
  roc_result_age <- roc(data_test[[outcome]] ~ data_test[[paste(outcome, "_riskscore_age", sep = "")]], data = data_test, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_age <- roc_result_age$auc
  ci_lower_age <- roc_result_age$ci[1]
  ci_upper_age <- roc_result_age$ci[3]
  test_age <- roc.test(roc_result_met,roc_result_age) #Delong's test
  p_age <- test_age[["p.value"]]
  compare_age <- ifelse(p_age >= 0.05, "Comparable", ifelse(auc_met > auc_age, "Superior", "-"))
  
  # 性别
  formula_sex <- as.formula(paste(outcome, "~ as.factor(sex)", sep = ""))
  model_sex <- glm(formula_sex, data = data_train, family = binomial)
  data_test[[paste(outcome, "_riskscore_sex", sep = "")]] <- predict(model_sex, newdata = data_test)
  roc_result_sex <- roc(data_test[[outcome]] ~ data_test[[paste(outcome, "_riskscore_sex", sep = "")]], data = data_test, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_sex <- roc_result_sex$auc
  ci_lower_sex <- roc_result_sex$ci[1]
  ci_upper_sex <- roc_result_sex$ci[3]
  test_sex <- roc.test(roc_result_met,roc_result_sex) #Delong's test
  p_sex <- test_sex[["p.value"]]
  compare_sex <- ifelse(p_sex >= 0.05, "Comparable", ifelse(auc_met > auc_sex, "Superior", "-"))
  
  # 收入
  formula_income <- as.formula(paste(outcome, "~ as.factor(income)", sep = ""))
  model_income <- glm(formula_income, data = data_train, family = binomial)
  data_test[[paste(outcome, "_riskscore_income", sep = "")]] <- predict(model_income, newdata = data_test)
  roc_result_income <- roc(data_test[[outcome]] ~ data_test[[paste(outcome, "_riskscore_income", sep = "")]], data = data_test, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_income <- roc_result_income$auc
  ci_lower_income <- roc_result_income$ci[1]
  ci_upper_income <- roc_result_income$ci[3]
  test_income <- roc.test(roc_result_met,roc_result_income) #Delong's test
  p_income <- test_income[["p.value"]]
  compare_income <- ifelse(p_income >= 0.05, "Comparable", ifelse(auc_met > auc_income, "Superior", "-"))
  
  # 吸烟
  formula_smoking <- as.formula(paste(outcome, "~ as.factor(smoking)", sep = ""))
  model_smoking <- glm(formula_smoking, data = data_train, family = binomial)
  data_test[[paste(outcome, "_riskscore_smoking", sep = "")]] <- predict(model_smoking, newdata = data_test)
  roc_result_smoking <- roc(data_test[[outcome]] ~ data_test[[paste(outcome, "_riskscore_smoking", sep = "")]], data = data_test, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_smoking <- roc_result_smoking$auc
  ci_lower_smoking <- roc_result_smoking$ci[1]
  ci_upper_smoking <- roc_result_smoking$ci[3]
  test_smoking <- roc.test(roc_result_met,roc_result_smoking) #Delong's test
  p_smoking <- test_smoking[["p.value"]]
  compare_smoking <- ifelse(p_smoking >= 0.05, "Comparable", ifelse(auc_met > auc_smoking, "Superior", "-"))
  
  # 喝酒
  formula_drinking <- as.formula(paste(outcome, "~ as.factor(drinking)", sep = ""))
  model_drinking <- glm(formula_drinking, data = data_train, family = binomial)
  data_test[[paste(outcome, "_riskscore_drinking", sep = "")]] <- predict(model_drinking, newdata = data_test)
  roc_result_drinking <- roc(data_test[[outcome]] ~ data_test[[paste(outcome, "_riskscore_drinking", sep = "")]], data = data_test, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_drinking <- roc_result_drinking$auc
  ci_lower_drinking <- roc_result_drinking$ci[1]
  ci_upper_drinking <- roc_result_drinking$ci[3]
  test_drinking <- roc.test(roc_result_met,roc_result_drinking) #Delong's test
  p_drinking <- test_drinking[["p.value"]]
  compare_drinking <- ifelse(p_drinking >= 0.05, "Comparable", ifelse(auc_met > auc_drinking, "Superior", "-"))
  
  # 教育
  formula_education <- as.formula(paste(outcome, "~ as.factor(education)", sep = ""))
  model_education <- glm(formula_education, data = data_train, family = binomial)
  data_test[[paste(outcome, "_riskscore_education", sep = "")]] <- predict(model_education, newdata = data_test)
  roc_result_education <- roc(data_test[[outcome]] ~ data_test[[paste(outcome, "_riskscore_education", sep = "")]], data = data_test, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_education <- roc_result_education$auc
  ci_lower_education <- roc_result_education$ci[1]
  ci_upper_education <- roc_result_education$ci[3]
  test_education <- roc.test(roc_result_met,roc_result_education) #Delong's test
  p_education <- test_education[["p.value"]]
  compare_education <- ifelse(p_education >= 0.05, "Comparable", ifelse(auc_met > auc_education, "Superior", "-"))
  
  # BMI
  formula_bmi <- as.formula(paste(outcome, "~ as.factor(bmi_c)", sep = ""))
  model_bmi <- glm(formula_bmi, data = data_train, family = binomial)
  data_test[[paste(outcome, "_riskscore_bmi", sep = "")]] <- predict(model_bmi, newdata = data_test)
  roc_result_bmi <- roc(data_test[[outcome]] ~ data_test[[paste(outcome, "_riskscore_bmi", sep = "")]], data = data_test, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_bmi <- roc_result_bmi$auc
  ci_lower_bmi <- roc_result_bmi$ci[1]
  ci_upper_bmi <- roc_result_bmi$ci[3]
  test_bmi <- roc.test(roc_result_met,roc_result_bmi) #Delong's test
  p_bmi <- test_bmi[["p.value"]]
  compare_bmi <- ifelse(p_bmi >= 0.05, "Comparable", ifelse(auc_met > auc_bmi, "Superior", "-"))
  
  # 降脂药
  formula_med_statin <- as.formula(paste(outcome, "~ as.factor(med_statin)", sep = ""))
  model_med_statin <- glm(formula_med_statin, data = data_train, family = binomial)
  data_test[[paste(outcome, "_riskscore_med_statin", sep = "")]] <- predict(model_med_statin, newdata = data_test)
  roc_result_med_statin <- roc(data_test[[outcome]] ~ data_test[[paste(outcome, "_riskscore_med_statin", sep = "")]], data = data_test, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_med_statin <- roc_result_med_statin$auc
  ci_lower_med_statin <- roc_result_med_statin$ci[1]
  ci_upper_med_statin <- roc_result_med_statin$ci[3]
  test_med_statin <- roc.test(roc_result_met,roc_result_med_statin) #Delong's test
  p_med_statin <- test_med_statin[["p.value"]]
  compare_med_statin <- ifelse(p_med_statin >= 0.05, "Comparable", ifelse(auc_met > auc_med_statin, "Superior", "-"))
  
  # 降压药
  formula_med_hbp <- as.formula(paste(outcome, "~ as.factor(med_hbp)", sep = ""))
  model_med_hbp <- glm(formula_med_hbp, data = data_train, family = binomial)
  data_test[[paste(outcome, "_riskscore_med_hbp", sep = "")]] <- predict(model_med_hbp, newdata = data_test)
  roc_result_med_hbp <- roc(data_test[[outcome]] ~ data_test[[paste(outcome, "_riskscore_med_hbp", sep = "")]], data = data_test, print.thres = TRUE, print.auc = TRUE, ci = TRUE, plot = FALSE, smooth = FALSE)
  auc_med_hbp <- roc_result_med_hbp$auc
  ci_lower_med_hbp <- roc_result_med_hbp$ci[1]
  ci_upper_med_hbp <- roc_result_med_hbp$ci[3]
  test_med_hbp <- roc.test(roc_result_met,roc_result_med_hbp) #Delong's test
  p_med_hbp <- test_med_hbp[["p.value"]]
  compare_med_hbp <- ifelse(p_med_hbp >= 0.05, "Comparable", ifelse(auc_met > auc_med_hbp, "Superior", "-"))
  
  
  
  # 将结果添加到 auc_results 数据框中，并改名
  auc_results <- rbind(auc_results, c(outcome, auc_met, ci_lower_met, ci_upper_met,
                                      auc_age, ci_lower_age, ci_upper_age, p_age, compare_age,
                                      auc_sex, ci_lower_sex, ci_upper_sex, p_sex, compare_sex,
                                      auc_income, ci_lower_income, ci_upper_income, p_sex, compare_income,
                                      auc_smoking, ci_lower_smoking, ci_upper_smoking, p_sex, compare_smoking,
                                      auc_drinking, ci_lower_drinking, ci_upper_drinking, p_drinking, compare_drinking,
                                      auc_education, ci_lower_education, ci_upper_education, p_education, compare_education,
                                      auc_bmi, ci_lower_bmi, ci_upper_bmi, p_bmi, compare_bmi,
                                      auc_med_statin, ci_lower_med_statin, ci_upper_med_statin, p_med_statin, compare_med_statin,
                                      auc_med_hbp, ci_lower_med_hbp, ci_upper_med_hbp, p_med_hbp, compare_med_hbp
  ))
  colnames(auc_results) <- c("outcome", "MET", "met_lower", "met_upper",
                             "AGE", "age_lower", "age_upper", "Pvalue_AGE", "MET_vs._AGE",
                             "SEX", "sex_lower", "sex_upper", "Pvalue_SEX", "MET_vs._SEX",
                             "INCOME", "income_lower", "income_upper", "Pvalue_INCOME", "MET_vs._INCOME",
                             "SMOKING", "smoking_lower", "smoking_upper", "Pvalue_SMOKING", "MET_vs._SMOKING",
                             "DRINKING", "drinking_lower", "drinking_upper", "Pvalue_DRINKING", "MET_vs._DRINKING",
                             "EDUCATION", "education_lower", "education_upper", "Pvalue_EDUCATION", "MET_vs._EDUCATION",
                             "BMI", "bmi_lower", "bmi_upper", "Pvalue_BMI", "MET_vs._BMI",
                             "ANTI_LIPID", "anti_lipid_lower", "anti_lipid_upper", "Pvalue_ANTI_LIPID", "MET_vs._ANTI_LIPID",
                             "ANTI_BP", "anti_bp_lower", "anti_bp_upper", "Pvalue_ANTI_BP", "MET_vs._ANTI_BP")
}

#转换数值变量
columns_to_convert <- c("MET", "met_lower", "met_upper",
                        "AGE", "age_lower", "age_upper", "Pvalue_AGE",
                        "SEX", "sex_lower", "sex_upper", "Pvalue_SEX",
                        "INCOME", "income_lower", "income_upper", "Pvalue_INCOME",
                        "SMOKING", "smoking_lower", "smoking_upper", "Pvalue_SMOKING",
                        "DRINKING", "drinking_lower", "drinking_upper", "Pvalue_DRINKING",
                        "EDUCATION", "education_lower", "education_upper", "Pvalue_EDUCATION",
                        "BMI", "bmi_lower", "bmi_upper", "Pvalue_BMI",
                        "ANTI_LIPID", "anti_lipid_lower", "anti_lipid_upper", "Pvalue_ANTI_LIPID",
                        "ANTI_BP", "anti_bp_lower", "anti_bp_upper", "Pvalue_ANTI_BP")
auc_results[columns_to_convert] <- lapply(auc_results[columns_to_convert], as.numeric)

write.xlsx(auc_results, "xxx.xlsx")



models <- create_models()$models

# * DCA ####
# 清除除了data_all以外的所有数据
objects_to_keep <- c("data_all","create_models")  #要保留的对象名称
objects_to_remove <- setdiff(ls(), objects_to_keep)
rm(list = objects_to_remove)

# 开始
library(rmda)
models <- create_models()$models
outcomes <- c("cvd","dr_progression","vtdr","ckd")

# 存储DCA计算结果
results_baseline <- list()
results_full <- list()
for (outcome in outcomes) {
  
  if (outcome == "cvd" | outcome == "dr_progression" | outcome == "vtdr" | outcome == "ckd") {
    model_names <- c("Age_Sex","FGCRS","SCORE2","AHA_ASCVD","WHO_CVD")
  }
  if (outcome == "ckd") {
    model_names <- c("Age_Sex","Nelson_model","Chien_model","O_Seaghdha_model")
  }
  
  for (model_name in model_names) {
    
    #baseline
    formula <- formula(paste(outcome, "~", models[[model_name]], sep = ""))
    result <- decision_curve(formula, data = data_all, thresholds = seq(0, 1, by = .01))
    results_baseline[[outcome]][[model_name]] <- result
    
    #full
    formula <- formula(paste(outcome, "~", models[[model_name]], "+ score_", outcome, sep = ""))
    result <- decision_curve(formula, data = data_all, thresholds = seq(0, 1, by = .01))
    results_full[[outcome]][[model_name]] <- result
  }
  gc()
}

library(rmda)
pdf("xxx.pdf", width = 20/2, height = 25/2)
par(mfrow = c(2,2))  # 将图分成两行两列
plot_decision_curve(list(results_baseline$cvd$Age_Sex, results_full$cvd$Age_Sex, results_baseline$cvd$FGCRS, results_full$cvd$FGCRS, results_baseline$cvd$SCORE2, results_full$cvd$SCORE2, results_baseline$cvd$AHA_ASCVD, results_full$cvd$AHA_ASCVD, results_baseline$cvd$WHO_CVD, results_full$cvd$WHO_CVD),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "FGCRS", "FGCRS + PMW", "SCORE2", "SCORE2 + PMW", "AHA/ASCVD", "AHA/ASCVD + PMW", "WHO-CVD", "WHO-CVD + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.2),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("Cardiovascular disease")

plot_decision_curve(list(results_baseline$dr_progression$Age_Sex, results_full$dr_progression$Age_Sex, results_baseline$dr_progression$FGCRS, results_full$dr_progression$FGCRS, results_baseline$dr_progression$SCORE2, results_full$dr_progression$SCORE2, results_baseline$dr_progression$AHA_ASCVD, results_full$dr_progression$AHA_ASCVD, results_baseline$dr_progression$WHO_CVD, results_full$dr_progression$WHO_CVD),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "FGCRS", "FGCRS + PMW", "SCORE2", "SCORE2 + PMW", "AHA/ASCVD", "AHA/ASCVD + PMW", "WHO-CVD", "WHO-CVD + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.45),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("Diabetic retinopathy progression")

plot_decision_curve(list(results_baseline$vtdr$Age_Sex, results_full$vtdr$Age_Sex, results_baseline$vtdr$FGCRS, results_full$vtdr$FGCRS, results_baseline$vtdr$SCORE2, results_full$vtdr$SCORE2, results_baseline$vtdr$AHA_ASCVD, results_full$vtdr$AHA_ASCVD, results_baseline$vtdr$WHO_CVD, results_full$vtdr$WHO_CVD),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "FGCRS", "FGCRS + PMW", "SCORE2", "SCORE2 + PMW", "AHA/ASCVD", "AHA/ASCVD + PMW", "WHO-CVD", "WHO-CVD + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#9a144a","#9a144a","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.1),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("Vision-threatening diabetic retinopathy")

plot_decision_curve(list(results_baseline$ckd$Age_Sex, results_full$ckd$Age_Sex, results_baseline$ckd$Nelson_model, results_full$ckd$Nelson_model, results_baseline$ckd$Chien_model, results_full$ckd$Chien_model, results_baseline$ckd$O_Seaghdha_model, results_full$ckd$O_Seaghdha_model),
                    curve.names = c("Age&Sex", "Age&Sex + PMW", "Nelson model", "Nelson model + PMW", "Chien model", "Chien model + PMW", "O'Seaghdha model", "O'Seaghdha model + PMW"),
                    col = c("#5c58a6","#5c58a6","#3589bc","#3589bc","#FDDC89","#FDDC89","#d64758","#d64758","#aea4a3","#aea4a3"),
                    lty = c(4,1,4,1,4,1,4,1),
                    lwd = c(1,2,1,2,1,2,1,2,1,1),
                    xlim = c(0,0.3),
                    legend.position = "topright", cost.benefit.axis = F, confidence.intervals = F, xlab = "Risk threshold", ylab = "Standardized net benifit")
title("Chronic kidney disease")

dev.off()