## THIS SCRIPT ESTIMATES A DiD MODEL FOR EU COMPANIES THAT RECEIVED A GRANT

rm(list = ls())
library(dplyr)
library(readr)
library(ggplot2)
library(tidyr)
library(sjlabelled)
library(ggrepel)
library(scales)
library(ggpubr)
library(plm)
library(readxl)
library(lmtest)
library(stargazer)


# Set the current working directory as the folder that contains this R file
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

## 1) IMPORT DATA =========================================================================

# NEW Method: Successful companies - This is the new way to import successful companies thorugh
# the file "Successful_Applicants.xlsx"
sheets = c("DE_Marg", "IT_Marg", "FR_Marg", "ES_Marg")
successfull = list()
for (i in 1:length(sheets)){
  # Read the excel sheets of the file with the companies successfull
  successfull[[i]] <- read_excel("Successful_Applicants.xlsx", sheet = sheets[i])
  name_cols <- c("name", "margin_post", "margin_pre")
  # Change the name of the columns
  colnames(successfull[[i]]) <- name_cols
  # If the column must be numeric, trnasofrm it as numeric (sometimes it is saved as character)
  for (l in 1:length(name_cols)){
    if (name_cols[l] != "name"){
      successfull[[i]][[l]] = as.numeric(successfull[[i]][[l]]) 
    }
  }
  # Add column of nationality
  successfull[[i]][["Nationality"]] <- substr(sheets[i],1,2)
}

succ_df = bind_rows(successfull)
succ_df[["success"]] = 1



# Control companies
sheets = c("Comp_DE", "Comp_IT", "Comp_ES", "Comp_FR")
control = list()
for (i in 1:length(sheets)){
  # Read the excel sheets of the file with the companies successfull
  control[[i]] <- read_excel("Control_Group.xlsx", sheet = sheets[i])
  name_cols <- c("name", "prod_post", "prod_pre","margin_post", "margin_pre")
  # Change the name of the columns
  colnames(control[[i]]) <- name_cols
  # If the column must be numeric, trnasofrm it as numeric (sometimes it is saved as character)
  for (l in 1:length(name_cols)){
    if (name_cols[l] != "name"){
      control[[i]][[l]] = as.numeric(control[[i]][[l]]) 
    }
  }
  # Add column of nationality
  control[[i]][["Nationality"]] <- substr(sheets[i],6,7)
}

control_df = bind_rows(control)
control_df[["success"]] = 0

## 2) MERGE TREATMENT AND CONTROL ====================================================
## See the columns with same names
columns_final = intersect(colnames(control_df), colnames(succ_df))

# Keep the columns which appear in both
df_final = rbind(succ_df[columns_final],
                 control_df[columns_final])

# From wide to long dataframe
df_final = df_final %>% pivot_longer(c("margin_pre", "margin_post"), names_to = "margin_pre_post", values_to = "margin") %>%
  # Createa new dummy variable which is equal to 1 if the value of the margin is post treatment
  mutate(margin_dummy = as.numeric(margin_pre_post == "margin_post")) %>%
  # Remove the column 'margin_pre_post' which is not more useful
  mutate(margin_pre_post = NULL)

## rename dummies
df_final <- df_final %>% 
  rename(pre_post = margin_dummy, 
         treated = success)

write.csv(df_final, "df_final.csv")

## 3) ESTIMATE THE DID MODEL =====================================================
##estimate did via standard OLS
did_model <- lm(margin ~ pre_post + treated + pre_post:treated, data = df_final)
summary(did_model)
#did_model2 <- lm(margin ~ pre_post + treated + pre_post:treated + Nationality, data = df_final)
#summary(did_model2)

#check robustness to outliers
length(df_final$margin)
quartiles <- quantile(df_final$margin, probs=c(.25, .75), na.rm = T)
IQR <- IQR(df_final$margin, na.rm = T)
Lower <- quartiles[1] - 1.5*IQR
Upper <- quartiles[2] + 1.5*IQR 
data_no_outlier <- subset(df_final, margin > Lower & margin < Upper)
nrow(data_no_outlier)

did_model3 <- lm(margin ~ pre_post + treated + pre_post:treated, data = data_no_outlier)
summary(did_model3)
#did_model4 <- lm(margin ~ pre_post + treated + pre_post:treated + Nationality, data = data_no_outlier)
#summary(did_model4)

# Save to Latex the output without outliers
stargazer(did_model3, type="latex", # type = "text" shows a preview
          title="Table X: Average profit margin before and after the EU grant", 
          align=TRUE)
          #dep.var.labels = "Profit Margin", 
          #covariate.labels=c("Change in profit margin"))


## This is what interests you ---------------------------------------
# Estimate via fixed effects (see https://en.wikipedia.org/wiki/Fixed_effects_model)
# the main idea is that if there are omitted variables that are invariant across time, to avoid omitted variable bias, 
# we can demean the data and run the usual regression. Since the omitted variables are invariant, 
# when the data is demeaned they disappear. 
panel <- pdata.frame(df_final, "name") # declare as panel data
did.reg <- plm(margin ~ pre_post + treated + pre_post:treated, data = panel, model = "within") # within model
coeftest(did.reg, vcov = function(x) # obtain clustered standard errors
  vcovHC(x, cluster = "group", type = "HC1"))

# This is what really interests you --------------------------
# Estimate via fixed effects without outliers
panel_no_out <- pdata.frame(data_no_outlier, "name") # declare as panel data
did.reg2 <- plm(margin ~ pre_post + treated + pre_post:treated, data = panel_no_out, model = "within") # within model
coeftest(did.reg2, vcov = function(x) # obtain clustered standard errors
  vcovHC(x, cluster = "group", type = "HC1"))
stargazer(did.reg2, type="text", # type = "text" shows a preview, while type = "latex" shows the code to insert in Latex 
          title="Table X: Regression results estimated via fixed effetcts without outliers", 
          align=TRUE)

stargazer(did.reg2, type="latex", # type = "text" shows a preview, while type = "latex" shows the code to insert in Latex 
          title="Table X: Regression results estimated via fixed effetcts without outliers", 
          align=TRUE)


# Notice that the coefficient linked to pre_post:treated is positive and significant

#dep.var.labels = "Profit Margin", 
#covariate.labels=c("Change in profit margin"))

##############################################################################

## 3.5) Charts with the lines ===============================================================

# Treated data: observed effect
treat = data_no_outlier %>%
  filter(treated == 1)

treat_pre = treat %>% 
  filter(pre_post == 0)

treat_post = treat %>% 
  filter(pre_post == 1)

treat_avg = c(mean(treat_pre$margin),
              mean(treat_post$margin))

treat_obs = data.frame("Time" = c(1,2),
           "Margin Average" = treat_avg)

# Control data: observed changes
control = data_no_outlier %>%
  filter(treated == 0)

control_pre = control %>% 
  filter(pre_post == 0)

control_post = control %>% 
  filter(pre_post == 1)

control_avg = c(mean(control_pre$margin),
              mean(control_post$margin))

control_obs = data.frame("Time" = c(1,2),
                       "Margin Average" = control_avg)

# Treated: with the trend of the control

treat_trend = data.frame("Time" = c(1,2),
                         "Margin Average" = c(treat_avg[1], treat_avg[1] + diff(control_avg))) 

# Plot with the three lines

plot(treat_obs, 
     type = "l",
     xlim = c(0.5, 2.5),
     ylim = c(3, 7),
     col = "orange",
     lwd = 2,
     xlab = "",
     xaxt = "n",
     ylab = "Margin Average")

axis(1, at = seq(1, 2, by = 1), c("Time 1", "Time 2"))

grid()

lines(control_obs,
      col = "blue",
      lwd = 2)

lines(treat_trend,
      col = "orange",
      lwd = 1, 
      lty = "dashed")

leg.txt = c("Actual Trend of Treatment Group", 
            "Trend of Control Group",
            "Expected Trend of Treatment Group")
leg.col = c("orange", 
            "blue",
            "orange")

leg.lty = c("solid",
            "solid",
            "dashed")

leg.lwd = c(2,
            2,
            1)

legend("bottom", 
       leg.txt,
       lty = leg.lty,
       lwd = leg.lwd,
       col = leg.col,
       cex = 0.65
       )

## 4) OPTIONAL. ================================================================
# This is to merge the information from organization.xksx file with the information on the companies winning the grant
# Not necessary to run the diff in diff.
orgaz <- read_excel("organization.xlsx")
df_orgaz <- orgaz %>% 
  dplyr::select(name, netEcContribution)
df_orgaz$netEcContribution = as.numeric(df_orgaz$netEcContribution) 
df_orgaz = df_orgaz %>%
  group_by(name) %>%
  summarise(grant_tot = sum(netEcContribution),
            projects_num = length(netEcContribution))%>%
  ungroup()

# Merge organization with successfull
succ_org = merge(succ_df, df_orgaz, by.x = c("name"), by.y = c("name"))

## 5) OLD Method NOT TO RUN ---------------------------------------------------
# OLD Method: Successful companies - This is an old way to import the file "ORBIS - Successful.xlsx"
sheets = c("DE Final", "IT Final", "FR Final", "ES Final")
successfull = list()
for (i in 1:length(sheets)){
  # Read the excel sheets of the file with the companies successfull
  successfull[[i]] <- read_excel("ORBIS - Successful.xlsx", sheet = sheets[i])
  name_cols <- c("name", "prod_post", "prod_pre", "Change_Prod", "utile_post", "utile_pre", "Change_Utile",
                 "Empl-2", "Empl-7", "Change_Empl")
  # Change the name of the columns
  colnames(successfull[[i]]) <- name_cols
  # If the column must be numeric, trnasofrm it as numeric (sometimes it is saved as character)
  for (l in 1:length(name_cols)){
    if (name_cols[l] != "name"){
      successfull[[i]][[l]] = as.numeric(successfull[[i]][[l]]) 
    }
  }
  # Add column of nationality
  successfull[[i]][["Nationality"]] <- substr(sheets[i],1,2)
  
  # Add column of profit margin
  successfull[[i]][["margin_pre"]] <- successfull[[i]][["utile_pre"]]/successfull[[i]][["prod_pre"]]*100
  successfull[[i]][["margin_post"]] <- successfull[[i]][["utile_post"]]/successfull[[i]][["prod_post"]]*100
}

succ_df = bind_rows(successfull)
succ_df[["success"]] = 1


