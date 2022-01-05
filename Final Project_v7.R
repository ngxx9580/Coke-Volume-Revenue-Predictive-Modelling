library("tidyverse")
library("sqldf")
library("readxl")
library("car")
library("estimatr")
library("caret")
library("lubridate")
library("mice")
library("writexl")

cor(d_main_data)
# load file
a_raw <- read_excel(file.choose(),sheet=3)

# region features
a_raw$region_quebec_ind <- ifelse(a_raw$REGION == 'QUEBEC',1,0)
a_raw$region_west_ind <- ifelse(a_raw$REGION == 'WEST',1,0)
a_raw$region_east_ind <- ifelse(a_raw$REGION == 'EAST',1,0)

# channel features
a_raw$discount_channel_ind <- ifelse(a_raw$CHANNEL == 'DISCOUNT',1,0)
a_raw$wholesale_channel_ind <-ifelse(a_raw$CHANNEL == 'WHOLESALE',1,0)

# category features
a_raw$frozen_ind <-ifelse(a_raw$CATEGORY == 'JUICE - FROZEN',1,0)
a_raw$shelf_stable_ind <-ifelse(a_raw$CATEGORY == 'JUICE - SHELF STABLE',1,0)

# brand features
a_raw$minute_maid_brand_ind <- ifelse(a_raw$BRAND == 'MINUTE MAID',1,0)

# week_name features
a_raw$holiday_ind <- ifelse(str_starts(a_raw$'WEEK_NAME', 'Week'),0,1)

# flyer features
#a_raw$flyer_ind <- ifelse(a_raw$FEATURE == '(blank)',0,1)
#a_raw$front_page_ind <- ifelse(a_raw$FEATURE =='FRONT PAGE',1,0)
#a_raw$back_page_ind <- ifelse(a_raw$FEATURE =='BACK PAGE',1,0)


# aggregate to product_group level to capture seasonality effect
b_seaonality <- a_raw %>%
  group_by(CATEGORY, BRAND,YEAR,WEEK) %>%
  summarise(MEAN_UNITS = mean(UNITS),
            SD_UNITS = sd(UNITS)) 

# join dataframe
c_main_data <- merge(b_seaonality,a_raw, by.x=c('CATEGORY','BRAND','YEAR','WEEK'),
                     by.y=c('CATEGORY','BRAND','YEAR','WEEK'))


# Seasonality
c_main_data$seasonality <- round(((c_main_data$UNITS - c_main_data$MEAN_UNITS)/c_main_data$SD_UNITS),digits = 4)
c_main_data$seasonality [is.na(c_main_data$seasonality)] <- 0

# log transformation on units due to right-skewed density plot based on previous run
c_main_data$logunits <- log(c_main_data$UNITS)
c_main_data$logprice <- log(c_main_data$PRICE)


# Final data frame
d_main_data <- select(c_main_data,-CATEGORY,-BRAND,-MEAN_UNITS,-SD_UNITS,-REGION,-CHANNEL,-WEEK_NAME,
                      -UNITS,-REVENUE,-PRICE)

# test-train split
e_train_data <- filter(d_main_data, YEAR!=2019) #2017&2018
f_test_data <- filter(d_main_data, YEAR==2019)

# train linear model on logunits
g_train_reg <- lm(logunits~.,e_train_data)
summary(g_train_reg) # Adjusted R-squared:  0.7058, standard error: 1.382

plot(g_train_reg)
plot(density(resid(g_train_reg)))

# test
h_pred_test <- predict(g_train_reg, f_test_data)

# model evaluation metrics
data.frame( R2 = R2(h_pred_test, f_test_data$logunits),
            RMSE = RMSE(h_pred_test, f_test_data$logunits),
            MAE = MAE(h_pred_test, f_test_data$logunits))


# MAPE
f_test_data$pred <- predict(g_train_reg, f_test_data)
f_test_data$UNITS <- round((exp(f_test_data$logunits)),digits = 0)
f_test_data$UNITS2 <- round((exp(f_test_data$pred)),digits = 0)
#mean(abs((f_test_data$UNITS-f_test_data$UNITS2)/f_test_data$UNITS))*100
(mean(abs((f_test_data$UNITS-f_test_data$UNITS2)/f_test_data$UNITS))*100)/(nrow(f_test_data))

confint(g_train_reg)

#write_xlsx(f_test_data,"C:\\Users\\Leon\\Desktop\\pred_test.xlsx")



# 2022 promotion file setup
h_raw <- read_excel(file.choose(),sheet=4)

# region features
h_raw$region_quebec_ind <- ifelse(h_raw$REGION == 'QUEBEC',1,0)
h_raw$region_west_ind <- ifelse(h_raw$REGION == 'WEST',1,0)
h_raw$region_east_ind <- ifelse(h_raw$REGION == 'EAST',1,0)

# channel features
h_raw$discount_channel_ind <- ifelse(h_raw$CHANNEL == 'DISCOUNT',1,0)
h_raw$wholesale_channel_ind <-ifelse(h_raw$CHANNEL == 'WHOLESALE',1,0)

# category features
h_raw$frozen_ind <-ifelse(h_raw$CATEGORY == 'JUICE - FROZEN',1,0)
h_raw$shelf_stable_ind <-ifelse(h_raw$CATEGORY == 'JUICE - SHELF STABLE',1,0)

# brand features
h_raw$minute_maid_brand_ind <- ifelse(h_raw$BRAND == 'MINUTE MAID',1,0)

# week_name features
h_raw$holiday_ind <- ifelse(str_starts(h_raw$'WEEK_NAME', 'Week'),0,1)

# flyer features
#h_raw$flyer_ind <- ifelse(h_raw$FEATURE == '(blank)',0,1)
#h_raw$front_page_ind <- ifelse(h_raw$FEATURE =='FRONT PAGE',1,0)
#h_raw$back_page_ind <- ifelse(h_raw$FEATURE =='BACK PAGE',1,0)

# log transformation on units due to right-skewed density plot based on previous run
h_raw$logprice <- log(h_raw$PRICE)


# Historical seasonality
i_hist_seasonality <- c_main_data %>%
  select(CATEGORY,BRAND,WEEK,seasonality) %>% 
  group_by(CATEGORY, BRAND,WEEK) %>% 
  summarise(seasonality= mean(seasonality))


# final df for prediction regression
j_final_data <- merge(i_hist_seasonality,h_raw, by.x=c('CATEGORY', 'BRAND','WEEK'),
                      by.y=c('CATEGORY', 'BRAND','WEEK'))



# Final data frame
k_final_data <- select(j_final_data,-CATEGORY,-BRAND,-REGION,-CHANNEL,-WEEK_NAME,-PRICE)
                       

# final regression
l_final_reg <- lm(logunits~.,d_main_data)
summary(l_final_reg)

# Prediction
h_raw$pred <- predict(l_final_reg,k_final_data)
h_raw$UNITS <- round((exp(h_raw$pred)),digits = 0)

m_result <- select(h_raw,REGION, CHANNEL, CATEGORY, BRAND, WEEK_NAME, YEAR, WEEK, PRICE, UNITS)

#export result
write_xlsx(m_result,"C:\\Users\\isaac\\Desktop\\2022_prediction.xlsx")

