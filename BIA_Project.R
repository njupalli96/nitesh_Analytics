cald = read.csv("D://IIMS//Term 5//BIA//Term Project (Group)-20221114//calibration.csv")
cs_test = read.csv("D://IIMS//Term 5//BIA//Term Project (Group)-20221114//currrent_score.csv")
fs_test = read.csv("D://IIMS//Term 5//BIA//Term Project (Group)-20221114//future_score.csv")
col=!(names(cald) %in% c("callwait_Mean","callwait_Range","ccrndmou_Mean","ccrndmou_Range","complete_Mean","complete_Range","da_Mean
","da_Range","inonemin_Mean","inonemin_Range","retdays","rmcalls","rmmou","rmrev","unan_dat_Mean","unan_dat_Range","unan_vce_Mean
","unan_vce_Range","cartype","children","crtcount","div_type","educ1","last_swap","mailflag","mailordr","mailresp","numbcars","occu1","pcowner","pre_hnd_price","proptype","REF_QTY","solflag","tot_acpt","tot_ret","wrkwoman","Customer_ID","crclscod","csa","ownrent","lor","dwlltype","dwllsize","infobase","HHstatin","prizm_social_one","area","ethnic","dualband","refurb_new","hnd_webcap"))

colc=!(names(cs_test) %in% c("callwait_Mean","callwait_Range","ccrndmou_Mean","ccrndmou_Range","complete_Mean","complete_Range","da_Mean
","da_Range","inonemin_Mean","inonemin_Range","retdays","rmcalls","rmmou","rmrev","unan_dat_Mean","unan_dat_Range","unan_vce_Mean
","unan_vce_Range","cartype","children","crtcount","div_type","educ1","last_swap","mailflag","mailordr","mailresp","numbcars","occu1","pcowner","pre_hnd_price","proptype","REF_QTY","solflag","tot_acpt","tot_ret","wrkwoman","Customer_ID","crclscod","csa","ownrent","lor","dwlltype","dwllsize","infobase","HHstatin","prizm_social_one","area","ethnic","dualband","refurb_new","hnd_webcap"))

colf=!(names(fs_test) %in% c("callwait_Mean","callwait_Range","ccrndmou_Mean","ccrndmou_Range","complete_Mean","complete_Range","da_Mean
","da_Range","inonemin_Mean","inonemin_Range","retdays","rmcalls","rmmou","rmrev","unan_dat_Mean","unan_dat_Range","unan_vce_Mean
","unan_vce_Range","cartype","children","crtcount","div_type","educ1","last_swap","mailflag","mailordr","mailresp","numbcars","occu1","pcowner","pre_hnd_price","proptype","REF_QTY","solflag","tot_acpt","tot_ret","wrkwoman","Customer_ID","crclscod","csa","ownrent","lor","dwlltype","dwllsize","infobase","HHstatin","prizm_social_one","area","ethnic","dualband","refurb_new","hnd_webcap"))

caldn = cald[col]
cs_testn = cs_test[colc]
fs_testn = fs_test[colf]

facs = (names(caldn) %in% c("adults","asl_flag","car_buy","churn","creditcd"
                            ,"forgntvl","income","kid0_2","kid3_5","kid6_10","kid11_15","kid16_17","marital","mtrcycle"
                            ,"new_cell","rv","truck"))

facsc = (names(cs_testn) %in% c("adults","asl_flag","car_buy","creditcd"
                                ,"forgntvl","income","kid0_2","kid3_5","kid6_10","kid11_15","kid16_17","marital","mtrcycle"
                                ,"new_cell","rv","truck"))

facsf = (names(fs_testn) %in% c("adults","asl_flag","car_buy","creditcd"
                                ,"forgntvl","income","kid0_2","kid3_5","kid6_10","kid11_15","kid16_17","marital","mtrcycle"
                                ,"new_cell","rv","truck"))

caldn[facs] = lapply(caldn[facs],as.factor)
caldn[!facs] = lapply(caldn[!facs], as.numeric)

caldn = caldn[!is.na(caldn$phones),]
caldn = caldn[!is.na(caldn$models),]
caldn = caldn[!is.na(caldn$eqpdays),]

cs_testn[facsc] = lapply(cs_testn[facsc],as.factor)
cs_testn[!facsc] = lapply(cs_testn[!facsc], as.numeric)

fs_testn[facsf] = lapply(fs_testn[facsf],as.factor)
fs_testn[!facsf] = lapply(fs_testn[!facsf], as.numeric)

sapply(caldn, function(x) sum(is.na(x)))
sapply(cs_testn, function(x) sum(is.na(x)))
sapply(fs_testn, function(x) sum(is.na(x)))

library("mice")

caldn_multimp <- mice(caldn,meth='cart')
caldn_fin<-mice::complete(caldn_multimp,1)

cs_testn_multimp <- mice(cs_testn,meth='cart')
cs_testn_fin<-mice::complete(cs_testn_multimp,1)

fs_testn_multimp <- mice(fs_testn,meth='cart')
fs_testn_fin<-mice::complete(fs_testn_multimp,1)

library(writexl)
write_xlsx(caldn_fin,"D://IIMS//Term 5//BIA//Term Project (Group)-20221114//caldn_fin.xlsx")
write_xlsx(cs_testn_fin,"D://IIMS//Term 5//BIA//Term Project (Group)-20221114//cs_testn_fin.xlsx")
write_xlsx(fs_testn_fin,"D://IIMS//Term 5//BIA//Term Project (Group)-20221114//fs_testn_fin.xlsx")
library(readxl)
caldn_fin = read_excel("D://IIMS//Term 5//BIA//Term Project (Group)-20221114//caldn_fin.xlsx")
cs_testn_fin = read_excel("C:\Users\HP\Downloads\a.xlsx")
fs_testn_fin = read_excel("C:\Users\HP\Downloads\fs_testn_fin.xlsx")

facs = (names(caldn_fin) %in% c("adults","asl_flag","car_buy","churn","creditcd"
                            ,"forgntvl","income","kid0_2","kid3_5","kid6_10","kid11_15","kid16_17","marital","mtrcycle"
                            ,"new_cell","rv","truck"))

facsc = (names(cs_testn_fin) %in% c("adults","asl_flag","car_buy","creditcd"
                                ,"forgntvl","income","kid0_2","kid3_5","kid6_10","kid11_15","kid16_17","marital","mtrcycle"
                                ,"new_cell","rv","truck"))

facsf = (names(fs_testn_fin) %in% c("adults","asl_flag","car_buy","creditcd"
                                ,"forgntvl","income","kid0_2","kid3_5","kid6_10","kid11_15","kid16_17","marital","mtrcycle"
                                ,"new_cell","rv","truck"))

caldn_fin[facs] = lapply(caldn_fin[facs],as.factor)
caldn_fin[!facs] = lapply(caldn_fin[!facs], as.numeric)

cs_testn_fin[facsc] = lapply(cs_testn_fin[facsc],as.factor)
cs_testn_fin[!facsc] = lapply(cs_testn_fin[!facsc], as.numeric)

fs_testn_fin[facsf] = lapply(fs_testn_fin[facsf],as.factor)
fs_testn_fin[!facsf] = lapply(fs_testn_fin[!facsf], as.numeric)

library(arsenal)
summary(comparedf(caldn_fin,fs_testn_fin))
summary(comparedf(caldn_fin,cs_testn_fin))

library(caret)
samplesize= 0.80 * nrow(caldn_fin)
set.seed(80)
index = sample(seq_len(nrow(caldn_fin)),size = samplesize)

train = caldn_fin[index,]
test = caldn_fin[-index,]
summary.factor(train$churn)

#Random Forest
library(randomForest)
set.seed(120)  # Setting seed

mtry <- tuneRF(train[-83],train$churn, ntreeTry=500,stepFactor=1.5,improve=0.01, trace=TRUE, plot=TRUE)

best.m <- mtry[mtry[, 2] == min(mtry[, 2]), 1]
print(best.m)

classifier_RF_2 = randomForest(churn~.,train,na.action = na.omit,mtry = 16,importance = TRUE)
y_pred = predict(classifier_RF_2, newdata = test)
confusionMatrix(y_pred,test$churn, positive = "1")

#Logistic Regression

logr = glm(churn~.,data = train, family = "binomial",na.action = na.omit)
pred.log = predict(logr,test)
predict_reg <- ifelse(pred.log >0.5, 1, 0)
predict_reg = as.factor(predict_reg)
confusionMatrix(predict_reg,test$churn, positive = "1")

#Naive
library(e1071)
set.seed(120)  # Setting Seed
classifier_NB = naiveBayes(churn ~ ., data = train)
y_pred_NB <- predict(classifier_NB, newdata = test)

# Confusion Matrix
cm <- table(test$churn, y_pred_NB)
cm

# Model Evaluation
confusionMatrix(cm,positive = "1")

library("lift")
TopDecileLift(y_pred_NB,test$churn)
plotLift(y_pred_NB,test$churn,cumulative = TRUE)

# Plotting model
plot(classifier_NB)

# Importance plot
importance(classifier_RF_2)

# Variable importance plot
varImpPlot(classifier_RF_2)

# Create ROC Curve
library(pROC)
roc_score = roc(as.numeric(test$churn), as.numeric(y_pred_NB))  # AUC score
plot(roc_score)

#Gini_Coefficient
library(ineq)
ineq(y_pred_NB, type = "Gini")

# Predicting the Test set results
y_pred_cs = predict(classifier_RF_2, newdata = cs_testn_fin)
y_pred_fs = predict(classifier_RF_2, newdata = fs_testn_fin)

y_pred_csD = data.frame(y_pred_cs)
y_pred_fsD = data.frame(y_pred_fs)

write_xlsx(y_pred_csD,"D://IIMS//Term 5//BIA//Term Project (Group)-20221114//y_pred_cs.xlsx")
write_xlsx(y_pred_fsD,"D://IIMS//Term 5//BIA//Term Project (Group)-20221114//y_pred_fs.xlsx")
