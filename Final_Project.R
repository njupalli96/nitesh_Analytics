##Set Seed
set.seed(100)

##Read the Logistic Regression Data
lr_data = read.csv(file.choose())

##Modifying variables to remove skewness
lr_data$Number.of.Orders2 = lr_data$Number.of.Orders^2
lr_data$lnrevenue = log(lr_data$Revenue + 1)

View(lr_data)

lr_data$Gender = as.factor(lr_data$Gender)
lr_data$Customer.Type = as.factor(lr_data$Customer.Type)
lr_data$Satisfaction = as.factor(lr_data$Satisfaction)

str(lr_data)

##Run Logistics Regression using GLM
lr_model = glm(formula = Satisfaction ~ lnrevenue+Number.of.Orders2+Number.of.Orders+Gender+Customer.Type+Age+Income, data = lr_data, family = "binomial")
summary(lr_model)

##Psuedo R-square calculation
null_model = glm(formula = Satisfaction ~ 1, data = lr_data, family = "binomial")
1 - logLik(lr_model)/logLik(null_model)

##Odds Ratio
exp(lr_model$coefficients)

##Predicted Probability
lr_data$predict = predict(lr_model,lr_data,type = "response")

##Export the results
write.csv(lr_data,file = file.choose(),row.names = FALSE)
