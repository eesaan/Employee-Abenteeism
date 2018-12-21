            #################              Employee Absenteeism Analysis            ################


# Removing all the prior stored objects:
rm(list=ls())

# Setting the working directory:
setwd("W:/DataScience/edWisor/Project")

# Importing Libraries:

library(xlsx)
library(dplyr)
library(ggplot2)
library(DMwR)
library(gridExtra)
library(ggcorrplot)
library(caTools)
library(dummies)
library(devtools)
library(Metrics)
library(caret)
library(rpart)
library(randomForest)
library(fastDummies)


# Loading the data:
df = read.xlsx("Absenteeism_at_work_Project.xls", sheetIndex = 1, header = TRUE)


# Checking the dataframe head:
head(df)


                             ################### DATA PRE PROCESSING ####################

# Checking the strucutre & datatype information of data:
str(df)

# Replacing the . in colnames with _ :

colnames(df) = gsub(".", "_", colnames(df), fixed = TRUE)
colnames(df)


# Vector of numerical variables:
num_var = c("Transportation_expense", "Distance_from_Residence_to_Work", "Service_time", "Age", "Work_load_Average_day_","Hit_target","Son","Pet","Weight","Height","Body_mass_index","Absenteeism_time_in_hours")

# Vector of categorical variables:
cat_var = c("ID","Reason_for_absence", "Month_of_absence", "Day_of_the_week", "Seasons", "Disciplinary_failure", "Education", "Social_drinker", "Social_smoker")



#### Missing Value Analysis ####

# Creating dataframe of missing values for every variable in decreasing order:

mis_val = data.frame(apply(df,2,function(x){sum(is.na(x))}))
colnames(mis_val) = "mis_percent"
mis_val$mis_percent = ( mis_val$mis_percent/nrow(df) )*100
mis_val$variables = row.names(mis_val)
row.names(mis_val) = NULL
mis_val = mis_val[order(-mis_val$mis_percent),]
mis_val = mis_val[,c(2,1)]
View(mis_val)

# Droping observation in which "Absenteeism time in hours" has missing value:

df = df[is.na(df$Absenteeism_time_in_hours)==0,]


#In "Reason of absence", the encoded category 20 is not there and 0 is there, so encoded value of 0 should be replaced with 20.

df$Reason_for_absence[df$Reason_for_absence == 0] = 20

# Converting in proper datatype:
df =  df %>% mutate_at(.vars = cat_var,.funs = as.factor)
str(df)


# KNN Imputation
df = knnImputation(data = df, k = 5)
sum(is.na(df))


#### Outlier Analysis ####

# Checking for outliers in the numerical variables & Using the KNN Imputation method for outlier analysis.

# Boxplot of numerical variables:

for (i in 1:length(num_var))
{
  assign(paste0("gn",i), ggplot(aes_string(y = (num_var[i])), data = subset(df))+ 
           stat_boxplot(geom = "errorbar", width = 0.5) +
           geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,
                        outlier.size=3, notch=FALSE) +
           theme(legend.position="bottom")+
           labs(y=num_var[i])+
           ggtitle(num_var[i]))
}

# Plotting plots together
gridExtra::grid.arrange(gn1,gn2, gn3,gn4,nrow=1)
gridExtra::grid.arrange(gn5,gn6,gn7,gn8,nrow=1)
gridExtra::grid.arrange(gn9,gn10,gn11,gn12,nrow=1)



# Replace all outliers with NA and impute
for(i in num_var)
{
  val = df[,i][df[,i] %in% boxplot.stats(df[,i])$out]
  
  df[,i][df[,i] %in% val] = NA
}

# Imputing missing values
df = knnImputation(df,k=5)

####  Visualisations for exploring relationships b/w the variables: ####

# Barplot of Mean Absent Time v/s Reason:

  ggplot(data=df, aes(x=Reason_for_absence, y=Absenteeism_time_in_hours, fill=Reason_for_absence)) + 
  stat_summary(fun.y = mean, geom = "bar") + ggtitle("Mean Absent Time vs Reason") +
  labs(y="Mean Absent Time")

# Barplot of Mean Absent Time v/s Month of Absence:

  ggplot(data=df, aes(x=Month_of_absence, y=Absenteeism_time_in_hours, fill=Month_of_absence)) + 
  stat_summary(fun.y = mean, geom = "bar") + ggtitle("Mean Absent Time vs Month Absent") +
  labs(y="Mean Absent Time")

# Barplot of Mean Absent Time v/s Day of week:  
  
  ggplot(data=df, aes(x=Day_of_the_week, y=Absenteeism_time_in_hours, fill=Day_of_the_week)) + 
  stat_summary(fun.y = mean, geom = "bar") + ggtitle("Mean Absent Time vs Day of week") +
  labs(y="Mean Absent Time")
  
# Barplot of Mean Absent Time v/s Disciplinary Failure:
    
  ggplot(data=df, aes(x=Disciplinary_failure, y=Absenteeism_time_in_hours, fill=Disciplinary_failure)) + 
  stat_summary(fun.y = mean, geom = "bar") + ggtitle("Mean Absent Time vs Disciplinary Failure") +
  labs(y="Mean Absent Time")

# Barplot of Mean Absent Time v/s Drinking:
  ggplot(data=df, aes(x=Social_drinker, y=Absenteeism_time_in_hours, fill=Social_drinker)) + 
  stat_summary(fun.y = mean, geom = "bar") + ggtitle("Mean Absent Time vs Drinking") +
  labs(y="Mean Absent Time")
  
# Barplot of Mean Absent Time v/s Smoking:
  ggplot(data=df, aes(x=Social_smoker, y=Absenteeism_time_in_hours, fill=Social_smoker)) + 
  stat_summary(fun.y = mean, geom = "bar") + ggtitle("Mean Absent Time vs Smoking") +
  labs(y="Mean Absent Time")

# Barplot of Mean Absent Time v/s Education:
  ggplot(data=df, aes(x=Education, y=Absenteeism_time_in_hours, fill=Education)) + 
    stat_summary(fun.y = mean, geom = "bar") + ggtitle("Mean Absent Time vs Education") +
    labs(y="Mean Absent Time")
  
# Histogram showing distribution of Absent Time:
  ggplot(data=df, aes(x=Absenteeism_time_in_hours)) + geom_histogram(bins=15)


#### Feature Scaling ####  
# Using normalization method since the varibles do NOT follow Normal Distribution.
for (i in num_var)
{
  if(i=="Absenteeism_time_in_hours")
    { break } 
  df[,i] = (df[,i] - min(df[,i]))/(max(df[,i])-min(df[,i]))
}
  
#### Feature Selection ####

# Correlation heatmap for numerical variables in the dataset:
  
corr_mat = cor(select(df, num_var))
ggcorrplot(corr_mat,title = "Correlation Heatmap",type = "upper",ggtheme = theme_classic(),lab=TRUE)

# Removing variables "Weight" from the dataset:
df = df %>% select(-c("Weight"))

# ANOVA test for Categorical variable:

summary(aov(formula = Absenteeism_time_in_hours~Reason_for_absence, data = df))
summary(aov(formula = Absenteeism_time_in_hours~Month_of_absence, data = df))
summary(aov(formula = Absenteeism_time_in_hours~Day_of_the_week, data = df))
summary(aov(formula = Absenteeism_time_in_hours~Seasons, data = df))
summary(aov(formula = Absenteeism_time_in_hours~Disciplinary_failure, data = df))
summary(aov(formula = Absenteeism_time_in_hours~Education, data = df))
summary(aov(formula = Absenteeism_time_in_hours~Social_drinker, data = df))
summary(aov(formula = Absenteeism_time_in_hours~Social_smoker, data = df))




                            ################### MODEL DEVELOPMENT ####################

# Creating dummy variables for categorical variables
df = dummy_cols(df , select_columns = cat_var , remove_first_dummy = TRUE)
knitr::kable(df)

# Deleting the columns for which dummies are created
df = select(df, -cat_var)

# Splitting the dataset into training & testing set:
set.seed(123)
split = sample.split(df$Absenteeism_time_in_hours, SplitRatio = 0.30)
df_train = subset(df, split== TRUE)
df_test = subset(df, split== FALSE)

##### Using PCA

#principal component analysis
prin_comp = prcomp(select(df_train,-c("Absenteeism_time_in_hours")))

# Computing standard deviation of each principal component
std_dev = prin_comp$sdev

# Computing variance
pr_var = std_dev^2

# Proportion of variance explained
prop_varex = pr_var/sum(pr_var)

# cdf plot for principle components
plot(cumsum(prop_varex), xlab = "Principal Component",
     ylab = "Cumulative Proportion of Variance Explained",
     type = "b")

# Add a training set with principal components
pca_train = data.frame(prin_comp$x)

# From the above plot selecting 45 components since it explains almost 95+ % data variance
pca_train = pca_train[,1:45]

# Transforming test into PCA
pca_test = predict(prin_comp, newdata = select(df_test,-c("Absenteeism_time_in_hours")))
pca_test = as.data.frame(pca_test)

#select the first 45 components
pca_test = pca_test[,1:45]


# Including the target variable:
pca_train$Absenteeism_time_in_hours = paste(df_train$Absenteeism_time_in_hours)
pca_test$Absenteeism_time_in_hours  = paste(df_test$Absenteeism_time_in_hours)



#### Linear Regression Model ####

# Fitting the model to the training set:
lin_reg = lm(formula= Absenteeism_time_in_hours ~., data= pca_train)

# Predicting the values for the test set:
pred_lm = predict(lin_reg, newdata = pca_test)

# Results 
print(postResample(pred = pred_lm, obs= df_test$Absenteeism_time_in_hours))

#### Decision Tree Regression Model ####

# Fitting the model to the training set:
d_tree = rpart(formula= Absenteeism_time_in_hours ~., data= pca_train, method = "anova")

# Predicting the values for the test set:
pred_dt = predict(d_tree, newdata = pca_test)

# Results 
print(postResample(pred = pred_dt, obs= df_test$Absenteeism_time_in_hours))