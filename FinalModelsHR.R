#' ---
#' title: "Case 2: HR Analytics"
#' author:  "Hunter Berberich"
#' date: 8/1/2024
#' output:
#'  pdf_document:
#'   pandoc_args: --listings
#' header-includes:
#'  \usepackage{fancyhdr}
#'  \usepackage{background}
#'  \lstset{breaklines=true, keywordstyle=\color{blue}, commentstyle=\color{teal}, numberstyle=\color{blue},stringstyle=\color{teal},}
#'  \pagestyle{fancy}
#'  \fancyhead[LO,LE]{BAN 840 - Predictive Analytics for Business}
#'  \fancyhead[RO,RE]{Homework 10}
#'  \fancypagestyle{plain}{\pagestyle{fancy}}
#'  \backgroundsetup{scale = 8, color=gray, opacity=0.3, contents={}}
#'  \let\OldRule\rule
#'  \renewcommand{\rule}[2]{\OldRule{\linewidth}{#2}}
#' ---

#+ echo = FALSE
# DO NOT CHANGE THE CODE IN THIS SECTION --------------------------------------

# include formatR and styler for text wrapping in PDF output
library(formatR)
library(styler)

# set working dir to source file location
library("rstudioapi")
setwd(dirname(getSourceEditorContext()$path))


#+ echo = FALSE
# clear all objects
rm(list = ls())

#+ echo = TRUE
#+ message = FALSE

#' 
#' ---
#' 
#' # \textcolor{blue}{Setup code}
#' 
#' **Add code here to load required packages, read data, etc. - preliminary tasks you need to perform to answer the HW question.**
#' 

# Your code here.

# Load necessary libraries
library(dplyr)
library(ggplot2)
library(caret)
library(MASS)
library(pROC)
library(openxlsx)
library(rpart)
library(rpart.plot)
library(reshape2)

#' # \textcolor{blue}{Data Cleaning and Preprocessing}

data <- read.csv("/Users/hunterberberich/Desktop/School Stuff!/Data/Case2DataClean.csv")

data$Education <- factor(data$Education, levels = 1:5, labels = c("Below College", "College", "Bachelor", "Master", "Doctor"))
data$EnvironmentSatisfaction <- factor(data$EnvironmentSatisfaction, levels = 1:4, labels = c("Low", "Medium", "High", "Very High"))
data$JobInvolvement <- factor(data$JobInvolvement, levels = 1:4, labels = c("Low", "Medium", "High", "Very High"))
data$JobSatisfaction <- factor(data$JobSatisfaction, levels = 1:4, labels = c("Low", "Medium", "High", "Very High"))
data$PerformanceRating <- factor(data$PerformanceRating, levels = 1:4, labels = c("Low", "Good", "Excellent", "Outstanding"))
data$RelationshipSatisfaction <- factor(data$RelationshipSatisfaction, levels = 1:4, labels = c("Low", "Medium", "High", "Very High"))
data$WorkLifeBalance <- factor(data$WorkLifeBalance, levels = 1:4, labels = c("Bad", "Good", "Better", "Best"))
data$BusinessTravel <- factor(data$BusinessTravel)
data$Department <- factor(data$Department)
data$EducationField <- factor(data$EducationField)
data$Gender <- factor(data$Gender)
data$JobRole <- factor(data$JobRole)
data$MaritalStatus <- factor(data$MaritalStatus)
data$OverTime <- factor(data$OverTime)
data$Attrition <- factor(data$Attrition, levels = c("No", "Yes"))

#' # \textcolor{blue}{Model Building}
set.seed(123)
index <- createDataPartition(data$Attrition, p = 0.8, list = FALSE)
train_data <- data[index, ]
test_data <- data[-index, ]

initial_model <- glm(Attrition ~ ., data = train_data, family = binomial)
print(initial_model)

stepwise_model <- step(initial_model, direction = "both")
print(stepwise_model)

basic_tree <- rpart(Attrition ~ ., data = train_data, method = "class", control = rpart.control(minsplit = 2, minbucket = 1, maxsurrogate = 0))
rpart.plot(basic_tree, main = "Basic Classification Tree")

full_tree <- rpart(Attrition ~ ., data = train_data, method = "class", cp = 0, control = rpart.control(minsplit = 2, minbucket = 1, maxsurrogate = 0))
best_cp <- full_tree$cptable[which.min(full_tree$cptable[,"xerror"]),"CP"]
pruned_tree <- prune(full_tree, cp = best_cp)
rpart.plot(pruned_tree, main = "Pruned Classification Tree")
print(pruned_tree)
#' # \textcolor{blue}{Model Evaluation}
create_confusion_matrices <- function(model, test_data, thresholds, model_type) {
  results <- list()
  if (model_type == "logistic") {
    predictions <- predict(model, newdata = test_data, type = "response")
  } else {
    predictions <- predict(model, newdata = test_data, type = "prob")[,2]
  }
  for (threshold in thresholds) {
    predicted_classes <- factor(ifelse(predictions > threshold, "Yes", "No"), levels = c("No", "Yes"))
    confusion <- confusionMatrix(predicted_classes, test_data$Attrition, positive = "Yes")
    results[[paste0("Threshold_", threshold)]] <- confusion
  }
  return(results)
}

evaluate_model_performance <- function(confusion) {
  list(
    Accuracy = confusion$overall["Accuracy"],
    BalancedAccuracy = confusion$byClass["Balanced Accuracy"]
  )
}

thresholds <- seq(0.05, 0.95, by = 0.05)

confusion_matrices_lr <- create_confusion_matrices(stepwise_model, test_data, thresholds, "logistic")
performance_lr <- do.call(rbind, lapply(names(confusion_matrices_lr), function(x) {
  metrics <- evaluate_model_performance(confusion_matrices_lr[[x]])
  data.frame(
    Model = "Logistic Regression",
    Threshold = gsub("Threshold_", "", x),
    Accuracy = metrics$Accuracy,
    BalancedAccuracy = metrics$BalancedAccuracy
  )
}))

confusion_matrices_pruned <- create_confusion_matrices(pruned_tree, test_data, thresholds, "tree")
performance_pruned <- do.call(rbind, lapply(names(confusion_matrices_pruned), function(x) {
  metrics <- evaluate_model_performance(confusion_matrices_pruned[[x]])
  data.frame(
    Model = "Pruned Tree",
    Threshold = gsub("Threshold_", "", x),
    Accuracy = metrics$Accuracy,
    BalancedAccuracy = metrics$BalancedAccuracy
  )
}))

performance_combined <- rbind(performance_lr, performance_pruned)
write.xlsx(performance_combined, "/Users/hunterberberich/Desktop/School Stuff!/Data/ModelPerformanceCombined.xlsx")

print(performance_combined)

options(repr.plot.width = 12, repr.plot.height = 6)

performance_lr_long <- melt(performance_lr, id.vars = c("Model", "Threshold"), measure.vars = c("Accuracy", "BalancedAccuracy"))
ggplot(performance_lr_long, aes(x = as.factor(Threshold), y = value, color = variable, group = variable)) +
  geom_line(size = 1.0) +
  geom_point(size = 2) +
  theme_minimal() +
  labs(title = "Logistic Regression Performance", x = "Threshold", y = "Value") +
  scale_color_brewer(palette = "Set1") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1, size = 9),
        axis.text.y = element_text(size = 12),
        plot.title = element_text(size = 16, face = "bold"),
        legend.title = element_text(size = 12),
        legend.text = element_text(size = 10))

ggsave("Logistic_Regression_Performance_Metrics.png", width = 12, height = 6)

performance_tree_long <- melt(performance_pruned, id.vars = c("Model", "Threshold"), measure.vars = c("Accuracy", "BalancedAccuracy"))
ggplot(performance_tree_long, aes(x = as.factor(Threshold), y = value, color = variable, group = variable)) +
  geom_line(size = 1.0) +
  geom_point(size = 2) +
  theme_minimal() +
  labs(title = "Pruned Tree Performance Metrics", x = "Threshold", y = "Value") +
  scale_color_brewer(palette = "Set1") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1, size = 12),
        axis.text.y = element_text(size = 12),
        plot.title = element_text(size = 16, face = "bold"),
        legend.title = element_text(size = 12),
        legend.text = element_text(size = 10))



plot_roc <- function(model, test_data, model_type) {
  if (model_type == "logistic") {
    predictions <- predict(model, newdata = test_data, type = "response")
  } else {
    predictions <- predict(model, newdata = test_data, type = "prob")[,2]
  }
  roc_obj <- roc(test_data$Attrition, predictions, levels = rev(levels(test_data$Attrition)))
  plot(roc_obj, main = paste(model_type, "ROC Curve"), col = "blue")
  return(roc_obj)
}

roc_lr <- plot_roc(stepwise_model, test_data, "logistic")
roc_pruned <- plot_roc(pruned_tree, test_data, "tree")

#' # \textcolor{blue}{Cost Analysis}
calculate_cost_savings <- function(confusion_matrix, total_leaving_employees, replacement_cost, raise_cost) {
  false_negatives <- confusion_matrix$table[1, 2]
  false_positives <- confusion_matrix$table[2, 1]
  
  cost_without_model <- total_leaving_employees * replacement_cost
  
  cost_with_model <- (false_negatives * replacement_cost) + (false_positives * raise_cost)
  
  cost_savings <- cost_without_model - cost_with_model
  
  return(list(cost_without_model = cost_without_model, cost_with_model = cost_with_model, cost_savings = cost_savings))
}

annual_salary <- 78035
replacement_cost <- 1.5 * annual_salary
raise_cost <- 0.1 * annual_salary

cost_savings_results_lr <- lapply(names(confusion_matrices_lr), function(threshold) {
  confusion_matrix <- confusion_matrices_lr[[threshold]]
  total_leaving_employees <- sum(test_data$Attrition == "Yes")
  cost_savings <- calculate_cost_savings(confusion_matrix, total_leaving_employees, replacement_cost, raise_cost)
  return(data.frame(Threshold = as.numeric(gsub("Threshold_", "", threshold)), 
                    Model = "Logistic Regression",
                    Cost_Savings = cost_savings$cost_savings))
})

cost_savings_results_pruned <- lapply(names(confusion_matrices_pruned), function(threshold) {
  confusion_matrix <- confusion_matrices_pruned[[threshold]]
  total_leaving_employees <- sum(test_data$Attrition == "Yes")
  cost_savings <- calculate_cost_savings(confusion_matrix, total_leaving_employees, replacement_cost, raise_cost)
  return(data.frame(Threshold = as.numeric(gsub("Threshold_", "", threshold)), 
                    Model = "Pruned Tree",
                    Cost_Savings = cost_savings$cost_savings))
})

cost_savings_df_lr <- do.call(rbind, cost_savings_results_lr)
cost_savings_df_pruned <- do.call(rbind, cost_savings_results_pruned)
cost_savings_df <- rbind(cost_savings_df_lr, cost_savings_df_pruned)

ggplot(cost_savings_df, aes(x = Threshold, y = Cost_Savings, color = Model, group = Model)) +
  geom_line(size = 1.0) +
  geom_point(size = 2) +
  scale_x_continuous(breaks = seq(0, 1, by = 0.05)) +
  theme_minimal() +
  labs(title = "Cost Savings at Different Thresholds", x = "Threshold", y = "Cost Savings") +
  scale_color_brewer(palette = "Set1") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1, size = 9),
        axis.text.y = element_text(size = 12),
        plot.title = element_text(size = 16, face = "bold"),
        legend.title = element_text(size = 12),
        legend.text = element_text(size = 10))

