library(readxl) #untuk membaca data excel
library(writexl)
library(lubridate) #untuk extract fitur dari tanggal
library(DescTools)
library(tidyr)
library(dplyr)
library(rpart) #package decision tree
library(rpart.plot) #plot decision Tree
library(randomForest) #random forest
library(xgboost) #package xgboost
library(e1071) #package SVM
library(rsample) #untuk split data
library(corrplot)
library(plotly)
library(caret)
library(ggplot2)
library(pROC)
library(keras)
library(tensorflow)
library(Rtsne)
library(cluster)
library(factoextra)
library(klaR)
library(reshape2)

# DATA RAW ----------------------------------------------------------------
data <- read_excel("C:/Kuliah/Sem 8/TA 2/data/FINAL DATA.xlsx")
str(data)
summary(data)


# PRE PROCESSING ----------------------------------------------------------
# TIDAK = 0
# YA = 1

# Variabel Target (y) -----------------------------------------------------
# Hitung jumlah transaksi unik per nomor telepon
loyalty_counts <- data %>%
  group_by(No_Telp_Pembeli) %>%
  summarise(Nomor_Invoice_Count = n_distinct(Nomor_Invoice))

# Buat kolom baru 'Loyalty_Label'
data <- data %>%
  left_join(loyalty_counts, by = "No_Telp_Pembeli") %>%
  mutate(Loyalty_Label = ifelse(Nomor_Invoice_Count > 1, 1, 0))

data <- data %>% dplyr::select(-Nomor_Invoice_Count)

# Diubah menjadi faktor
data$Loyalty_Label <- as.factor(data$Loyalty_Label)
plot(x = data$Loyalty_Label, main = 'Loyalty Customer')

# Encoding ----------------------------------------------
#Mengubah format tanggal
data$Tanggal_Pembayaran <- as.POSIXct(data$Tanggal_Pembayaran, format = "%d-%m-%Y %H:%M:%S")

# Combine date and time into a single column
data <- data %>%
  mutate(
    Tanggal_Pengiriman = as.POSIXct(paste(Tanggal_Pengiriman_Barang, Waktu_Pengiriman_Barang), format="%d-%m-%Y %H:%M:%S"),
    Hour_Difference = as.numeric(difftime(Tanggal_Pengiriman, Tanggal_Pembayaran, units = "hours")), # Calculate hours difference
    `Total_Penjualan_(IDR)` = as.numeric(`Total_Penjualan_(IDR)`)
    )

data <- data %>%
  mutate(Kategori_Penjualan = case_when(
    `Total_Penjualan_(IDR)` < 500000 ~ "Sparepart",
    `Total_Penjualan_(IDR)` >= 500000 & `Total_Penjualan_(IDR)` <= 1500000 ~ "Low End",
    `Total_Penjualan_(IDR)` > 1500000 ~ "High End"
  ))

data <- data %>%
  mutate(
    Kategori_Penjualan_Sparepart = ifelse(Kategori_Penjualan == "Sparepart", 1, 0),
    Kategori_Penjualan_Low_End = ifelse(Kategori_Penjualan == "Low End", 1, 0),
    Kategori_Penjualan_High_End = ifelse(Kategori_Penjualan == "High End", 1, 0)
  )

data$Weekday <- weekdays(data$Tanggal_Pembayaran) # Weekday Name
# Weekend Binary Encoding (1 = Saturday/Sunday, 0 = Weekday)
data$IsWeekend <- ifelse(data$Weekday %in% c("Saturday", "Sunday"), 1, 0)

# Ekstrak jam dari Tanggal_Pembayaran
data$Jam <- format(data$Tanggal_Pembayaran, "%H")
data$Jam <- as.numeric(data$Jam)

# Buat kategori waktu
data$Waktu <- cut(data$Jam, 
                breaks = c(-Inf, 4.99, 10.99, 15.99, 18.99, 24.99),
                labels = c("Dini_Hari", "Pagi", "Siang", "Sore", "Malam"))

# Konversi ke dummy variabel
dummy_waktu <- model.matrix(~Waktu - 1, data = data)

# Ubah kolom 'Tipe_Bebas_Ongkir' menjadi 0 dan 1
data <- data %>%
  mutate(Tipe_Bebas_Ongkir = ifelse(Tipe_Bebas_Ongkir == "Bebas Ongkir", 1, 0))

# Create dummy variables for Kota
dummy_kota <- model.matrix(~ Kota - 1, data = data)  # Removes intercept

# Convert matrix to dataframe and rename columns
dummy_kota <- as.data.frame(dummy_kota)

# Merge dummy variables with original dataset
data <- cbind(data, dummy_waktu, dummy_kota)

# Remove Original Column
data <- subset(data, select = -c(Nomor_Invoice, No_Telp_Pembeli, `Total_Penjualan_(IDR)`,
                                 Kategori_Penjualan, Tanggal_Pembayaran, 
                                 Tanggal_Pengiriman_Barang, Waktu_Pengiriman_Barang, 
                                 Tanggal_Pengiriman, Weekday, Kota, Provinsi, Jam, Waktu))

#write_xlsx(data, "C:/Kuliah/Sem 8/TA 2/data/FINAL DATA CLEANING2.xlsx")

# DATA CLEAN --------------------------------------------------------------
#data <- read_excel("C:/Kuliah/Sem 8/TA 2/data/FINAL DATA CLEANING.xlsx")
#str(data)
#summary(data)

# Diubah menjadi faktor
#data$Loyalty_Label <- as.factor(data$Loyalty_Label)
#plot(x = data$Loyalty_Label, main = 'Loyalty Customer')




# Hapus NA ----------------------------------------------------------------
Abstract(data)
PlotMiss(data) #plot missing values
data <- na.omit(data)

# Transformasi var numerik  ------------------------------------------------------------
hist(data$Hour_Difference,
     main = expression("Histogram dari " * italic("Hour Difference")),
     xlab = expression(italic("Hour Difference")),
     ylab = "Frekuensi") #skewed ke kanan
# Transformasi logaritma untuk Hour_Difference
data$Hour_Difference <- log1p(data$Hour_Difference)
# Visualisasi ulang untuk melihat perubahan distribusi
hist(data$Hour_Difference, 
     main = "Histogram Setelah Transformasi Logaritma",
     xlab = expression(italic("Hour Difference")),
     ylab = "Frekuensi",
     col = "lightblue")

# Membagi dataset menjadi train dan test ----------------------------------
set.seed(123)
data_split <- initial_split(data, prop = .80)
data_train <- training(data_split)
data_test <- testing(data_split)

colnames(data_train) <- make.names(colnames(data_train))
colnames(data_test) <- make.names(colnames(data_test))

# Decision Tree ----------------------------------------------------------
fit <- rpart(Loyalty_Label ~ ., 
                    data = data_train, 
                    method = "class", 
                    control = rpart.control(cp = 0.001, minsplit = 1, maxdepth = 5),
                    parms = list(split = "gini"))  # Kembali ke default


summary(fit)
fit$variable.importance
barplot(fit$variable.importance)

# Plot the tree
rpart.plot(fit, box.palette = "Blues", cex= 0.75)
#rpart.plot(fit, type=4, extra=2, clip.right.labs=FALSE, varlen=0, faclen=3)

# Random Forest -----------------------------------------------------------
forest_fit1 <- randomForest(formula = Loyalty_Label~.,
                            data = data_train,
                            ntree = 100) #menggunakan 100 tree
plot(forest_fit1)

#importance_scores <- importance(forest_fit1)
#selected_features <- names(sort(importance_scores[,1], decreasing = TRUE))[1:2]
#print(selected_features)

# SVM ---------------------------------------------------------------------
# Menyiapkan data untuk SVM
data_train$Loyalty_Label <- as.factor(data_train$Loyalty_Label)  # Pastikan target adalah faktor (kategori)
data_test$Loyalty_Label <- as.factor(data_test$Loyalty_Label)

# Pisahkan fitur dan target
train_features <- data_train %>% dplyr::select(-Loyalty_Label)  # Ambil semua kolom kecuali kolom target
test_features <- data_test %>% dplyr::select(-Loyalty_Label)

# Pisahkan label
train_label <- data_train$Loyalty_Label
test_label <- data_test$Loyalty_Label

# creating kernels list
kernel_settings <- c("linear", "radial", "polynomial", "sigmoid")

# testing each kernels
for (kernel in kernel_settings) {
  # Train model SVM
  svm_model <- svm(
    Loyalty_Label ~ .,
    data = data_train,
    kernel = kernel
  )
  
  # performing predictions with each kernels
  predictions <- predict(svm_model, data_test)
  
  # counting accuracy from each kernels
  accuracy <- sum(predictions == data_test$Loyalty_Label) / nrow(data_test)
  
  # Print accruracy for each kernels
  cat("Kernel:", kernel, "| Accuracy:", accuracy, "\n")
  
}

##Kernel linear 0.6363636
svm_model <- svm(Loyalty_Label ~ ., data = data_train, 
                 kernel = "linear", cost = 1, gamma = 0.1)
svm_model

# Prediksi menggunakan model SVM
svm_pred <- predict(svm_model, newdata = test_features)
confusionMatrix(svm_pred, test_label)

# If using a linear kernel, we can extract the coefficients
if (svm_model$kernel == "linear") {
  # Get the coefficients
  coefficients <- as.data.frame(t(svm_model$coefs) %*% svm_model$SV)
  variable_importance <- abs(coefficients)
  
  # Create a data frame for better visualization
  importance_df <- data.frame(Variable = rownames(variable_importance), Importance = variable_importance)
  importance_df <- importance_df[order(-importance_df$Importance), ]  # Sort by importance
  
  # Print variable importance
  print(importance_df)
}


##T-test for SVM result
#Comparing Predictions vs. Actual Labels
# Convert factors to numeric for comparison
svm_pred_numeric <- as.numeric(svm_pred)
test_label_numeric <- as.numeric(test_label)

# Perform a paired t-test
t_test_result <- t.test(svm_pred_numeric, test_label_numeric, paired = TRUE)

# Print the t-test result
print(t_test_result)

#p-value ≥ 0.05 → No significant difference, meaning the two models perform similarly.

# SVM Visualization -------------------------------------------------------
library(ggplot2)
library(dplyr)
library(reshape2)
library(caret)
library(e1071)
library(pROC)
library(ggpubr)

# 1. Confusion Matrix Heatmap
svm_pred <- predict(svm_model, newdata = test_features)
conf_matrix <- table(Predicted = svm_pred, Actual = test_label)
conf_matrix_df <- as.data.frame(conf_matrix)

ggplot(conf_matrix_df, aes(x = Actual, y = Predicted, fill = Freq)) +
  geom_tile() +
  scale_fill_gradient(low = "white", high = "red") +
  geom_text(aes(label = Freq), color = "black", size = 5) +
  theme_minimal() +
  ggtitle("Confusion Matrix Heatmap")

# 2. ROC Curve
svm_prob <- attributes(predict(svm_model, newdata = test_features, decision.values = TRUE))$decision.values
roc_curve <- roc(test_label, as.numeric(svm_prob))
plot(roc_curve, main = "ROC Curve for SVM", col = "blue", lwd = 2)

# 3. Barplot: Perbandingan Prediksi vs Aktual untuk Variabel Tipe_Bebas_Ongkir
ggplot(data_test, aes(x = Tipe_Bebas_Ongkir, fill = as.factor(svm_pred))) +
  geom_bar(position = "dodge") +
  theme_minimal() +
  ggtitle("Prediksi SVM terhadap Tipe Bebas Ongkir")


# 4. Barplot: Perbandingan Prediksi vs Aktual untuk Variabel IsWeekend
ggplot(data_test, aes(x = IsWeekend, fill = as.factor(svm_pred))) +
  geom_bar(position = "dodge") +
  theme_minimal() +
  ggtitle("Prediksi SVM terhadap IsWeekend")



# 5. Barplot: Perbandingan Prediksi vs Aktual untuk Variabel Dummies (Kategori Penjualan)
kategori_cols <- grep("^Kategori_Penjualan_", names(data_test), value = TRUE)

data_kategori <- data_test %>%
  pivot_longer(cols = kategori_cols, names_to = "Kategori", values_to = "Value") %>%
  filter(Value == 1)

ggplot(data_kategori, aes(x = Kategori, fill = as.factor(svm_pred))) +
  geom_bar(position = "dodge") +
  theme_minimal() +
  ggtitle("Prediksi SVM terhadap Kategori Penjualan")

# 6. Bar Plot: Prediksi SVM vs Dummy Variable (Waktu)
waktu_cols <- grep("^Waktu", names(data_test), value = TRUE)

data_waktu <- data_test %>%
  pivot_longer(cols = waktu_cols, names_to = "Waktu", values_to = "Value") %>%
  filter(Value == 1)

ggplot(data_waktu, aes(x = Waktu, fill = as.factor(svm_pred))) +
  geom_bar(position = "dodge") +
  theme_minimal() +
  ggtitle("Prediksi SVM terhadap Waktu")

# 7. Facet Grid: Kota terhadap Prediksi SVM
kota_cols <- grep("^Kota", names(data_test), value = TRUE)

data_kota <- data_test %>%
  pivot_longer(cols = kota_cols, names_to = "Kota", values_to = "Value") %>%
  filter(Value == 1)

ggplot(data_kota, aes(x = Kota, fill = as.factor(svm_pred))) +
  geom_bar(position = "dodge") +
  facet_wrap(~ Loyalty_Label, scales = "free_x") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
  ggtitle("Distribusi Prediksi SVM berdasarkan Kota")

# 8. Boxplot: Hour_Difference 
ggplot(data_test, aes(x = Hour_Difference, fill = as.factor(svm_pred))) +
  geom_boxplot() +
  theme_minimal() +
  labs(title = "Distribusi Prediksi SVM berdasarkan Hour Difference")

# XGBoost -----------------------------------------------------------------
# Menyiapkan data untuk XGBoost
data_train$Loyalty_Label <- as.numeric(data_train$Loyalty_Label) - 1
data_test$Loyalty_Label <- as.numeric(data_test$Loyalty_Label) - 1

train_matrix <- xgb.DMatrix(
  data = as.matrix(data_train %>% dplyr::select(-c("Loyalty_Label"))),
  label = data_train$Loyalty_Label
)

test_matrix <- xgb.DMatrix(
  data = as.matrix(data_test %>% dplyr::select(-c("Loyalty_Label"))),
  label = data_test$Loyalty_Label
)

# Model XGBoost
xgb_model <- xgboost(data = train_matrix, 
                     nrounds = 100, objective = "binary:logistic")
# Prediksi XGBoost
xgb_pred_prob <- predict(xgb_model, test_matrix)
xgb_pred <- ifelse(xgb_pred_prob > 0.5, 1, 0)
xgb_pred <- as.factor(xgb_pred)

# Plot Feature Importance
xgb.importance(model = xgb_model) %>%
  xgb.plot.importance(top_n = 10)


# Evaluasi Metrik ---------------------------------------------------------
# Ensure the actual labels are factors
data_test$Loyalty_Label <- as.factor(data_test$Loyalty_Label)

###Decision Tree
dt_pred <- predict(fit, data_test, type = "class")
dt_pred <- factor(dt_pred, levels = levels(data_test$Loyalty_Label))
conf_matrix_dt <- confusionMatrix(dt_pred, data_test$Loyalty_Label)

# Extracting metrics
dt_accuracy <- conf_matrix_dt$overall['Accuracy']
dt_precision <- conf_matrix_dt$byClass['Precision']
dt_recall <- conf_matrix_dt$byClass['Recall']
dt_f1 <- conf_matrix_dt$byClass['F1']

###Random Forest
rf_pred <- predict(forest_fit1, data_test)
#rf_pred <- factor(rf_pred, levels = levels(data_test$Loyalty_Label))
conf_matrix_rf <- confusionMatrix(rf_pred, data_test$Loyalty_Label)

# Extracting metrics
rf_accuracy <- conf_matrix_rf$overall['Accuracy']
rf_precision <- conf_matrix_rf$byClass['Precision']
rf_recall <- conf_matrix_rf$byClass['Recall']
rf_f1 <- conf_matrix_rf$byClass['F1']


###SVM
conf_matrix_svm <- confusionMatrix(svm_pred, test_label)

# Extracting metrics
svm_accuracy <- conf_matrix_svm$overall['Accuracy']
svm_precision <- conf_matrix_svm$byClass['Precision']
svm_recall <- conf_matrix_svm$byClass['Recall']
svm_f1 <- conf_matrix_svm$byClass['F1']


###XGBoost
conf_matrix_xgb <- confusionMatrix(xgb_pred, as.factor(data_test$Loyalty_Label))

# Extracting metrics
xgb_accuracy <- conf_matrix_xgb$overall['Accuracy']
xgb_precision <- conf_matrix_xgb$byClass['Precision']
xgb_recall <- conf_matrix_xgb$byClass['Recall']
xgb_f1 <- conf_matrix_xgb$byClass['F1']


##SUMMARY
results <- data.frame(
  Model = c("Decision Tree", "Random Forest", "SVM", "XGBoost"),
  Accuracy = c(dt_accuracy, rf_accuracy, svm_accuracy, xgb_accuracy),
  Precision = c(dt_precision, rf_precision, svm_precision, xgb_precision),
  Recall = c(dt_recall, rf_recall, svm_recall, xgb_recall),
  F1_Score = c(dt_f1, rf_f1, svm_f1, xgb_f1)
)

print(results)

# Variable Importance -----------------------------------------------------
importance_dt <- fit$variable.importance
importance_dt <- as.data.frame(importance_dt)
colnames(importance_dt) <- "Decision_Tree"
importance_dt$Feature <- rownames(importance_dt)

importance_rf <- as.data.frame(importance(forest_fit1)[, 1])  # Ambil MeanDecreaseGini
colnames(importance_rf) <- "Random_Forest"
importance_rf$Feature <- rownames(importance_rf)

importance_xgb <- xgb.importance(model = xgb_model)
importance_xgb <- importance_xgb[, c("Feature", "Gain")]
colnames(importance_xgb) <- c("Feature", "XGBoost")

# Gabungkan semua berdasarkan kolom "Feature"
importance_all <- Reduce(function(x, y) merge(x, y, by = "Feature", all = TRUE),
                         list(importance_dt, importance_rf, importance_xgb))

# Mengganti NA dengan 0 supaya bisa dibandingkan
importance_all[is.na(importance_all)] <- 0

# Urutkan berdasarkan rata-rata importance
importance_all$Average_Importance <- rowMeans(importance_all[, -1])
importance_all <- importance_all[order(-importance_all$Average_Importance), ]

print(importance_all)

library(reshape2)
library(ggplot2)

# Ubah ke format long
importance_long <- melt(importance_all[, -ncol(importance_all)], id.vars = "Feature")

# Plot
ggplot(importance_long, aes(x = reorder(Feature, value), y = value, fill = variable)) +
  geom_bar(stat = "identity", position = "dodge") +
  coord_flip() +
  labs(title = "Perbandingan Variable Importance",
       x = "Feature", y = "Importance", fill = "Model") +
  theme_minimal()


##SVM
library(mlbench)
library(caret)

# Fit ulang dengan caret agar bisa pakai varImp
ctrl <- trainControl(method = "cv", number = 5)
svm_fit <- caret::train(
  Loyalty_Label ~ ., data = data_train,
  method = "svmLinear",
  trControl = ctrl
)

# Importance dari model linear SVM via caret
importance_svm <- varImp(svm_fit)
print(importance_svm)
plot(importance_svm, top = 10)


# Distribusi var. Kota ----------------------------------------------------
# Cek nama kolom untuk identifikasi target dan kota
colnames(data)

# Gantilah "Target_Variable" dengan nama kolom target yang sesuai
target_col <- "Loyalty_Label"

# Pilih hanya kolom kota dan target
city_cols <- grep("KotaKota", names(data), value = TRUE)
data_city <- data %>% select(all_of(c(city_cols, target_col)))

# Hitung distribusi target berdasarkan kota
city_distribution <- data_city %>%
  pivot_longer(cols = all_of(city_cols), names_to = "City", values_to = "Presence") %>%
  filter(Presence == 1) %>%
  group_by(City, !!sym(target_col)) %>%
  summarise(Count = n(), .groups = 'drop') %>%
  mutate(Percentage = Count / sum(Count) * 100)

# Visualisasi distribusi target berdasarkan kota
ggplot(city_distribution, aes(x = City, y = Percentage, fill = as.factor(!!sym(target_col)))) +
  geom_bar(stat = "identity", position = "dodge") +
  coord_flip() +
  labs(title = "Distribusi Target Berdasarkan Kota", x = "Kota", y = "Persentase", fill = "Target") +
  theme_minimal()

