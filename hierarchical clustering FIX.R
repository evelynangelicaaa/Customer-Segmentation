# Load required libraries
library(readxl)
library(cluster)
library(factoextra)
library(dendextend)
library(caret)

# Load data
data <- read_excel("C:/Kuliah/Sem 8/TA 2/data/FINAL DATA CLEANING.xlsx")
data$Hour_Difference <- log1p(data$Hour_Difference)

# Remove target variable (Loyalty_Label)
X <- data[, !(names(data) %in% c("Loyalty_Label"))]

# Perform hierarchical clustering using Ward's method
set.seed(123)
hc <- hclust(dist(X), method = "ward.D2")
#hc_complete <- hclust(dist(X), method = "complete")
#hc_single <- hclust(dist(X), method = "single")
#hc_avg <- hclust(dist(X), method = "average")
#set.seed(100)
#hc_centroid <- hclust(dist(X), method = "centroid")

# Plot dendrogram
plot(hc,
     main = expression(italic("Dendrogram of Hierarchical Clustering")),
     xlab = expression(italic("Data Points")),
     ylab = expression(italic("Height")),
     sub = "",
     cex = 0.6,         # ukuran label data (default: 1)
     cex.main = 1.5,    # ukuran judul
     cex.lab = 1.3,     # ukuran label sumbu
     cex.axis = 1.1     # ukuran angka pada sumbu
)

par(mar = c(10, 4, 4, 2))  # Increase bottom margin for better label spacing
plot(cut(as.dendrogram(hc), h = 4)$lower[[1]], cex = 0.1)

# Determine optimal number of clusters using Elbow Method
p <- fviz_nbclust(X, FUN = hcut, method = "wss")
p + 
  labs(title = "Optimal number of clusters",
       x = "Number of clusters k",
       y = "Total Within Sum of Square") +
  theme(
    plot.title = element_text(face = "italic"),
    axis.title.x = element_text(face = "italic"),
    axis.title.y = element_text(face = "italic")
  )

# Determine optimal number of clusters using Silhouette Method
fviz_nbclust(X, FUN = hcut, method = "silhouette")

# Cut tree into optimal number of clusters (adjust k based on analysis)
k_optimal <- 2  # Ganti dengan jumlah klaster optimal berdasarkan metode di atas
clusters <- cutree(hc, k = k_optimal)

# Visualize clusters with colored dendrogram
dend <- as.dendrogram(hc)
dend <- color_branches(dend, k = k_optimal)
par(cex = 1.0)
plot(dend, main = "Dendrogram Hasil Pengelompokan", ylab = expression(italic("Euclidean Distance")), 
     xlab = "Pelanggan", sub = "")

par(mar = c(10, 4, 4, 2))  # Increase bottom margin for better label spacing
plot(cut(as.dendrogram(dend), h = 4)$lower[[1]], cex = 0.5)
plot(cut(as.dendrogram(dend), h = 4)$lower[[35]], cex = 0.5)

plot(cut(as.dendrogram(dend), h = 4)$lower[[50]], cex = 1.0)
plot(cut(as.dendrogram(dend), h = 4)$lower[[93]], cex = 1.0)


# Append cluster labels to original data
data$Cluster <- as.factor(clusters)

# Plot frequency of clusters
cluster_counts <- as.data.frame(table(clusters))
names(cluster_counts) <- c("Cluster", "Frekuensi")

ggplot(cluster_counts, aes(x = Cluster, y = Frekuensi, fill = Cluster)) +
  geom_bar(stat = "identity") +
  theme_minimal() +
  ggtitle("Frekuensi Data pada Masing-Masing Cluster (Agglomerative)") +
  xlab("Klaster") +
  ylab("Frekuensi")

# Create scatter plot
# Create a scatterplot to visualize the clusters
ggplot(data = data, aes(x = X, y = Loyalty_Label, color = as.factor(clusters))) +
  geom_point(size = 4) +
  scale_color_manual(values = c("#1f77b4", "#ff7f0e", "#2ca02c")) +
  labs(title = "Car Segmentation Based on MPG and Horsepower",
       x = "Miles per Gallon (mpg)",
       y = "Horsepower (hp)") +
  theme_minimal()


# Transpose data (columns) ------------------------------------------------
hc.t <- hclust(dist(t(X)), method = "ward.D2")
plot(hc.t,
     main = expression(italic("Dendrogram of Hierarchical Clustering")),
     xlab = expression(italic("Data Points")),
     ylab = expression(italic("Height")),
     sub = "",
     cex = 1.0,         # ukuran label data (default: 1)
     cex.main = 1.5,    # ukuran judul
     cex.lab = 1.3,     # ukuran label sumbu
     cex.axis = 1.1     # ukuran angka pada sumbu
)

par(mar = c(10, 4, 4, 2))  # Increase bottom margin for better label spacing
plot(cut(as.dendrogram(hc.t), h = 5)$lower[[1]], cex = 0.1)

# Cut tree into optimal number of clusters (adjust k based on analysis)
k_optimal <- 2  # Ganti dengan jumlah klaster optimal berdasarkan metode di atas
clusters <- cutree(hc.t, k = k_optimal)

# Visualize clusters with colored dendrogram
dend.t <- as.dendrogram(hc.t)
dend.t <- color_branches(dend.t, k = k_optimal)
par(cex = 0.6)
par(mar = c(15, 4, 4, 2)) # Increase bottom margin for better label spacing
plot(dend.t, main = "Dendrogram Hasil Pengelompokan", ylab = expression(italic("Euclidean Distance")), 
     sub = "")

plot(cut(as.dendrogram(dend.t), h = 5)$lower[[1]], cex = 0.1)

data.t <- t(data)

# Append cluster labels to original data
data.t$Cluster <- as.factor(clusters)

# Plot frequency of clusters
cluster_counts.t <- as.data.frame(table(clusters))
names(cluster_counts.t) <- c("Cluster", "Frekuensi")

ggplot(cluster_counts.t, aes(x = Cluster, y = Frekuensi, fill = Cluster)) +
  geom_bar(stat = "identity") +
  theme_minimal() +
  ggtitle("Frekuensi Data pada Masing-Masing Cluster (Agglomerative)") +
  xlab("Klaster") +
  ylab("Frekuensi")

# Save clustered data to a new file
#write.csv(data, "Clustered_Data.csv", row.names = FALSE)

# Clustering Visualization ----------------------------------------------------
library(ggplot2)
library(dplyr)
library(tidyr)

# Ensure Cluster is a factor
data$Cluster <- as.factor(data$Cluster)

### 1️⃣ Bar Chart: Tipe_Bebas_Ongkir by Cluster ###
ggplot(data, aes(x = Cluster, fill = as.factor(Tipe_Bebas_Ongkir))) +
  geom_bar(position = "fill") +
  theme_minimal() +
  labs(title = "Distribusi Tipe Bebas Ongkir Berdasarkan Klaster", x = "Klaster", y = "Proporsi", fill = "Tipe Bebas Ongkir")

aggregate(Tipe_Bebas_Ongkir ~ Cluster, data = data, FUN = function(x) round(mean(x) * 100, 2))

### 2️⃣ Bar Chart: Loyalty_Label (Target) by Cluster ###
ggplot(data, aes(x = Cluster, fill = as.factor(Loyalty_Label))) +
  geom_bar(position = "fill") +
  theme_minimal() +
  labs(title = expression("Distribusi " * italic("Loyalty Label" ) * "Berdasarkan Klaster"), x = "Klaster", y = "Proporsi", fill = expression(italic("Loyalty Label")))

### 3️⃣ Boxplot: Hour_Difference by Cluster ###
ggplot(data, aes(x = Cluster, y = Hour_Difference, fill = Cluster)) +
  geom_boxplot() +
  theme_minimal() +
  labs(title = expression("Distribusi " * italic("Hour Difference" ) * "Berdasarkan Klaster"), x = "Klaster", y = "Proporsi", fill = expression(italic("Hour Difference")))

data %>%
  group_by(Cluster) %>%
  summarise(
    Mean = mean(Hour_Difference),
    Median = median(Hour_Difference),
    SD = sd(Hour_Difference),
    Min = min(Hour_Difference),
    Max = max(Hour_Difference)
  )

### 4️⃣ Stacked Bar Chart: Kategori_Penjualan_ Dummies by Cluster ###
kategori_cols <- grep("^Kategori_Penjualan_", names(data), value = TRUE)  # Select all Kategori_Penjualan_ dummies

data_kategori <- data %>%
  pivot_longer(cols = kategori_cols, names_to = "Kategori", values_to = "Value") %>%
  filter(Value == 1)  # Keep only rows where dummy is 1

ggplot(data_kategori, aes(x = Cluster, fill = Kategori)) +
  geom_bar(position = "fill") +
  theme_minimal() +
  labs(title = "Distribusi Kategori Penjualan Berdasarkan Klaster", x = "Klaster", y = "Proporsi")

aggregate(data[, c("Kategori_Penjualan_Sparepart","Kategori_Penjualan_Low_End","Kategori_Penjualan_High_End")], 
          by = list(Klaster = clusters), 
          FUN = mean)

### 5️⃣ Bar Chart: IsWeekend by Cluster ###
ggplot(data, aes(x = Cluster, fill = as.factor(IsWeekend))) +
  geom_bar(position = "fill") +
  theme_minimal() +
  labs(title = expression("Distribusi " * italic("IsWeekend " ) * "Berdasarkan Klaster"), x = "Klaster", y = "Proporsi", fill = expression(italic("IsWeekend")))

aggregate(IsWeekend ~ Cluster, data = data, FUN = function(x) round(mean(x) * 100, 2))

### 6️⃣ Stacked Bar Chart: Waktu Dummies by Cluster ###
waktu_cols <- grep("^Waktu", names(data), value = TRUE)  # Select all Waktu dummies

data_waktu <- data %>%
  pivot_longer(cols = waktu_cols, names_to = "Waktu", values_to = "Value") %>%
  filter(Value == 1)

ggplot(data_waktu, aes(x = Cluster, fill = Waktu)) +
  geom_bar(position = "fill") +
  theme_minimal() +
  labs(title = "Distribusi Waktu Berdasarkan Klaster", x = "Klaster", y = "Proporsi")

aggregate(data[, c("WaktuDini_Hari","WaktuPagi", "WaktuSiang", "WaktuMalam", "WaktuSore")], 
          by = list(Klaster = clusters), 
          FUN = mean)

### 7️⃣ Stacked Bar Chart: Kota Dummies by Cluster ###
kota_cols <- grep("^Kota", names(data), value = TRUE)  # Select all Kota dummies

data_kota <- data %>%
  pivot_longer(cols = kota_cols, names_to = "Kota", values_to = "Value") %>%
  filter(Value == 1)

ggplot(data_kota, aes(x = Cluster, fill = Kota)) +
  geom_bar(position = "fill") +
  theme_minimal() +
  labs(title = "Kota Berdasarkan Klaster", x = "Klaster", y = "Proportion")

#15 kota teratas
top_kota <- data_kota %>%
  count(Kota, sort = TRUE) %>%
  top_n(15, n)  # Show only the top 15 most frequent cities

data_kota_filtered <- data_kota %>%
  filter(Kota %in% top_kota$Kota)

ggplot(data_kota_filtered, aes(x = Cluster, fill = Kota)) +
  geom_bar(position = "fill") +
  theme_minimal() +
  labs(title = "Distribusi 15 Kota Teratas Berdasarkan Klaster", x = "Klaster", y = "Proporsi", fill = "Kota") +
  theme(legend.position = "bottom")

total_kota <- colSums(data[, kota_cols])
top_15_kota <- names(sort(total_kota, decreasing = TRUE)[1:15])
data_kota_top15 <- data[, c("Cluster", top_15_kota)]
proporsi_kota_top15 <- aggregate(. ~ Cluster, data = data_kota_top15, FUN = mean)
as.data.frame(proporsi_kota_top15)

# Metrik Evaluasi ---------------------------------------------------------

# Ambil label asli
true_labels <- data$Loyalty_Label

# Lihat perbandingan hasil klaster dan label asli
table(clusters, true_labels)

# 1. Mapping hasil cluster ke label prediksi
predicted_labels <- ifelse(clusters == 1, 0, 1)  # Cluster 1 = not loyal, lainnya = loyal

# 2. Ubah jadi faktor agar cocok dengan confusionMatrix
true_labels <- as.factor(data$Loyalty_Label)
predicted_labels <- as.factor(predicted_labels)

# 3. Evaluasi menggunakan confusionMatrix 
conf_res <- confusionMatrix(predicted_labels, true_labels)

# 4. Ambil metrik-metrik evaluasi
accuracy <- conf_res$overall['Accuracy']
precision <- conf_res$byClass['Precision']
recall <- conf_res$byClass['Recall']
f1_score <- conf_res$byClass['F1']

# 5. Buat summary dalam bentuk data frame
eval_summary <- data.frame(
  Metric = c("Accuracy", "Precision (Class 0)", "Recall (Class 0)", "F1-Score (Class 0)"),
  Value = c(accuracy, precision, recall, f1_score)
)

# 6. Tampilkan tabel
print(eval_summary)



# Variable Importance -----------------------------------------------------

# Bind cluster labels to the feature data
X_with_cluster <- X
X_with_cluster$cluster <- as.factor(clusters)

# Build random forest
set.seed(123)
rf_model <- randomForest(cluster ~ ., data = X_with_cluster, importance = TRUE)

# Show variable importance
importance(rf_model)
varImpPlot(rf_model)
