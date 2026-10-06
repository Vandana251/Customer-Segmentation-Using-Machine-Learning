# Customer Segmentation using Unsupervised Learning

An end-to-end Machine Learning project analyzing customer behaviors and demographics to perform market segmentation using Unsupervised Learning techniques.

---

## 📌 Project Overview
Customer segmentation is the practice of dividing a customer base into distinct groups that share similar characteristics, purchasing behaviors, and demographic profiles. This project leverages unsupervised clustering algorithms and dimensionality reduction to identify high-value customer groups and actionable business segments.

---

## 🛠️ Tech Stack & Libraries
- **Language**: Python 3.x
- **Data Manipulation**: `pandas`, `numpy`
- **Data Visualization**: `matplotlib`, `seaborn`
- **Machine Learning**: `scikit-learn`, `scipy`
  - **Clustering**: K-Means, Agglomerative (Hierarchical) Clustering, DBSCAN
  - **Dimensionality Reduction**: Principal Component Analysis (PCA)
  - **Evaluation Metrics**: Silhouette Score, Calinski-Harabasz Score, Davies-Bouldin Index

---

## 📊 Dataset Description
- **File**: `customer_data.csv`
- **Shape**: 10,000 rows × 42 columns
- **Features Include**:
  - **Demographics**: Birth Year, Education Level, Marital Status, Income, Household composition
  - **Spending Behavior**: Amount spent across different categories (Wine, Fruits, Meat, Fish, Sweet, Gold, etc.)
  - **Engagement & Channels**: Web purchases, Catalog purchases, Store purchases, Web visits, Campaign acceptance rates

---

## 🔬 Methodology & Workflow
1. **Exploratory Data Analysis (EDA)**:
   - Data cleaning, handling missing values, and outlier inspection.
   - Distribution analysis and correlation analysis across spending and demographic metrics.
2. **Feature Engineering & Preprocessing**:
   - Creating derived features (e.g., Total Spending, Total Children, Age, Customer Tenure).
   - Standard scaling (`StandardScaler`) for distance-based clustering algorithms.
3. **Dimensionality Reduction**:
   - Applying **PCA** to reduce high-dimensional feature space while preserving maximum variance.
4. **Clustering Models**:
   - **K-Means Clustering**: Finding optimal $K$ using the Elbow Method and Silhouette Analysis.
   - **Hierarchical / Agglomerative Clustering**: Constructing Dendrograms using Ward's linkage.
   - **DBSCAN**: Density-based clustering for identifying core clusters and noise/outliers.
5. **Cluster Profiling & Evaluation**:
   - Evaluating cluster separation using **Silhouette Score**, **Davies-Bouldin**, and **Calinski-Harabasz**.
   - Interpreting segments for targeted marketing and business strategy.

---

## 🚀 How to Run Locally

### 1. Clone the repository
```bash
git clone https://github.com/Vandana251/Customer-Segmentation.git
cd Customer-Segmentation
```

### 2. Install Dependencies
```bash
pip install numpy pandas matplotlib seaborn scikit-learn scipy jupyter
```

### 3. Open the Notebook
```bash
jupyter notebook Customer_Segmentation_Unsupervised_Vandana.ipynb
```

---

## 👤 Author
- **Vandana Illipilla**
- GitHub: [@Vandana251](https://github.com/Vandana251)
