import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats

sns.set(style="whitegrid")

df = pd.read_csv(r"C:\Users\ASTHA BAHETI\Downloads\sales_dataset.csv")


print("First 10 records")
print(df.head(10))


print("\nDataset info")
print(df.info())


print("Descriptive statistics")
print(df.describe(include="all"))


print("Missing values per column")
print(df.isnull().sum())


print("\nDuplicate rows:", df.duplicated().sum())


df["Order_Date"] = pd.to_datetime(df["Order_Date"])
print("\nOrder_Date dtype after conversion:", df["Order_Date"].dtype)


numerical_vars = df.select_dtypes(include=np.number).columns.tolist()
categorical_vars = df.select_dtypes(include="object").columns.tolist()
print("\nNumerical variables:", numerical_vars)
print("Categorical variables:", categorical_vars)