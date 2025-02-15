# -*- coding: utf-8 -*-
"""
Created on Sat Feb 15 11:19:49 2025

@author: Asus
"""
import pandas as pd

# Define file paths for urban and rural data
path_r = r"C:\Users\Asus\Desktop\ترم2اقتصاد\AppliedEconometrics\EX1\SumR1401.xlsx"
path_u = r"C:\Users\Asus\Desktop\ترم2اقتصاد\AppliedEconometrics\EX1\SumU1401.xlsx"

# Load Excel files separately
df_r = pd.read_excel(path_r)
df_u = pd.read_excel(path_u)

# Function to process each dataset separately (urban or rural)
def process_dataset(df, dataset_type):
    # Convert all column names to uppercase to standardize naming
    df.columns = df.columns.str.upper()

    # Extract province codes (2nd and 3rd digits from left in ADDRESS column)
    df["PROVINCE CODE"] = df["ADDRESS"].astype(str).str[1:3]

    # Dictionary for province codes
    province_codes = {
        "00": "مرکزی", "01": "گیلان", "02": "مازندران", "03": "آذربایجان‌شرقی", "04": "آذربایجان‌غربی",
        "05": "کرمانشاه", "06": "خوزستان", "07": "فارس", "08": "کرمان", "09": "خراسان‌رضوی",
        "10": "اصفهان", "11": "سیستان و بلوچستان", "12": "کردستان", "13": "همدان", "14": "چهارمحال و بختیاری",
        "15": "لرستان", "16": "ایلام", "17": "کهگیلویه و بویراحمد", "18": "بوشهر", "19": "زنجان",
        "20": "سمنان", "21": "یزد", "22": "هرمزگان", "23": "تهران", "24": "اردبیل",
        "25": "قم", "26": "قزوین", "27": "گلستان", "28": "خراسان‌شمالی", "29": "خراسان‌جنوبی", "30": "البرز"
    }

    df["PROVINCE"] = df["PROVINCE CODE"].map(province_codes)

    # Calculate total household income (only using existing columns that contain "DARAMAD")
    income_columns = df.filter(like="DARAMAD").columns
    df["TOTAL HOUSEHOLD INCOME"] = df[income_columns].sum(axis=1, min_count=1)

    # Ensure numeric columns are properly formatted
    df["TOTAL HOUSEHOLD INCOME"] = pd.to_numeric(df["TOTAL HOUSEHOLD INCOME"], errors='coerce')
    df["EDUCATION LEVEL"] = pd.to_numeric(df["A05"], errors='coerce')  # Keep education level as numeric
    df["HOUSEHOLD SIZE"] = pd.to_numeric(df["C01"], errors='coerce')

    # Group by province and calculate all summary statistics
    summary = df.groupby("PROVINCE").agg(
        MIN_INCOME=("TOTAL HOUSEHOLD INCOME", "min"),
        MAX_INCOME=("TOTAL HOUSEHOLD INCOME", "max"),
        MEAN_INCOME=("TOTAL HOUSEHOLD INCOME", "mean"),
        STD_INCOME=("TOTAL HOUSEHOLD INCOME", "std"),
        MIN_EDU=("EDUCATION LEVEL", "min"),
        MAX_EDU=("EDUCATION LEVEL", "max"),
        MEAN_EDU=("EDUCATION LEVEL", "mean"),
        STD_EDU=("EDUCATION LEVEL", "std"),
        MIN_HH_SIZE=("HOUSEHOLD SIZE", "min"),
        MAX_HH_SIZE=("HOUSEHOLD SIZE", "max"),
        MEAN_HH_SIZE=("HOUSEHOLD SIZE", "mean"),
        STD_HH_SIZE=("HOUSEHOLD SIZE", "std")
    ).reset_index()

    # Save output to separate Excel files
    output_filename = f"{dataset_type}_Province_Summary.xlsx"
    summary.to_excel(output_filename, index=False)
    print(f"✅ Summary saved as '{output_filename}'")

# Process urban and rural datasets separately
process_dataset(df_u, "Urban")  # برای داده‌های شهری
process_dataset(df_r, "Rural")  # برای داده‌های روستایی


