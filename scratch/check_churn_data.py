import pandas as pd
import glob
import os

files = glob.glob("**/*Churn*.xlsx", recursive=True)
print("Churn files:", files)

for f in files:
    try:
        df = pd.read_excel(f)
        print(f"\nFile: {f}, shape: {df.shape}")
        print("Columns:", df.columns.tolist()[:10])
        if "Churn" in df.columns:
            churn_yes = df[df["Churn"] == "Yes"]
            print(f"Total customers: {len(df)}, Churned: {len(churn_yes)} ({len(churn_yes)/len(df):.2%})")
            if "MonthlyCharges" in df.columns:
                mrr_total = df["MonthlyCharges"].sum()
                mrr_churn = churn_yes["MonthlyCharges"].sum()
                print(f"Total MonthlyCharges (MRR): ${mrr_total:,.2f}")
                print(f"Churned MonthlyCharges (At-Risk MRR): ${mrr_churn:,.2f}")
                print(f"Annualized At-Risk ARR (MRR * 12): ${mrr_churn * 12:,.2f}")
            if "TotalCharges" in df.columns:
                # convert to numeric
                tc = pd.to_numeric(df["TotalCharges"], errors="coerce")
                tc_churn = pd.to_numeric(churn_yes["TotalCharges"], errors="coerce").sum()
                print(f"Total Lifetime Charges: ${tc.sum():,.2f}")
                print(f"Total Lifetime Charges of Churned: ${tc_churn:,.2f}")
    except Exception as e:
        print(f"Error reading {f}: {e}")
