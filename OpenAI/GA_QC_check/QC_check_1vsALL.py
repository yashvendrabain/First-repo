import pandas as pd
import numpy as np

# === STEP 1: Load Excel file (keep native datatypes) ===
file_path = r"C:\Users\76000\OneDrive - Bain\Desktop\Python\OpenAI\GA_QC_check\Binary Data 11.20 PM.xlsx"
df = pd.read_excel(file_path)

# === STEP 2: Start from 'Q1_EmploymentStatus' onward ===
start_col = df.columns.get_loc("Q1_EmploymentStatus")
responses = df.iloc[:, start_col:]

# === STEP 3: Define comparison function (blank = blank → True) ===
def similarity_score(row1, row2):
    comparison = []
    for a, b in zip(row1, row2):
        if (pd.isna(a) or a == "") and (pd.isna(b) or b == ""):
            comparison.append(True)
        else:
            comparison.append(a == b)
    matches = sum(comparison)
    total = len(comparison)
    score = matches / total if total > 0 else 0
    return score, matches, total

# === STEP 4: Fix respondent X = 1 (row 2 in Excel) ===
row_x = 0  # Respondent 1
target_row = responses.iloc[row_x]

# === STEP 5: Compare Respondent 1 with every other respondent ===
results = []
for i, row in responses.iterrows():
    if i != row_x:
        score, matches, total = similarity_score(target_row, row)
        results.append({
            "Respondent_X": row_x + 2,
            "Compared_with_Row": i + 2,
            "Matching_Responses": matches,
            "Total_Questions": total,
            "Similarity_Score (%)": round(score * 100, 2)
        })

# === STEP 6: Output results ===
similarity_df = pd.DataFrame(results)
print(f"\n✅ Respondent 1 vs All Other Respondents (Blank = Blank → True):\n")
print(similarity_df)

output_file = "Respondent_1_vs_All_NativeComparison.xlsx"
similarity_df.to_excel(output_file, index=False)
print(f"\n📄 Saved as '{output_file}'")
