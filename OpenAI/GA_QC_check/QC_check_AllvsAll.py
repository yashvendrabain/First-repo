import pandas as pd
import numpy as np
from itertools import combinations


file_path = r"C:\Users\76000\OneDrive - Bain\Desktop\Python\OpenAI\GA_QC_check\Binary Data (2).xlsx"
df = pd.read_excel(file_path)


start_col = df.columns.get_loc("Q1_EmploymentStatus")
responses = df.iloc[:, start_col:]


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


results = []
for (i, row1), (j, row2) in combinations(responses.iterrows(), 2):
    score, matches, total = similarity_score(row1, row2)
    results.append({
        "Respondent_1": i + 2, 
        "Respondent_2": j + 2,
        "Matching_Responses": matches,
        "Total_Questions": total,
        "Similarity_Score (%)": round(score * 100, 2)
    })


similarity_df = pd.DataFrame(results)
output_file = "AllRespondents_vs_All_Other_Comparison.xlsx"
similarity_df.to_excel(output_file, index=False)

print("\nCompleted: Every Respondent vs Every Other Respondent")
print(similarity_df.head())
print(f"\n📄 Saved as '{output_file}'")
