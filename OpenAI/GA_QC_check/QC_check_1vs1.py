import pandas as pd

# === STEP 1: Load Excel file ===
file_path = r"C:\Users\76000\OneDrive - Bain\Desktop\Python\OpenAI\GA_QC_check\Binary Data 11.20 PM.xlsx"
df = pd.read_excel(file_path, dtype=str)  # read everything as string so blanks are visible as NaN or empty

# === STEP 2: Identify where 'Q1_EmploymentStatus' column starts ===
start_col = df.columns.get_loc("Q1_EmploymentStatus")

# Keep all columns from that point onward
responses = df.iloc[:, start_col:]

# === STEP 3: Clean blanks consistently ===
# Convert NaN, "nan", and None to empty string for consistent comparison
responses = responses.fillna("").astype(str).replace("nan", "")

# === STEP 4: Define comparison function ===
def similarity_score(row1, row2):
    comparison = row1 == row2
    matches = comparison.sum()
    total = len(row1)
    score = matches / total if total > 0 else 0
    return comparison, score, matches, total

# === STEP 5: Compare Record 1 (row 2) with Record 2 (row 3) ===
row1 = responses.iloc[0]   # Respondent 1
row2 = responses.iloc[1]   # Respondent 2

comparison, score, matches, total = similarity_score(row1, row2)

print(f"\n✅ Record 1 vs Record 2:")
print(f"Matching responses (including blanks): {matches} / {total}")
print(f"Similarity score: {score*100:.2f}%")

# === STEP 6: Save detailed results for manual check ===
comparison_df = pd.DataFrame({
    "Question": responses.columns,
    "Respondent_1": row1.values,
    "Respondent_2": row2.values,
    "Match": comparison.values
})

comparison_df.to_excel("Record1_vs_Record2_BlankAwareComparison.xlsx", index=False)
print("\n📄 Saved as 'Record1_vs_Record2_BlankAwareComparison.xlsx'")
