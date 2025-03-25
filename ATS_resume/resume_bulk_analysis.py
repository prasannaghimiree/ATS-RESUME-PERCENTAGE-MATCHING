import pandas as pd
import os
from similarity import calculate_similarity
from stability import calculate_stability
from utils import read_resume_file

def analyze_resumes(input_excel, output_excel):
    """Reads an Excel file, processes each resume, and writes results back."""
    df = pd.read_excel(input_excel)

    # Ensure necessary columns exist
    required_columns = {'Name', 'Job_Description', 'Resume_Path'}
    if not required_columns.issubset(df.columns):
        raise ValueError(f"Excel must have these columns: {required_columns}")

    # Initialize lists to store scores
    match_percentages = []
    stability_scores = []

    for index, row in df.iterrows():
        job_desc = str(row['Job_Description'])
        resume_path = str(row['Resume_Path']).strip()

        if not os.path.exists(resume_path):
            print(f"❌ Warning: Resume not found at {resume_path}")
            match_percentages.append(None)
            stability_scores.append(None)
            continue

        resume_text = read_resume_file(resume_path)

        # Calculate scores
        match_percentage = calculate_similarity(job_desc, resume_text)
        stability_score = calculate_stability(resume_text)

        print(f"✅ {row['Name']} - Match: {match_percentage}%, Stability: {stability_score}%")

        match_percentages.append(match_percentage)
        stability_scores.append(stability_score)

    # Append new columns
    df["Match_Percentage"] = match_percentages
    df["Stability_Score"] = stability_scores

    # Save results
    df.to_excel(output_excel, index=False)
    print(f"\n📊 Analysis complete! Results saved to {output_excel}")
