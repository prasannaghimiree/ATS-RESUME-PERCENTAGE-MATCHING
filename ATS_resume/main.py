import os
from resume_bulk_analysis import analyze_resumes

# Paths
INPUT_EXCEL = "cvs.xlsx" 
OUTPUT_EXCEL = "resume_analysis_results.xlsx"

if __name__ == "__main__":
    analyze_resumes(INPUT_EXCEL, OUTPUT_EXCEL)
