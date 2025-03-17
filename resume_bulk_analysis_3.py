import pandas as pd
import time
import random
from ats_func_3 import analyze_resume

def process_resumes(input_file, output_file):
    df = pd.read_excel(input_file)
    results = []
    
    for index, row in df.iterrows():
        try:
            
            delay = min((2 ** index) + random.uniform(0, 1), 60)
            time.sleep(delay)
            
            result = analyze_resume(row['Resume'], row['JobDescription'])
            results.append({
                "Applicant": row['Applicant'],
                "Position": row['Position'],
                "Match_Percentage": result["Overall_Match"],
                "Stability_Score": result["Stability_Score"],
                "Total_Experience": result["Total_Experience"],
                "Companies_Count": result["Companies_Count"]
            })
            print(f"Processed: {row['Applicant']}")
        except Exception as e:
            print(f"Error processing {row['Applicant']}: {str(e)}")
            results.append({
                "Applicant": row['Applicant'],
                "Position": row['Position'],
                "Match_Percentage": "Error",
                "Stability_Score": "Error",
                "Total_Experience": "Error",
                "Companies_Count": "Error"
            })
    
    pd.DataFrame(results).to_excel(output_file, index=False)
    print(f"Analysis complete. Results saved to {output_file}")

if __name__ == "__main__":
    process_resumes(
        input_file="cvs.xlsx",
        output_file="ats_results.xlsx"
    )