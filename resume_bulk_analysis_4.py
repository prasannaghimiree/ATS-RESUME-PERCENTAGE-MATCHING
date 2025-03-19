import pandas as pd
import cx_Oracle
import os
import time
import random
import json
from ats_func_4 import analyze_resume
from dotenv import load_dotenv

load_dotenv()


def get_db_connection():
    """Establish a new database connection."""
    DB_USER = os.getenv("DB_USER")
    DB_PASS = os.getenv("DB_PASS")
    HOST = os.getenv("HOST")
    PORT = os.getenv("PORT")
    SERVICE = os.getenv("SERVICE")

    dsn_tns = cx_Oracle.makedsn(HOST, PORT, service_name=SERVICE)
    return cx_Oracle.connect(user=DB_USER, password=DB_PASS, dsn=dsn_tns)


def update_database(id_value, overall_match):
    """Update the database with match percentage and flag using a new connection each time."""
    try:
        with get_db_connection() as connection:
            with connection.cursor() as cursor:
                write_query = f"""
                UPDATE RESUME_DETAILS  
                SET PERCENTAGE_MATCH = :1, FLAG = 'Y'  
                WHERE ID = :2 AND FLAG = 'N'
                """
                cursor.execute(write_query, (overall_match, id_value))
                connection.commit()
    except Exception as e:
        print(f"Database update failed for ID {id_value}: {str(e)}")


def process_resumes(input_file, output_file):
    """Processes resumes, analyzes them, and updates the database."""
    df = pd.read_excel(input_file)
    results = []

    for index, row in df.iterrows():
        try:

            delay = min((2**index) + random.uniform(0, 1), 60)
            time.sleep(delay)

            result = analyze_resume(row["RESUME"], row["JOBDESCRIPTION"])

            results.append(
                {
                    "ID": row["ID"],
                    "APPLICANT": row["APPLICANT"],
                    "POSITION": row["POSITION"],
                    "Match_Percentage": result["Overall_Match"],
                    "Stability_Score": result["Stability_Score"],
                    "Total_Experience": result["Total_Experience"],
                    "Companies_Count": result["Companies_Count"],
                    "Strengths": ", ".join(result["Strengths"]),
                    "Weaknesses": ", ".join(result["Weaknesses"]),
                    "Score_Breakdown": json.dumps(result["Score_Breakdown"]),
                    "Detailed_Analysis": result["Detailed_Analysis"],
                }
            )

            print(f"Processed: {row['APPLICANT']}")

            update_database(row["ID"], result["Overall_Match"])

        except Exception as e:
            print(f"Error processing {row['APPLICANT']}: {str(e)}")
            results.append(
                {
                    "ID": row.get("ID", "Error"),
                    "APPLICANT": row["APPLICANT"],
                    "POSITION": row["POSITION"],
                    "Match_Percentage": "Error",
                    "Stability_Score": "Error",
                    "Total_Experience": "Error",
                    "Companies_Count": "Error",
                    "Strengths": "Error",
                    "Weaknesses": "Error",
                    "Score_Breakdown": "Error",
                    "Detailed_Analysis": "Error",
                }
            )

    pd.DataFrame(results).to_excel(output_file, index=False)
    print(f"Analysis complete. Results saved to {output_file}")


if __name__ == "__main__":

    with get_db_connection() as connection:
        query = "SELECT * FROM RESUME_DETAILS WHERE FLAG='N'"
        df = pd.read_sql(query, con=connection)
        df.to_excel("data_extracted_from_database.xlsx", index=False)

    process_resumes(
        input_file="data_extracted_from_database.xlsx",
        output_file="new_result_from_extracted_data_5.xlsx",
    )
