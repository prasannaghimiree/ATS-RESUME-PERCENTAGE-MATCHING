import pandas as pd
from flask import Flask, jsonify
from resume_bulk_analysis import process_resumes, get_db_connection

app = Flask(__name__)

@app.route("/process_resume", methods=["Get"])
def process_resumes_endpoint():

    try:
        with get_db_connection() as connection:
            query = "SELECT * FROM RESUME_DETAILS WHERE FLAG='N'"
            # query = "SELECT * FROM RESUME_DETAILS"
            df = pd.read_sql(query, con=connection)
            df.to_excel("data_extracted_from_database.xlsx", index=False)
        process_resumes(
            input_file="data_extracted_from_database.xlsx",
            output_file="new_result_from_extracted_data.xlsx",
        )
        df = pd.read_excel("new_result_from_extracted_data.xlsx")
        results = df[["ID", "APPLICANT", "Match_Percentage"]].to_dict(orient="records")

        return jsonify({"status": "completed", "results": results})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500
    
    
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
