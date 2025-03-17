import pandas as pd
import pdfplumber
import google.generativeai as genai
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os
import re
import json

genai.configure(api_key=os.environ['GOOGLE_API_KEY'])
model = genai.GenerativeModel('gemini-1.5-flash')

def extract_resume_data(resume_text):
    prompt = f"""**Resume Analysis Task**
    
    Return ONLY VALID JSON with:
    {{
        "education": [{{"field": "Field of Study", "institution": "Institution Name"}}],
        "skills": ["skill1", "skill2"],
        "experience": [
            {{
                "job_title": "Position",
                "company": "Company Name",
                "start_date": "MM/YYYY",
                "end_date": "MM/YYYY or 'Present'",
                "is_relevant": boolean
            }}
        ]
    }}
    
    Rules:
    - Current date: {datetime.now().strftime('%m/%Y')}
    - Exclude freelance/volunteer roles
    - Format dates as MM/YYYY
    - is_relevant: true only for professional IT roles
    
    Resume:
    {resume_text}
    """
    
    try:
        response = model.generate_content(prompt)
        json_str = re.search(r'\{.*\}', response.text, re.DOTALL).group()
        return json.loads(json_str)
    except Exception as e:
        print(f"Error parsing resume: {str(e)}")
        return {"education": [], "skills": [], "experience": []}

def parse_job_description(job_desc):
    prompt = f"""**Job Description Analysis Task**
    
    Return ONLY VALID JSON with:
    {{
        "required_education": ["Degree 1", "Degree 2"],
        "required_skills": ["skill1", "skill2"],
        "min_experience": years,
        "relevant_titles": ["Title 1", "Title 2"]
    }}
    
    Job Description:
    {job_desc}
    """
    
    try:
        response = model.generate_content(prompt)
        json_str = re.search(r'\{.*\}', response.text, re.DOTALL).group()
        return json.loads(json_str)
    except Exception as e:
        print(f"Error parsing job description: {str(e)}")
        return {"required_education": [], "required_skills": [], "min_experience": 0}

def parse_job_description(job_desc):
    prompt = f"""**Job Description Analysis Task**
    
    Extract from this job description as JSON:
    1. required_education: List of required degrees/qualifications
    2. required_skills: List of required technical skills (normalized)
    3. min_experience: Minimum years of experience required (extract number)
    4. relevant_titles: List of relevant job titles/positions
    
    Job Description:
    {job_desc}
    """
    
    response = model.generate_content(prompt)
    return eval(response.text)

def calculate_experience(periods):
    if not periods:
        return 0.0
    
    sorted_periods = sorted(periods, key=lambda x: x[0])
    merged = [sorted_periods[0]]
    
    for current_start, current_end in sorted_periods[1:]:
        last_start, last_end = merged[-1]
        
        if current_start <= last_end:
            merged[-1] = (last_start, max(last_end, current_end))
        else:
            merged.append((current_start, current_end))
    
    total_months = 0
    for start, end in merged:
        delta = relativedelta(end, start)
        total_months += delta.years * 12 + delta.months
        
    return round(total_months / 12, 2)

def calculate_match_score(job_reqs, resume_data):
    edu_intersection = set(resume_data['education']) & set(job_reqs['required_education'])
    edu_score = len(edu_intersection) / len(job_reqs['required_education']) if job_reqs['required_education'] else 0
    
    skill_intersection = set(resume_data['skills']) & set(job_reqs['required_skills'])
    skill_score = len(skill_intersection) / len(job_reqs['required_skills']) if job_reqs['required_skills'] else 0
    
    exp_score = min(resume_data['total_experience'] / job_reqs['min_experience'], 1.0) if job_reqs['min_experience'] > 0 else 1.0
    
    return (0.3 * edu_score + 0.5 * skill_score + 0.2 * exp_score) * 100

def calculate_stability(experiences):
    if not experiences:
        return 0.0
    
    total_points = 0
    for exp in experiences:
        if not exp['is_relevant']:
            continue
            
        start = datetime.strptime(exp['start_date'], '%m/%Y')
        end = datetime.strptime(exp['end_date'], '%m/%Y') if exp['end_date'].lower() != 'present' else datetime.now()
        
        years = (end - start).days / 365.25
        
        if years > 3:
            total_points += 1.5
        elif years >= 2:
            total_points += 1.0
        elif years >= 0.5:
            total_points += 0.5
        else:
            total_points += 0.25
    
    return min((total_points / len(experiences)) * 100, 100)

def process_resume(job_desc, resume_path):
    with pdfplumber.open(resume_path) as pdf:
        resume_text = "\n".join(page.extract_text() for page in pdf.pages)
    
    resume_data = extract_resume_data(resume_text)
    job_reqs = parse_job_description(job_desc)
    
    periods = []
    for exp in resume_data['experience']:
        if exp['is_relevant']:
            start = datetime.strptime(exp['start_date'], '%m/%Y')
            end = datetime.strptime(exp['end_date'], '%m/%Y') if exp['end_date'].lower() != 'present' else datetime.now()
            periods.append((start, end))
    
    resume_data['total_experience'] = calculate_experience(periods)
    
    match_score = calculate_match_score(job_reqs, resume_data)
    stability_score = calculate_stability(resume_data['experience'])
    
    return round(match_score, 2), round(stability_score, 2)

def main(input_file, output_file):
    df = pd.read_excel(input_file)
    results = []
    
    for _, row in df.iterrows():
        try:
            match, stability = process_resume(row['JobDescription'], row['Resume'])
            results.append({
                'Applicant': row['Applicant'],
                'Position': row['Position'],
                'Match%': match,
                'Stability%': stability
            })
        except Exception as e:
            print(f"Error processing {row['Applicant']}: {str(e)}")
    
    pd.DataFrame(results).to_excel(output_file, index=False)

if __name__ == "__main__":
    main('cvs.xlsx', 'output.xlsx')