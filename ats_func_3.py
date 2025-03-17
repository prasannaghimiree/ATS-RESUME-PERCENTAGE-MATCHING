import json
import google.generativeai as genai
import os
import pdfplumber
import re
import time
from dotenv import load_dotenv
from datetime import datetime
from dateutil.relativedelta import relativedelta

load_dotenv()

genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

MODEL_CONFIG = {
    "temperature": 0.1,
    "max_output_tokens": 2000,
    "response_mime_type": "application/json",
}

safety_settings = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_ONLY_HIGH"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_ONLY_HIGH"},
]

def safe_json_parse(response_text):
    """Robust JSON parsing with multiple fallback strategies"""
    try:
        
        return json.loads(response_text)
    except json.JSONDecodeError:
        try:
           
            json_match = re.search(r'```json\n(.*?)\n```', response_text, re.DOTALL)
            if json_match:
                return json.loads(json_match.group(1))
        except:
            try:
                json_str = re.search(r'\{.*\}', response_text, re.DOTALL)
                if json_str:
                    return json.loads(json_str.group())
            except:
                return {}
    return {}

def extract_structured_data(text):
    """Extract resume data with validation and retries"""
    prompt = f"""
    Extract resume data as VALID JSON with:
    Work experience (exclude freelancing/voluntary/college/Teaching/Intern/Internship/fellowship/Instructing/Projects work and out of context/field experience). For each position, label it as false in relevant.
    Terminologies similar to skills which are mention in job experience, projects are also contedas a skills.
    {{
        "skills": ["Python", "Machine Learning", ...],
        "education": ["Bachelor's in Computer Science", ...],
        "experience": [
            {{
                "company": "Company Name",
                "start": "MM/YYYY",
                "end": "MM/YYYY/Present",
                "position": "Job Title",
                "relevant": true/false
            }}
        ]
    
    }}
    Resume Text (truncated):
    {text[:10000]}
    """
    
    model = genai.GenerativeModel('gemini-1.5-flash')
    for attempt in range(3):
        try:
            response = model.generate_content(prompt)
            if response.text:
                data = safe_json_parse(response.text)
                
                if all(key in data for key in ['skills', 'education', 'experience']):
                    return data
        except Exception as e:
            print(f"Extraction error (attempt {attempt+1}): {str(e)[:50]}")
            time.sleep(2 ** attempt)
    return {"skills": [], "education": [], "experience": []}

    

def parse_date(date_str):
    """Robust date parsing with multiple formats"""
    date_str = str(date_str).strip()
    if not date_str:
        return datetime.now()
    
    try:
        if date_str.lower() == "present":
            return datetime.now()
        
        
        for fmt in ("%m/%Y", "%Y/%m", "%b %Y", "%B %Y", "%Y"):
            try:
                return datetime.strptime(date_str, fmt)
            except:
                continue
    except:
        pass
    
    return datetime.now()

def calculate_experience(experiences):
    """Calculate total relevant experience with merged date ranges"""
    try:
        intervals = []
        for exp in experiences:
            if not exp.get('relevant', False):
                continue
            
            start = parse_date(exp.get('start', '01/2000'))
            end = parse_date(exp.get('end', 'Present'))
            intervals.append((start, end))
        
        if not intervals:
            return 0.0
        
        sorted_intervals = sorted(intervals, key=lambda x: x[0])
        merged = [sorted_intervals[0]]
        
        for current in sorted_intervals[1:]:
            last = merged[-1]
            if current[0] <= last[1]:
                merged[-1] = (min(current[0], last[0]), max(current[1], last[1]))
            else:
                merged.append(current)
        
       
        total_months = 0
        for start, end in merged:
            delta = relativedelta(end, start)
            total_months += delta.years * 12 + delta.months
        
        return (total_months / 12)
    except Exception as e:
        print(f"Experience calculation error: {str(e)[:50]}")
        return 0.0

def calculate_stability(experiences):
    """Calculate stability score based on tenure points"""
    try:
        points = []
        for exp in experiences:
            if not exp.get('relevant', False):
                continue
            
            start = parse_date(exp.get('start', '01/2000'))
            end = parse_date(exp.get('end', 'Present'))
            months = relativedelta(end, start).years * 12 + relativedelta(end, start).months
            
            if months < 6:
                points.append(0.25)
            elif 6 <= months < 12:
                points.append(0.5)
            elif 12 <= months < 24:
                points.append(1)
            elif 24 <= months < 36:
                points.append(1.5)
            else:
                points.append(2)
        
        if not points:
            return 0
        
        avg_score = sum(points) / len(points)
        return min(avg_score * 100, 100)
    except Exception as e:
        print(f"Stability calculation error: {str(e)[:50]}")
        return 0

def get_match_percentage(job_desc, resume_text):
    """Calculate match percentage with validation"""
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    prompt = f"""
    For example: Return EXACTLY like this JSON format:
    {{
        "match": maching_percentage,
        "stability": stability_percentage
    }}
    
    Calculation Rules:
    - Match%: Skills(40%) + Education(20%) + Experience(40%)
    - Stability%: Tenure points converted to percentage. 
    
    Job Description:
    {job_desc[:10000]}
    
    Resume Content:
    {resume_text[:10000]}
    """
    
    for attempt in range(5):
        try:
            response = model.generate_content(prompt)
            if response.text:
                data = safe_json_parse(response.text)
                if "match" in data and "stability" in data:
                    return {
                        "match": max(0, min(100, int(data["match"]))),
                        "stability": max(0, min(100, int(data["stability"])))
                    }
        except Exception as e:
            sleep_time = min(2 ** attempt + 5, 60)  
            print(f"API Error (attempt {attempt+1}): Sleeping {sleep_time}s")
            time.sleep(sleep_time)
    
    return {"match": 0, "stability": 0}

def analyze_resume(file_path, job_description):
    """Main analysis function with error handling"""
    try:
        with pdfplumber.open(file_path) as pdf:
            text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
        
        resume_data = extract_structured_data(text)
        print("************************************************************")
        print(resume_data)
        print("************************************************************")
        print("##############################################################")
        print(job_description)
        print("##############################################################")
        scores = get_match_percentage(job_description, text)
        
        
        return {
            "Overall_Match": scores["match"],
            "Stability_Score": scores["stability"],
            "Total_Experience": calculate_experience(resume_data.get("experience", [])),
            "Companies_Count": len([e for e in resume_data.get("experience", []) if e.get("relevant", False)])
        }
    
      
    except Exception as e:
        print(f"Analysis failed: {str(e)[:50]}")
        return {
            "Overall_Match": 0,
            "Stability_Score": 0,
            "Total_Experience": 0.0,
            "Companies_Count": 0
        }

