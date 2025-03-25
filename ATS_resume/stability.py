import re

def extract_experience_details(resume_text):
    """Extract number of companies and total experience from resume."""
    companies = re.findall(r"\b(?:Company|Employer|Organization)\s*:\s*(.+)", resume_text, re.IGNORECASE)
    experience = re.findall(r"(\d+)\s*years?\s*experience", resume_text, re.IGNORECASE)

    total_experience = sum(map(int, experience)) if experience else 0
    num_companies = len(set(companies)) if companies else 1

    return total_experience, num_companies

def calculate_stability(resume_text):
    """Calculate stability score based on job switches."""
    total_experience, num_companies = extract_experience_details(resume_text)
    
    if num_companies == 0:  # Edge case handling
        return 100
    
    stability_score = min((total_experience / num_companies) * 25, 100)  # Cap at 100%
    
    return round(stability_score)
