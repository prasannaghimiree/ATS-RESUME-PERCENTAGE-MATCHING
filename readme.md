# ATS Resume Percentage Matcher

## Overview
The **ATS Resume Percentage Matcher** is a tool designed to analyze resumes against job descriptions and provide a percentage match. It simulates an Applicant Tracking System (ATS) to help job seekers and recruiters assess the relevance of a resume for a specific job role.

## Features
- **Text Parsing**: Extracts text from resumes and job descriptions.
- **Keyword Matching**: Compares keywords and key phrases between the resume and job description.
- **Similarity Scoring**: Uses NLP techniques to compute a match percentage.
- **Stopword Removal**: Filters out common words that do not contribute to relevance.
- **Weighting Mechanism**: Assigns weights to different sections of the resume for improved accuracy.
- **User-Friendly Output**: Provides a clear percentage score indicating match strength.

## Technologies Used
- **Python** (Core Logic)
- **NLTK / SpaCy** (Natural Language Processing)
- **TF-IDF / Word2Vec** (Feature Extraction)
- **Flask / FastAPI** (Optional Web Interface)
- **Pandas / NumPy** (Data Handling)

## Installation
```bash
# Clone the repository
git clone https://github.com/your-repo/ATS-Resume-Matcher.git
cd ATS-Resume-Matcher

# Create a virtual environment
python -m venv venv
source venv/bin/activate  # On Windows use: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
