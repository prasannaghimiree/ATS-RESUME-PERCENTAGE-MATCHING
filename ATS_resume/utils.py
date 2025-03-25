import os
import docx

def read_text_file(file_path):
    """Reads and returns the text content of a .txt file."""
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            return file.read()
    except Exception as e:
        print(f"❌ Error reading {file_path}: {e}")
        return ""

def read_docx_file(file_path):
    """Reads and returns text from a .docx file."""
    try:
        doc = docx.Document(file_path)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        print(f"❌ Error reading {file_path}: {e}")
        return ""

def read_resume_file(file_path):
    """Reads resume content based on file type (.txt or .docx)."""
    _, ext = os.path.splitext(file_path.lower())

    if ext == ".txt":
        return read_text_file(file_path)
    elif ext == ".docx":
        return read_docx_file(file_path)
    else:
        print(f"❌ Unsupported file format: {ext}")
        return ""
    
# def read_resume_file(file_path):
#     _, ext = os.path.splitext(file_path.lower())

#     if ext == ".txt":
#         return read_text_file(file_path)
#     elif ext==".docx":
#         return read_docx_file(file_path)
#     else:
#         print(f"Unsupported file format: {ext}")
#         return ""
