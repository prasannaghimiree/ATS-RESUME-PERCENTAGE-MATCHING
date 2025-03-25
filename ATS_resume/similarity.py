from gensim.models.doc2vec import Doc2Vec, TaggedDocument
from gensim.utils import simple_preprocess

def train_doc2vec_model(job_description, resumes):
    """Train a Doc2Vec model using job descriptions and resumes."""
    documents = [TaggedDocument(simple_preprocess(job_description), ["JD"])]
    for i, resume in enumerate(resumes):
        documents.append(TaggedDocument(simple_preprocess(resume), [f"Resume_{i}"]))

    model = Doc2Vec(vector_size=100, window=5, min_count=1, workers=4, epochs=20)
    model.build_vocab(documents)
    model.train(documents, total_examples=model.corpus_count, epochs=model.epochs)
    
    return model

def calculate_similarity(job_description, resume_text):
    """Compute similarity score between job description and a resume using Doc2Vec."""
    model = train_doc2vec_model(job_description, [resume_text])

    job_vector = model.infer_vector(simple_preprocess(job_description))
    resume_vector = model.infer_vector(simple_preprocess(resume_text))
    
    similarity_score = model.wv.cosine_similarities(job_vector, [resume_vector])[0]
    
    return round(similarity_score * 100)  

# def similarity(job_description, resume_text):
#     model = train_doc2vec_model(job_description, [resume_text])
#     job_vector = model.infer_vector(simple_preprocess(resume_text))
#     resume_vector = model.infer_vector(simple_preprocess(resume_text))
#     similarity_score = model.wv.cosine_similarities(job_vector,[resume_vector])[0]

#     return round(similarity_score*100)
