from docx.docx import Document


def get_comments(document_list):
    for document in document_list:
        doc = Document(document)
        comments = doc.comments
        yield comments
