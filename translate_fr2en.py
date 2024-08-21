from tqdm import tqdm
from rich import print as rprint

from docx import Document
from deep_translator import GoogleTranslator

# Load the document
doc_path = "data/document_in_french.docx"
doc = Document(doc_path)

# Initialize the translator
translator = GoogleTranslator(source='fr', target='en')

# Translate paragraphs while keeping the formatting
for paragraph in tqdm(doc.paragraphs):
    if paragraph.text.strip():  # If the paragraph has text
        translation = translator.translate(paragraph.text)
        paragraph.text = translation

# Save the translated document
translated_doc_path = f"{doc_path.split('.')[0]}_translated.docx"
doc.save(translated_doc_path)

rprint(f"Translated document saved at '{translated_doc_path}'")
