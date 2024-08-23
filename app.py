from flask import Flask, request, render_template, send_file
from docx import Document
import os
import re

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'

# Ensure the upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    doc_file = request.files['doc']
    template_file = request.files['template']

    doc_path = os.path.join(app.config['UPLOAD_FOLDER'], doc_file.filename)
    template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_file.filename)
    
    doc_file.save(doc_path)
    template_file.save(template_path)
    
    # Process the document
    styled_doc_path = process_document(doc_path, template_path)
    
    return send_file(styled_doc_path, as_attachment=True)

def process_document(doc_path, template_path):
    # Load the document
    doc = Document(doc_path)
    # Extract the book ID from the file name
    book_id = os.path.basename(doc_path).split('_')[0]
    # Extract the chapter number from the first paragraph
    chapter_number = doc.paragraphs[0].text.strip()
    # Initialize a counter for figure numbers
    figure_counter = 1
    # Regular expression to match figure captions
    figure_caption_pattern = re.compile(rf"^Figure {chapter_number}\.\d+:")
    # Regular expression to identify code in text
    code_term_pattern = re.compile(r'\b\w+\(\)|\b[A-Za-z][a-z]+[A-Z]\w*\b|\b\w+_\w+\b')

    # Define the heading styles from the template
    heading_styles = {
        'Heading 1': 'HS - Heading_1 [PACKT]',
        'Heading 2': 'HS - Heading_2 [PACKT]',
        'Heading 3': 'HS - Heading_3 [PACKT]',
        'Heading 4': 'HS - Heading_4 [PACKT]',
        'Heading 5': 'HS - Heading_5 [PACKT]',
        'Heading 6': 'HS - Heading_6 [PACKT]',
    }

    # Iterate through paragraphs and apply styles
    for para in doc.paragraphs:
        # Check if the paragraph has a heading style
        if para.style.name in heading_styles:
            # Apply the corresponding template style if it's not already applied
            if para.style.name not in heading_styles.values():
                para.style = heading_styles[para.style.name]
    

    # Assuming Chapter Number is the first paragraph and Chapter Title is the second paragraph
    if len(doc.paragraphs) >= 2:
        chapter_number_paragraph = doc.paragraphs[0]
        chapter_title_paragraph = doc.paragraphs[1]

        # Apply styles
        chapter_number_paragraph.style = 'HS - ChapterNumber [PACKT]'
        chapter_title_paragraph.style = 'HS - ChapterTitle [PACKT]'

    
    # First, apply the "P0 - Normal [PACKT]" style to paragraphs
    for paragraph in doc.paragraphs:
        # Skip chapter number, title, and any headings
        if paragraph == doc.paragraphs[0] or paragraph == doc.paragraphs[1] or paragraph.style.name.startswith("HS"):
            continue
        
        # Skip figure captions and layout information
        if figure_caption_pattern.match(paragraph.text):
            continue

        # Apply "P0 - Normal [PACKT]" to regular paragraphs that are not indented or lists
        if not paragraph.paragraph_format.left_indent and not paragraph.style.name.startswith("List"):
            paragraph.style = 'P0 - Normal [PACKT]'
            # Apply ""CS - InlineCode [PACKT]" style to identified code within texts
            for match in code_term_pattern.finditer(paragraph.text):
                start, end = match.span()
                run = paragraph.add_run(paragraph.text[start:end])
                run.style = 'CS - InlineCode [PACKT]'
                paragraph.text = paragraph.text[:start] + paragraph.text[end:]

    # Next, handle adding layout information after figures
    for i, paragraph in enumerate(doc.paragraphs):
        if figure_caption_pattern.match(paragraph.text):
            # Create layout information string
            layout_info = f"{book_id}_{chapter_number.zfill(2)}_{str(figure_counter).zfill(2)}"

            # Create a new paragraph for the layout information
            layout_info_paragraph = doc.add_paragraph(layout_info)
            layout_info_paragraph.style = 'PF - LayoutInformation [PACKT]'

            # Move the new paragraph after the figure caption
            p = paragraph._element
            p.addnext(layout_info_paragraph._element)

            # Increment the figure counter
            figure_counter += 1
    
    # Save the styled document
    styled_doc_path = os.path.join(app.config['UPLOAD_FOLDER'], 'styled_' + os.path.basename(doc_path))
    doc.save(styled_doc_path)
    return styled_doc_path

if __name__ == '__main__':
    app.run(debug=True)