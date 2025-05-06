import os
from file_util import docx_to_markdown, markdown_to_docx

if __name__ == "__main__":
    input_docx = "./users/user1/input/William Coronado - CV.docx"
    output_md = "./users/user1/input/William Coronado - CV.md"
    input_md = "./users/user1/output/resume-813d4632c66ccbc2.txt.md" 
    output_docx = "./users/user1/output/813d4632c66ccbc2-CV.docx"

    # Convert from .docx to .md
    if not os.path.exists(input_docx):
        print(f"Error: The file '{input_docx}' does not exist.")
    else:
        docx_to_markdown(input_docx, output_md)
        print(f"Markdown file '{output_md}' has been created successfully.")

    # Convert from .md to .docx
    if not os.path.exists(input_md):
        print(f"Error: The file '{input_md}' does not exist.")
    else:
        markdown_to_docx(input_md, output_docx)
        print(f"Docx file '{output_docx}' has been created successfully.")