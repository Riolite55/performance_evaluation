import markdown
from weasyprint import HTML

def markdown_to_pdf(markdown_text: str, output_filename: str):
    """
    Converts a Markdown string into a PDF file with rendered formatting.
    
    :param markdown_text: A string containing Markdown content.
    :param output_filename: The name of the output PDF file.
    """
    # Convert Markdown to HTML
    html_content = markdown.markdown(markdown_text, extensions=['extra', 'tables'])
    
    # Wrap in a basic HTML template
    full_html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; }}
            h2 {{ color: #333; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            th, td {{ border: 1px solid black; padding: 8px; text-align: left; }}
        </style>
    </head>
    <body>
        {html_content}
    </body>
    </html>
    """
    
    # Convert HTML to PDF
    HTML(string=full_html).write_pdf(output_filename)
    print(f"PDF saved as {output_filename}")

# Example usage
markdown_text = """## Subtitle

This is a sample paragraph with **bold** and *italic* text.

- Item 1
- Item 2

| Column 1 | Column 2 |
|----------|----------|
| Data 1   | Data 2   |
"""

markdown_to_pdf(markdown_text, "output3.pdf")
