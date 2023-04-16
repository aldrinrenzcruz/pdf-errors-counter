import os
import pandas as pd
import PyPDF2

script_folder = os.path.dirname(os.path.abspath(__file__))
errors_file = os.path.abspath("_list.txt")

with open(errors_file, "r", encoding='utf-8') as f:
    errors = f.read().splitlines()

pdf_file = input("Enter the filename of the PDF: ")
print("Processing... This window will auto-close once the report is ready.")
pdf_file_path = os.path.abspath(pdf_file + '.pdf')

with open(pdf_file_path, "rb") as file:
    pdf_reader = PyPDF2.PdfReader(file)
    content = ''
    page_numbers = {}
    error_lines = {}
    for page in range(len(pdf_reader.pages)):
        current_page_content = pdf_reader.pages[page].extract_text()
        lines = current_page_content.split("\n")
        for line in lines:
            content += line + "\n"
            for error in errors:
                if error in line:
                    if error in page_numbers:
                        page_numbers[error].append(page + 1)
                    else:
                        page_numbers[error] = [page + 1]
                    if error in error_lines:
                        error_lines[error].append(line)
                    else:
                        error_lines[error] = [line]

error_counts = {error: content.count(error) for error in errors}
df = pd.DataFrame.from_dict(error_counts, orient='index', columns=['Count'])
df = df[df['Count'] > 0].sort_values(by='Count', ascending=False)
df.index.name = 'Actual Errors'
df['Errors'] = df.index.str.replace(" ", "\u00B7")
df['Page Numbers'] = df.index.map(lambda error: ', '.join(str(p) for p in page_numbers[error]))
df['Error Line'] = df.index.map(lambda error: '\n'.join(error_lines[error]))
df["Blank"] = " "
df = df.reset_index()
cols = ['Errors', 'Actual Errors', 'Count', 'Page Numbers', 'Error Line', 'Blank']
df = df[cols]
report_file = os.path.abspath(pdf_file + "_report.xlsx")
df.to_excel(report_file, index=False)