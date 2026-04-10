import zipfile
import re

def process_docx(filepath):
    print(f'\n--- Analyzing {filepath} ---')
    try:
        with zipfile.ZipFile(filepath, 'r') as docx:
            xml_str = docx.read('word/document.xml').decode('utf-8')
            raw_text = re.sub(r'<[^>]+>', '', xml_str)
            placeholders_in_raw = set(re.findall(r'{{(.*?)}}', raw_text))
            
            print('Placeholders:')
            for p in sorted(list(placeholders_in_raw)):
                print(f'  - {p}')
            
            t_texts = re.findall(r'<w:t(?:.*?)>(.*?)</w:t>', xml_str)
            splits = [t for t in t_texts if '{' in t and '{{' not in t and '}' not in t]
            if splits:
                print('Possible split curly braces found in runs:', splits)
            else:
                print('No split placeholders detected.')
            
            if '<a:ln' in xml_str:
                ln_tags = set(re.findall(r'<a:ln.*?</a:ln>', xml_str, re.DOTALL))
                print('Background image border/outline tags (<a:ln>):')
                for ln in ln_tags:
                    if 'w=\"' in ln:
                       print('  - Found <a:ln> with width (border might be visible)')
            else:
                print('No <a:ln> tags found.')
                
    except Exception as e:
        print(f'Error: {e}')

process_docx('c:/Users/nikun/OneDrive/Documents/Desktop/GSE new/templates/bank-quotation.docx')
process_docx('c:/Users/nikun/OneDrive/Documents/Desktop/GSE new/templates/client-quotation.docx')
