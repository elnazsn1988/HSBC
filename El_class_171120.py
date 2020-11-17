import fitz
from bs4 import BeautifulSoup
import shutil
import os 
import pandas as pd


def get_text_percentage(file_name: str) -> float:
    indic = ""
    """
    Calculate the percentage of document that is covered by (searchable) text.

    If the returned percentage of text is very low, the document is
    most likely a scanned PDF
    """
    total_page_area = 0.0
    total_text_area = 0.0

    doc = fitz.open(file_name)
    #doc = fitz.open(document)
    font_counts, styles = fonts(doc, granularity=False)
    size_tag = font_tags(font_counts, styles)
    elements = headers_para(doc, size_tag)
    elements2 = ' '.join(elements)
    soup = BeautifulSoup(elements2, 'lxml')
    root = soup.body
    try:
        for tag in soup.find_all('h1'):
            out_name = f'{tag.name}: {tag.text}'
    except:
        indic = -10
        
   
    if doc.isFormPDF != False:
        acr = 1
        print("woah")
        indic = -300
    else: 
        try:
            acr = 0
            for page_num, page in enumerate(doc):



                    total_page_area = total_page_area + abs(page.rect)
                    text_area = 0.0
                    if total_page_area == 0:
                        indic = -2
                    else :
                        for b in page.getTextBlocks():
                            r = fitz.Rect(b[:4])  # rectangle where block text appears
                            text_area = text_area + abs(r)
                        total_text_area = total_text_area + text_area
                    indic = total_text_area / total_page_area
        except:
            indic = -20
            
    doc.close()
    return indic , acr, out_name


from operator import itemgetter
import pandas as pd
import fitz
import json
def fonts(doc, granularity=False):
    """Extracts fonts and their usage in PDF documents.
    :param doc: PDF document to iterate through
    :type doc: <class 'fitz.fitz.Document'>
    :param granularity: also use 'font', 'flags' and 'color' to discriminate text
    :type granularity: bool
    :rtype: [(font_size, count), (font_size, count}], dict
    :return: most used fonts sorted by count, font style information
    """
    styles = {}
    font_counts = {}

    for page in doc:
        blocks = page.getText("dict")["blocks"]
        for b in blocks:  # iterate through the text blocks
            if b['type'] == 0:  # block contains text
                for l in b["lines"]:  # iterate through the text lines
                    for s in l["spans"]:  # iterate through the text spans
                        if granularity:
                            identifier = "{0}_{1}_{2}_{3}".format(s['size'], s['flags'], s['font'], s['color'])
                            styles[identifier] = {'size': s['size'], 'flags': s['flags'], 'font': s['font'],
                                                  'color': s['color']}
                        else:
                            identifier = "{0}".format(s['size'])
                            styles[identifier] = {'size': s['size'], 'font': s['font']}

                        font_counts[identifier] = font_counts.get(identifier, 0) + 1  # count the fonts usage

    font_counts = sorted(font_counts.items(), key=itemgetter(1), reverse=True)

    if len(font_counts) < 1:
        raise ValueError("Zero discriminating fonts found!")

    return font_counts, styles

def font_tags(font_counts, styles):
    """Returns dictionary with font sizes as keys and tags as value.
    :param font_counts: (font_size, count) for all fonts occuring in document
    :type font_counts: list
    :param styles: all styles found in the document
    :type styles: dict
    :rtype: dict
    :return: all element tags based on font-sizes
    """
    p_style = styles[font_counts[0][0]]  # get style for most used font by count (paragraph)
    p_size = p_style['size']  # get the paragraph's size

    # sorting the font sizes high to low, so that we can append the right integer to each tag 
    font_sizes = []
    for (font_size, count) in font_counts:
        font_sizes.append(float(font_size))
    font_sizes.sort(reverse=True)

    # aggregating the tags for each font size
    idx = 0
    size_tag = {}
    for size in font_sizes:
        idx += 1
        if size == p_size:
            idx = 0
            size_tag[size] = '<p>'
        if size > p_size:
            size_tag[size] = '<h{0}>'.format(idx)
        elif size < p_size:
            size_tag[size] = '<s{0}>'.format(idx)

    return size_tag

def headers_para(doc, size_tag):
    """Scrapes headers & paragraphs from PDF and return texts with element tags.
    :param doc: PDF document to iterate through
    :type doc: <class 'fitz.fitz.Document'>
    :param size_tag: textual element tags for each size
    :type size_tag: dict
    :rtype: list
    :return: texts with pre-prended element tags
    """
    header_para = []  # list with headers and paragraphs
    first = True  # boolean operator for first header
    previous_s = {}  # previous span

    for page in doc:
        blocks = page.getText("dict")["blocks"]
        for b in blocks:  # iterate through the text blocks
            if b['type'] == 0:  # this block contains text
                # REMEMBER: multiple fonts and sizes are possible IN one block
                block_string = ""  # text found in block
                for l in b["lines"]:  # iterate through the text lines
                    for s in l["spans"]:  # iterate through the text spans
                        if s['text'].strip():  # removing whitespaces:
                            if first:
                                previous_s = s
                                first = False
                                block_string = size_tag[s['size']] + s['text']
                            else:
                                if s['size'] == previous_s['size']:

                                    if block_string and all((c == "|") for c in block_string):
                                        # block_string only contains pipes
                                        block_string = size_tag[s['size']] + s['text']
                                    if block_string == "":
                                        # new block has started, so append size tag
                                        block_string = size_tag[s['size']] + s['text']
                                    else:  # in the same block, so concatenate strings
                                        block_string += " " + s['text']

                                else:
                                    header_para.append(block_string)
                                    block_string = size_tag[s['size']] + s['text']

                                previous_s = s

                    # new block started, indicating with a pipe
                    #block_string += "|"

                header_para.append(block_string)
    return list(header_para)
Listed_name  = []
def Eclass(inp_file_comp):
        El_log = pd.DataFrame(columns = ['doc_name','percentage text/ar','acro','filename'])
        #l_log.columns = 
        #inp_file_name = os.path.join(inp_file_path,inp_file_name)
        inp_filename = os.path.basename(inp_file_comp)
        extension = os.path.splitext(inp_file_comp)[1]
        #print(inp_filename)
        Listed_name.append(inp_filename)
        #print(Listed_name)
        try:
            text_perc = get_text_percentage(inp_file_comp)
            print(text_perc[0], " is the percentahge text to image")
            print(text_perc[1], " this is 0 if not acro, 1 if acro")
            print(inp_filename, " this is File Name")
            print(extension, " this is File Extension")
            
            if text_perc[1] == 1 or text_perc[0] == -300:
                print("ACRO")
                shutil.copy(inp_file_comp, os.path.join("Acro", inp_filename))
                
            if text_perc[0] == -10:
                print("let through but check")
                shutil.copy(inp_file_comp, os.path.join("checkdis", inp_filename))
            if text_perc[0] == -20:
                print("let through but check")
                shutil.copy(inp_file_comp, os.path.join("checkdis", inp_filename))
            
            if text_perc[0] == -2:
                print("Secured")
                shutil.copy(inp_file_comp, os.path.join("Secured", inp_filename))
                
            elif text_perc[0] < 0.01 or extension == '.jpg' or extension == '.png':
                print("fully scanned PDF - no relevant text")
                shutil.copy(inp_file_comp, os.path.join("Paper", inp_filename))
            #if text_perc[0] > 0.01 and text_perc[1] == 1:
         
                
            elif text_perc[0] > 0.01 and "Financial Planning Report" in text_perc[2]:
                print("FPR Present")
                shutil.copy(inp_file_comp, os.path.join("FPR", inp_filename))
            elif text_perc[0] > 0.01 and "Financial Planning Report" not in text_perc[2]:
                print("Non-FPR PDF")
                shutil.copy(inp_file_comp, os.path.join("Non-FPR PDF", inp_filename))
                

            else:
                shutil.copy(inp_file_comp, os.path.join("All_Else", inp_filename))

            #print("not fully scanned PDF - text is present")
                #shutil.copy(inp_file_path, os.path.join("Paper", inp_filename))
        except:
            shutil.copy(inp_file_comp, os.path.join("Errors", inp_filename))
        return (Listed_name) 

    #for file in os.listdir(path):
        #Eclass(file)

if __name__ == "__main__":
#import os
    #a_files = ['FPR', 'Acroform', 'All_else']
   # 
    #a_exist = [f for f in a_files if os.path.isfile(f)]
    if not os.path.exists('FPR'):
        os.makedirs('FPR')
    if not os.path.exists('Acro'):
        os.makedirs('Acro')
    if not os.path.exists('All_Else'):
        os.makedirs('All_Else')
    if not os.path.exists('Paper'):
        os.makedirs('Paper')
    if not os.path.exists('Errors'):
        os.makedirs('Errors')
    if not os.path.exists('Secured'):
        os.makedirs('Secured')
    if not os.path.exists('checkdis'):
        os.makedirs('checkdis')
    if not os.path.exists('Non-FPR PDF'):
        os.makedirs('Non-FPR PDF')

    path = "/home/jupyter/bucket_inp"
    for file in os.listdir(path):
        if not file.startswith('.'):
            Eclass(os.path.join(path, file))

