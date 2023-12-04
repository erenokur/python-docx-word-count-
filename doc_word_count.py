#!/usr/bin/python

import sys
import docx
import re
import os

def count_docx(file_name):
    document = docx.Document(file_name)

    newparatextlist = []
    for paratext in document.paragraphs:
        newparatextlist.append(paratext.text)
    
    return len(re.findall(r'\w+', '\n'.join(newparatextlist)))

def count_txt(file_name):
    wordcount = 0
    try:
        document = open(file_name)
    except:
        print('Cannot open file to read')
        exit()
    while 1:
        lines = document.readlines(100000)
        if not lines:
            break
        for line in lines:
            wordcount = wordcount + len(re.findall(r'\w+', line))
    return wordcount


if __name__ == '__main__':
    extensions = {
        'txt' : count_txt,
        'docx' : count_docx,
        'doc' : count_docx,
    }
    file_name, file_extension = os.path.splitext(sys.argv[1])
    print(extensions[file_extension.lower().replace('.','')](sys.argv[1]))
