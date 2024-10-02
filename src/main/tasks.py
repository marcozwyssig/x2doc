import os
from invoke import *
from pathlib import Path
from document import Document

@task
def create_word_from_x2doc(c, input_file, output_file):
    input_file = Path(input_file)
    if (not input_file.exists()):
        print(f"File {input_file} does not exist")
        return
    
    output_file = Path(output_file)
    if (output_file.exists()):
        print(f"File {output_file} already exists. Will be overwritten.")
        os.remove(output_file)
   
    doc = Document.from_xml(input_file.read_text())
    doc.to_word(output_file)
    
@task
def create_x2doc_from_word(c, input_file, output_file):
    input_file = Path(input_file)
    if (not input_file.exists()):
        print(f"File {input_file} does not exist")
        return
    
    output_file = Path(output_file)
    if (output_file.exists()):
        print(f"File {output_file} already exists. Will be overwritten.")
        os.remove(output_file)
   
    doc = Document.from_word(input_file)
    output_file.write_text(doc.to_xml())
    
