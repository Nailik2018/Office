# -*- coding: UTF-8 -*-
import os
import comtypes.client

class WordToPdf():

    def __init__(self, original_word_path, pdf_safe_path):

        self.original_word_path = original_word_path
        self.pdf_safe_path = pdf_safe_path

    def create_pdf(self):

        try:
            pdf_format_key = 17
            file_in = os.path.abspath(self.original_word_path)
            file_out = os.path.abspath(self.pdf_safe_path)
            worddoc = comtypes.client.CreateObject('Word.Application')
            doc = worddoc.Documents.Open(file_in)
            doc.SaveAs(file_out, FileFormat=pdf_format_key)
            doc.Close()
            worddoc.Quit()
        except Exception as e:
            print(e)