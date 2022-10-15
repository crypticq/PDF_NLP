import PyPDF2
import spacy
from spacy.lang.en.stop_words import STOP_WORDS
from string import punctuation
from heapq import nlargest
from docx import Document


class FrequencySummarizer:

    def __init__(self , pdf_file , output_file):
        self.pdf_file = pdf_file
        self.nlp = spacy.load('en_core_web_sm')
        self.stopwords = list(STOP_WORDS)
        self.punctuation = punctuation
        self.text = self.get_pdf_text()
        self.output_file = output_file + ".docx"



    def get_pdf_text(self):
        pdf_file_obj = open(self.pdf_file, 'rb')
        text = ""
        pdf_reader = PyPDF2.PdfFileReader(pdf_file_obj , strict=False)
        for page in range(pdf_reader.numPages):
            page_obj = pdf_reader.getPage(page)
            text += page_obj.extractText()

        return text


    def get_summary(self, sentences_count=5):
        doc = self.nlp(self.text)
        word_frequencies = {}
        for word in doc:
            if word.text not in self.stopwords and word.text not in self.punctuation:
                if word.text not in word_frequencies.keys():
                    word_frequencies[word.text] = 1
                else:
                    word_frequencies[word.text] += 1
        max_frequency = max(word_frequencies.values())
        for word in word_frequencies.keys():
            word_frequencies[word] = word_frequencies[word]/max_frequency
        sentence_list = [ sentence for sentence in doc.sents ]
        sentence_scores = {}
        for sent in sentence_list:
            for word in sent:
                if word.text.lower() in word_frequencies.keys():
                    if len(sent.text.split(' ')) < 30:
                        if sent not in sentence_scores.keys():
                            sentence_scores[sent] = word_frequencies[word.text.lower()]
                        else:
                            sentence_scores[sent] += word_frequencies[word.text.lower()]
        summary_sentences = nlargest(sentences_count, sentence_scores, key=sentence_scores.get)
        final_sentences = [ w.text for w in summary_sentences ]
        summary = ' '.join(final_sentences)
        return summary

    def write_summary(self, summary):
        document = Document()
        document.add_heading('Summary', 0)
        document.add_paragraph(summary)

        document.save(self.output_file)
        print("Summary saved as {}".format(self.output_file))




if __name__ == '__main__':

    file_input = input("Enter the name of the pdf file: ")
    file_output = input("Enter the name of the output file without .docx : ")
    fs = FrequencySummarizer(file_input , file_output)
    summary = fs.get_summary()
    fs.write_summary(summary)


