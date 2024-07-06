'''
Copyright (C) 2024 by Fengze Han. All rights reserved.
Description: A general class for analyzing text files and extracting word frequency.
'''

from collections import Counter
import fitz
from docx import Document
import time
import re
import os
import logging
import sys
import jieba
import jieba.analyse
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, scrolledtext

class FileAnalyzer:
    """
    A class for analyzing text files and extracting word frequency.

    Args:
        file_path (str): The path to the text file.

    Attributes:
        file_path (str): The path to the text file.
        logger (logging.Logger): The logger object for logging information.
    """

    def __init__(self, file_path):
        self.file_path = file_path
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)
        time_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        self.logger.addHandler(logging.FileHandler(time_str + 'log.txt'))
        self.stopwords = set()
        for root, dirs, files in os.walk('stopwords'):
            for file in files:
                if (file.endswith('.txt') == False):
                    continue
                
                with open(os.path.join(root, file), 'r', encoding='utf-8') as f:
                    for line in f:
                        self.stopwords.add(line.strip())
                        
    def _remove_stopwords(self, text):
        santi_words =[x for x in text if len(x) > 1 and x not in self.stopwords]
        return santi_words
        
    def _extract_text_from_pdf(self):
        """
        Extracts text from a PDF file.

        Returns:
            str: The extracted text.
        """
        doc = fitz.open(self.file_path)
        text = ""
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text += page.get_text()
        return text

    def _extract_text_from_docx(self):
        """
        Extracts text from a DOCX file.

        Returns:
            str: The extracted text.
        """
        doc = Document(self.file_path)
        return "\n".join([paragraph.text for paragraph in doc.paragraphs])

    def _word_frequency(self, words):
        """
        Calculates the word frequency in the given words.

        Args:
            words: The words to analyze.

        Returns:
            collections.Counter: A Counter object containing the word frequency.
        """
        return Counter(words)
    
    def set_logger_level(self, level):
        """
        Sets the logger level.

        Args:
            level (int): The logging level.
        """
        self.logger.setLevel(level)     

    def analyze_file(self, top_k_words=10):
        """
        Analyzes the text file and extracts the most common words.

        Args:
            top_k_words (int, optional): The number of most common words to extract. Defaults to 10.

        Returns:
            list: A list of tuples containing the most common words and their frequencies.
        """
        if self.file_path.endswith('.pdf'):
            text = self._extract_text_from_pdf()
        elif self.file_path.endswith('.docx'):
            text = self._extract_text_from_docx()
        else:
            text = ""
            self.logger.error("Unsupported file format")
            raise ValueError("Unsupported file format. Only .pdf and .docx are supported.")

        words = jieba.cut(text)
        cleaned_words = self._remove_stopwords(words)
        cleaned_text = ' '.join(cleaned_words)
        words = jieba.analyse.extract_tags(cleaned_text, topK=1000, withWeight=False, allowPOS=())
        
        # Regex pattern construction for non-Chinese text
        letter_pattern = r'[a-zA-Z]'  # English letters
        number_pattern = r'\d'  # Numbers
        special_char_pattern = r'[^\w\s]'  # Special characters (non-word characters excluding spaces)
        chinese_char_pattern = re.compile(r'[\u4e00-\u9fff]+')
        
        # Extract Chinese characters based on the pattern        
        chinese = [word for word in words if chinese_char_pattern.fullmatch(word)]

        # Extract numbers based on the pattern
        numbers = re.findall(number_pattern, text)
        
        # Extract special characters based on the pattern
        special_characters = re.findall(special_char_pattern, text)
        
        # count the frequency of each chinese character
        chinese_word_counts = self._word_frequency(chinese)
        chinese_common_words = chinese_word_counts.most_common(top_k_words)
        self.logger.info("Most common chinese characters: %s", chinese_common_words)
        
        # count the frequency of each number
        number_counts = self._word_frequency(numbers)
        number_common_words = number_counts.most_common(top_k_words)
        self.logger.info("Most common numbers: %s", number_common_words)
        
        # count the frequency of each special character
        special_char_counts = self._word_frequency(special_characters)
        special_char_common_words = special_char_counts.most_common(top_k_words)
        self.logger.info("Most common special characters: %s", special_char_common_words)
        
        return {"Chinese Words": chinese_common_words, "Numbers": number_common_words, "Special Characters": special_char_common_words}
        
    # method used to count the given word frequency
    def count_given_word_frequency(self, input_word):
        '''
        Count the frequency of the given word in the text file.
        Args:
            input_word (str): The word to count.
        '''
        if self.file_path.endswith('.pdf'):
            text = self._extract_text_from_pdf()
        elif self.file_path.endswith('.docx'):
            text = self._extract_text_from_docx()
        else:
            text = ""
            self.logger.error("Unsupported file format")
            raise ValueError("Unsupported file format. Only .pdf and .docx are supported.")
        
        punctuation_pattern = r"[。，、；：？！「」『』（）【】《》〈〉——……—·～“”‘’.,;:?!\"'()\[\]{}\-_=+&@#$%*~/|\\<>^`]"
        cleaned_text = re.sub(punctuation_pattern, ' ', text)
        
        words = jieba.cut(cleaned_text, cut_all=False)
        chinese = [word.strip() for word in words if word.strip()]
        count = chinese.count(input_word)
        print("The frequency of the given word is: ", count)
        
class GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Text File Analyzer")
        self.file_path = None

        # File selection
        self.select_button = tk.Button(root, text="Select File", command=self.select_file)
        self.select_button.pack(pady=10)

        # Analyze file button
        self.analyze_button = tk.Button(root, text="Analyze File", command=self.analyze_file)
        self.analyze_button.pack(pady=10)

        # Word frequency button
        self.word_freq_button = tk.Button(root, text="Count Word Frequency", command=self.count_word_frequency)
        self.word_freq_button.pack(pady=10)

        # Output area
        self.output_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=120, height=40)
        self.output_area.pack(pady=10, expand=True)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("Word files", "*.docx")])
        if self.file_path:
            messagebox.showinfo("File Selected", f"Selected file: {self.file_path}")

    def analyze_file(self):
        if not self.file_path:
            messagebox.showwarning("No File", "Please select a file first.")
            return
        try:
            analyzer = FileAnalyzer(self.file_path)
            results = analyzer.analyze_file()
            self.output_area.delete(1.0, tk.END)
            self.output_area.insert(tk.END, "Most Common Chinese Words:\n")
            for word, count in results["Chinese Words"]:
                self.output_area.insert(tk.END, f"{word}: {count}\n")

            self.output_area.insert(tk.END, "\nMost Common Numbers:\n")
            for number, count in results["Numbers"]:
                self.output_area.insert(tk.END, f"{number}: {count}\n")
            self.output_area.insert(tk.END, "\nMost Common Special Characters:\n")
            for char, count in results["Special Characters"]:
                self.output_area.insert(tk.END, f"{char}: {count}\n")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def count_word_frequency(self):
        if not self.file_path:
            messagebox.showwarning("No File", "Please select a file first.")
            return
        input_word = simpledialog.askstring("Input", "Enter the word to count its frequency:")
        if input_word:
            try:
                analyzer = FileAnalyzer(self.file_path)
                frequency = analyzer.count_given_word_frequency(input_word)
                messagebox.showinfo("Word Frequency", f"The frequency of '{input_word}' is: {frequency}")
            except Exception as e:
                messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    gui = GUI(root)
    root.mainloop()