import re
import pandas as pd

class TextProcessor:
    def __init__(self, abbreviation_file):
        self.abbreviation_dict = self.load_abbreviation_dict(abbreviation_file)

    @staticmethod
    def load_abbreviation_dict(file_path):
        df = pd.read_excel(file_path)
        abbreviation_dict = dict(zip(df['Abbreviation'], df['Full_word']))
        return abbreviation_dict

    def expand_abbreviations(self, text):
        words = text.split()
        expanded_words = [self.abbreviation_dict.get(word, word) for word in words]
        return " ".join(expanded_words)
        
    def clean_text(self, text):
        patterns_to_remove = [
            r"<.>", 
            r"<strong>",
            r"-", 
            r"</p>", 
            r"&nbsp;</p>", 
            r"\n", 
            r",", 
            r"#", 
            r"&nbsp;",
            r"\(", 
            r"\t", 
            r"\)", 
            r"\/", 
            r"<p>", 
            r"\.", 
            r"\>", 
            r"\<"
        ]
        
        combined_pattern = "|".join(patterns_to_remove)

        # Appliquez en une seule fois :
        text = re.sub(combined_pattern, " ", text).strip()
        
        # Convertir en minuscules
        text = text.lower()
        
        #on a 20% de moins en appliquant cette ligne          
        # Supprimer tout ce qui suit "IAW", "CMM", "OKIAW", ... 
        text = re.sub(r"\b(iaw|amm|i.a.w|a.m.m|serviaw|cmm|okiaw|ok|nowiaw|propiaw|propriaw|properlyiaw|iawamm)\b.*", " ", text, flags=re.IGNORECASE)
        
        # Étendre les abréviations
        text = self.expand_abbreviations(text)
        
        # Supprimer les espaces multiples
        text = re.sub(r"\s+", " ", text).strip()
        
        return text
    