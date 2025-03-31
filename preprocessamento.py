# preprocessamento.py
import re
from nltk.stem import WordNetLemmatizer
from nltk.corpus import stopwords
import nltk

nltk.download('stopwords')
nltk.download('wordnet')

class TextPreprocessor:
    def __init__(self):
        self.lemmatizer = WordNetLemmatizer()
        self.stop_words = set(stopwords.words('portuguese'))
        self.stop_words.update(["computador", "problema", "sistema"])

    def normalize_terms(self, text):
        text = text.lower()
        text = re.sub(r'\b(pc|notebook|desktop|máquina)\b', 'computador', text)
        text = re.sub(r'\b(app|aplicativo|programa)\b', 'software', text)
        text = re.sub(r'\b(wifi|wi-fi|wireless)\b', 'internet', text)
        text = re.sub(r'\b(travando|travada|congelando|demorando)\b', 'lentidão', text)  # Nova linha
        text = re.sub(r'\b(nao|ñ)\b', 'não', text)
        return text

    def preprocess(self, text):
        text = self.normalize_terms(text)
        text = re.sub(r'[^\w\s]', '', text)
        tokens = text.split()
        tokens = [self.lemmatizer.lemmatize(token) for token in tokens]
        tokens = [token for token in tokens if token not in self.stop_words and len(token) > 2]
        return ' '.join(tokens)