# features.py
from sklearn.base import BaseEstimator, TransformerMixin
import numpy as np

class FeatureExtractor(BaseEstimator, TransformerMixin):
    def fit(self, X, y=None):
        return self
    
    def transform(self, X):
        features = []
        for texto in X:
            tech_words = sum(1 for word in texto.split() if word in ['driver', 'erro', 'codigo', 'senha'])
            is_hardware = any(word in texto for word in ['conector', 'antena', 'fisic'])
            features.append([tech_words, int(is_hardware)])
        return np.array(features)