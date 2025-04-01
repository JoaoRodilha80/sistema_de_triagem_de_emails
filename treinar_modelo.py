# treinar_modelo.py
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.pipeline import make_pipeline
from sklearn.model_selection import cross_val_score
from sklearn.metrics import classification_report
import joblib
from preprocessamento import TextPreprocessor

# 1. Carregar dados
df = pd.read_csv("emails_treinamento.csv")
preprocessor = TextPreprocessor()
df['texto_processado'] = df['texto'].apply(preprocessor.preprocess)

# 2. Pipeline do modelo (original)
modelo = make_pipeline(
    TfidfVectorizer(max_features=5000, ngram_range=(1, 2)),
    MultinomialNB()
)

# 3. Treinar e avaliar (original)
modelo.fit(df['texto_processado'], df['categoria'])
scores = cross_val_score(modelo, df['texto_processado'], df['categoria'], cv=5)
print(f"Acurácia média: {scores.mean():.2f} (±{scores.std():.2f})")
print(classification_report(df['categoria'], modelo.predict(df['texto_processado'])))

# 4. Salvar modelo (original)
joblib.dump(modelo, "modelo_classificador.pkl")
joblib.dump(preprocessor, "preprocessor.pkl")
print("✅ Modelo treinado e salvo!")