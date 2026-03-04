import re

with open('app/utils.py', 'r', encoding='utf-8') as f:
    text = f.read()

text = text.replace('"Formato inválido. Use: RFC-AAAA-NNN"', '"Texto muito curto (mínimo 3 caracteres)"')
text = text.replace('"Invalid format. Use: RFC-YYYY-NNN"', '"Text too short (minimum 3 chars)"')
text = text.replace('"Format invalide. Utilisez: RFC-AAAA-NNN"', '"Texte trop court (min. 3 caractères)"')
text = text.replace('"Ungültiges Format. Verwenden Sie: RFC-JJJJ-NNN"', '"Text zu kurz (min. 3 Zeichen)"')

with open('app/utils.py', 'w', encoding='utf-8') as f:
    f.write(text)

print("Translations updated.")
