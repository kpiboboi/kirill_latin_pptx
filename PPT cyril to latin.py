from pptx import Presentation
from transliterate import translit
import re

def transliterate_word(word):
    translit_rules = {
        "ю" : "yu", "Ю" : "Yu",
        "ў" : "o'", "Ў" : "O'",
        "ё" : "yo", "Ё" : "Yo",
        "ғ" : "g'", "Ғ" : "G'",
        "қ" : "q", "Қ" : "Q",
        "ҳ" : "h", "Ҳ" : "H",
        "х" : "x", "Х" : "X",
        "ж" : "j", "Ж" : "J",
        "й" : "y", "Й" : "Y",
        "ы" : "i", 
        # Добавьте другие замены, если необходимо
    }

    # Заменяем символы в слове согласно правилам
    for cyrillic_char, latin_char in translit_rules.items():
        word = word.replace(cyrillic_char, latin_char)

    # Проверяем наличие "е" в начале слова и заменяем на "ye"
    if word.startswith("е"):
        if len(word) > 1:
            word = "ye" + word[1:]
        else:
            word = "ye"
    elif word.startswith("Е"):
        if len(word) > 1:
            word = "Ye" + word[1:]
        else:
            word = "Ye"
    return word
    

def transliterate_presentation(presentation_file):
    prs = Presentation(presentation_file)

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text_frame = shape.text_frame
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        cyrillic_text = run.text
                        words = re.split(r"(\s+)", cyrillic_text)
                        latin_words = [transliterate_word(word) for word in words]
                        latin_text = "".join(latin_words)
                        run.text = translit(latin_text, "ru", reversed=True)

    prs.save(presentation_file)

# Пример использования:
presentation_file = r"C:\Users\s.ibodov\Downloads\tst ppt\test.pptx"
transliterate_presentation(presentation_file)
print("Готово")