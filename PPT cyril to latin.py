from pptx import Presentation
from transliterate import translit
import re

def transliterate_presentation(presentation_file):
    prs = Presentation(presentation_file)

    def transliterate_word(word):
        if word.startswith("е") or word.startswith("Е"):
            if len(word) > 1:
                return "ye" + word[1:]
            else:
                return "ye"
        else:
            translit_rules = {
                "ю" : "yu",
                "ў" : "o'",
                "ё" : "yo",
                "ғ" : "g'",
                "ы" : "i",
                # Добавьте другие замены, если необходимо
            }
            return translit(word, "ru", reversed=True, schema=translit_rules)

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
                        run.text = latin_text

    new_presentation_file = r"C:\Users\s.ibodov\Downloads\tst ppt\test.pptx"
    prs.save(new_presentation_file)

# Пример использования:
presentation_file = r"C:\Users\s.ibodov\Downloads\tst ppt\test.pptx"
transliterate_presentation(presentation_file)
print("Готово")
