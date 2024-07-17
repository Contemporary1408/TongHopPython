# pip install num2words
# NUMBER TO WORD:
from num2words import num2words

def number_to_vietnamese_words(number):
    return num2words(number, lang='vi')

# Example usage
number = 123456
words = number_to_vietnamese_words(number)
print(f"{number} in Vietnamese words: {words}")
# #################################################################### Vice versa
# pip install word2number
# WORD TO NUMBER:
from word2number import w2n

def number_to_vietnamese_words(number):
    return w2n.word_to_num(number, lang='vi')

# Example usage
number = "một trăm hai mươi ba nghìn bốn trăm năm mươi sáu"
numeric_value = number_to_vietnamese_words(number)
print(f"{number} in numeric form: {numeric_value}")
