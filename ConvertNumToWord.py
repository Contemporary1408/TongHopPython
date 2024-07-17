# Install the num2words library (if you haven't already)
# You can install it using pip:
# pip install num2words

from num2words import num2words

def number_to_vietnamese_words(number):
    return num2words(number, lang='vi')

# Example usage
number = 123456
words = number_to_vietnamese_words(number)
print(f"{number} in Vietnamese words: {words}")
