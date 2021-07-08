from docx import Document

def count_words(file_name):
    # Count how many times each word is used in a docx file. Takes in a directory of file and returns dictionary with {word:number of word used}.
    count_dict = {}
    document = Document(file_name)

    for paragraph in document.paragraphs:
        words = paragraph.text.lower().split()
        words = separate_punc(words)
        for word in words:
            if word in count_dict:
                count_dict[word] += 1
            else:
                count_dict[word] = 1

    return(count_dict)

def separate_punc(word_list):
    # Separates punctuations from words. Takes in a list of words possibly with punctuations and returns a list with the words and punctuations separated. 
    return_word_list = []
    for word in word_list:
        new_list = []
        for letter in word:
            if letter.isalpha():    
                new_list.append(letter)
            if not letter.isalpha():
                new_list += [' ', letter, ' ']
        letter_str = ''.join(new_list)
        letter_list = letter_str.split()
        return_word_list += letter_list
    return return_word_list

# Running
file_name = input("Enter the file directory:")
print(count_words(file_name))
