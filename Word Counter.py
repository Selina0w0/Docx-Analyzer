from docx import Document

def separate_punc(word_list):
    # Separates punctuations and numbers from words. 
    # Takes in a list of words possibly with punctuations and returns a list with the words, numbers, and punctuations separated. 
    # new_list is used as temporary list of separated words.
    # digit_list is used as a variable to concatenate consecutive numbers in str format.

    return_word_list = []
    for word in word_list:
        print(f'word: {word}')
        new_list = []
        digit_list = ''

        for letter in word:
            if letter.isdigit():
                digit_list += letter

            elif digit_list == '':
                # When there is no stored numbers
                if letter.isalpha(): 
                    new_list.append(letter)
                else:
                    new_list += [' ', letter, ' ']
            else:
                # When there is stored numbers
                if letter.isalpha():    
                    new_list += [' ', digit_list, ' ']
                    digit_list = ''
                    new_list.append(letter)
                else:
                    new_list += [' ', digit_list, ' ']
                    digit_list = ''
                    new_list += [' ', letter, ' ']

        letter_str = ''.join(new_list)
        letter_list = letter_str.split()
        return_word_list += letter_list
    return return_word_list


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


# Running
file_name = input("Enter the file directory:")
# Sorts the dictionary of words in the order of most used to least used.
sorted_dict = dict(sorted(count_words(file_name).items(), key=lambda word: word[1], reverse=True))
print(sorted_dict)
