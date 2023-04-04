import re
import openaiAPI
import openai

openai.api_key = openaiAPI.api_key


def translation(sentence):
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=f"Translate the following string from Russiam into Serbian Latin, using only Serbian Latin characters. Cyrillic characters are not allowed. Preserve the code elements. If the word 'разворот' is encountered, use the word 'strana'. If the word 'обложка' is encountered, use the word 'korica':\n\n{sentence}\n\n",
        temperature=0.7,
        max_tokens=300,
        top_p=1.0,
        frequency_penalty=0.0,
        presence_penalty=0.0
    )
    return response['choices'][0]['text'].strip()


with open('C:/Users/Isk/Desktop/ИСК/Дела Аллы/trans/dict4.txt', 'r+',
          encoding='utf-8') as out, \
        open('C:/Users/Isk/Desktop/ИСК/Дела Аллы/trans/pid6166_lid7257_ru-ru_orig.xml', 'r',
             encoding='utf-8') as orig, \
        open('C:/Users/Isk/Desktop/ИСК/Дела Аллы/trans/pid6166_lid7257_ru-ru_finchange.xml', 'w',
             encoding='utf-8') as change:

    for line in orig:
        pattern = r'[А-Яа-я].*[А-Яа-я]'
        if re.findall(pattern, line):
            rus_part = ''.join(re.findall(pattern, line))
            translated_rus_part = translation(rus_part)
            print(line, end='')
            line = line.replace(rus_part, translated_rus_part)
            print(line, end='')
            print(line, end='', file=change)
        else:
            print(line, end='', file=change)
'''    for elem in line:
        translate = translation(elem)'''




'''    line_rus = map(lambda x: x.strip('\n'), rus.readlines())

    line_rs = map(lambda x: x.strip('\n'), rs.readlines())

    trans_dict = dict(zip(line_rus, line_rs))

    for line in orig:
        if '>' not in line:
            print(line, file=out, end='')
            continue
        if any(map(lambda x: 1071 < ord(x) < 1104, line)):
            true_line = line

            strt = true_line.index('>') + 1
            fnsh = -(list(reversed(true_line)).index('<') + 1)

            true_sentence = true_line[strt:fnsh]
            true_sentence_orig = true_sentence
            if true_sentence:
                if true_sentence[0].islower() and trans_dict[true_sentence][0].isupper():
                    trans_dict[true_sentence] = trans_dict[true_sentence][0].lower() + trans_dict[true_sentence][1:]
                true_sentence = trans_dict[true_sentence]

                line = line.replace(true_sentence_orig, true_sentence)

                print(line, file=out, end='')
            else:
                print(line, file=out, end='')
        else:
            print(line, file=out, end='')
with open('C:/Users/Isk/Desktop/ИСК/Дела Аллы/trans/pid6166_lid7257_ru-ru_change_all.xml', 'r',
          encoding='utf-8') as translate, \
open('C:/Users/Isk/Desktop/ИСК/Дела Аллы/trans/dict3.txt', 'w', encoding='utf-8') as rus:
    i = 0
    for line in translate:
        if any(map(lambda x: 1071 < ord(x) < 1104, line)):
            i += 1
            print(f'line {i}: {line}')
            print(line.strip(), file=rus)
        else:
            i += 1
'''

'''
            if true_sentence and true_sentence in trans_dict:
                true_sentence = trans_dict[true_sentence]
                for letter in true_sentence_orig:
                    if letter.isupper():
                        ind = true_sentence_orig.index(letter)
                        true_sentence = true_sentence.replace(true_sentence[ind], true_sentence[ind].upper(), 1)
                if true_sentence_orig[-1] in punctuation:
                    true_sentence += true_sentence_orig[-1]
                print(true_line)
                print(true_sentence_orig)
                print(true_sentence)
                print(line)
                line = line.replace(true_sentence_orig, true_sentence)
                print(line)
                print(line, file=out, end='')
            else:
                print(line, file=out, end='')
        else:
            print(line, file=out, end='')
with open('C:/Users/Isk/Desktop/ИСК/Дела Аллы/trans/pid6166_lid7257_ru-ru_change.xml', 'r', encoding='utf-8') as chek:
    for line in chek:
        print(line.strip())'''
