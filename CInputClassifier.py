import xlsxwriter
import xlrd
import pandas as pd
import time
import os

files = []

for file in os.listdir("/Users/mac/Desktop/repos/youtube-comprehensible-input/youtube-comprehensible-input/"):
    if file.endswith(".xlsx"):
        
        print(file)
        files.append(file)

def loop(file):

    #Open word list
    vocabDF = pd.read_excel('WordList.xlsx')

    #Open video list
    wordDF = pd.read_excel(file)

    #Set distinct and normal word count variables
    beginner_words_distinct = 0
    beginner_words = 0

    #Set distinct and normal word array
    vocab_arr_distinct = []
    word_arr_distinct = []

    vocab_arr = []
    word_arr = []


    #loop through vocab, add distinct words to array
    for v in vocabDF['Beginner']:
        v = v.replace(',', '')
        v = v.replace('¡', '')
        v = v.replace('.', '')
        v = v.replace('!', '')
        v = v.replace(' ', '')

        if v not in vocab_arr_distinct:
            vocab_arr_distinct.append(v.lower())

    #loop through video words, add distinct words to array
    for w in wordDF['Words']:
        w = w.replace(',', '')
        w = w.replace('¡', '')
        w = w.replace('!', '')
        w = w.replace('.', '')
        w = w.replace(' ', '')

        if w not in word_arr_distinct:
            word_arr_distinct.append(w.lower())
    counter = 0

    #loop and if video word in vocab, add 1 to beginner words
    for word in word_arr_distinct:
        if word in vocab_arr_distinct:
            beginner_words_distinct += 1
        
    for v in vocabDF['Beginner']:
        vocab_arr.append(v)

    for w in wordDF['Words']:
        word_arr.append(w)

    ##preprocesss
        #make all lower case
    vocab_arr = [x.lower() for x in vocab_arr]
    word_arr = [x.lower() for x in word_arr]

      
    counter = 0

    sentences_arr = []
    sentence = ''
    #create sentences of 8 words
    for word in word_arr:
        #replace invalid characters
        word = word.replace(',', '')
        word = word.replace('.', '')
        word = word.replace('¡', '')
        word = word.replace('!', '')
        word = word.replace(' ', '')
        
        if counter < 12:
            sentence = sentence +  word + ' '
        else:
            counter = 0
            sentences_arr.append(sentence)
            sentence = ''
        counter += 1


    hard_sentences = 0

    #loop through sentences if more than n words are not in vocab, that is a hard sentence 

    for sentence in sentences_arr:
        hard_words = 0
        
        x = sentence

        for word in x.split():
            if word not in vocab_arr:
                if word == 'se':
                    break
                else:
                    hard_words += 1
    ##            print(word)
                
                

            if hard_words == 4:
    #            print(hard_words)
##                print(word)
    ##            time.sleep(1)
                hard_sentences += 1
                hard_words = 0
                break


            
##    print('')            
##    print('Hard Sentences')
##    print(hard_sentences)
##    print('Total Sentences')
##    print(len(sentences_arr))






    for word in word_arr:
        if word in vocab_arr:
            beginner_words += 1
        

##    print('Distinct vocab words in video:')
##    print(beginner_words_distinct)
##    print('Distinct vocab')
##    print(len(vocab_arr_distinct))
##    print('Distinct words in video')
##    print(len(word_arr_distinct))

    percent_distinct = beginner_words_distinct/len(word_arr_distinct)

    ##print('Percentage of distinct beginner words in video')
    ##print(percent_distinct*100)
    
    print('\n')
    print(file)
    print('\n')

    print('Vocab words in video:')
    print(beginner_words)
    print('Total vocab')
    print(len(vocab_arr))
    print('Total words in video')
    print(len(word_arr))

    percent = beginner_words/len(word_arr)
    print('')
    print('Percentage of distinct beginner words in video')
    print(percent_distinct*100)
    print('')
    print('Percentage of beginner words in video')
    print(percent*100)
    print('')
    print('Speed of speech, WPM')
    time = wordDF['Time'][0]
    print(len(word_arr)/time)
    print('')

    print('Percent of hard sentences')
    print((hard_sentences/len(sentences_arr))*100)
    print('')


    vocab_percent = (beginner_words/len(word_arr))*100
    distinct_vocab_percent = (beginner_words_distinct/len(word_arr_distinct))*100
    wpm = len(word_arr)/time
    hard_sentences_percent = (hard_sentences/len(sentences_arr))*100

    score = wpm + hard_sentences_percent -vocab_percent - distinct_vocab_percent


    print('Score')
    print(score)

    return score


scores = []

for file in files:

    if file == 'WordList.xlsx':
        continue
    elif file == '~$WordList.xlsx':
        continue
    else:
        score = loop(file)
        scores.append(score)

results = []

counter = 0
for file in files:
    if file == 'WordList.xlsx':
        continue
    elif file == '~$WordList.xlsx':
        continue
    else:
        print('')
        print(file)
        print(scores[counter])
        results.append([scores[counter], file])
        print('')

    counter += 1
results.sort()
data = {'Video Title':[], 'Score': [], 'Level', 'URL', 'Words'}

for x in results:
    print('\n')
    data['Video Title'].append(x[1])
    data['Score'].append(x[0])

    print(x)

df = pd.DataFrame(data, columns=['Video Title', 'Score'])
df.to_excel('table.xlsx', index = False, header = True)



                    
                    
