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
    print(file)
    #Open word list
    vocabDF = pd.read_excel("/Users/mac/Desktop/repos/youtube-comprehensible-input/youtube-comprehensible-input/backup/WordList.xlsx")

    #Open video list
    wordDF = pd.read_excel("/Users/mac/Desktop/repos/youtube-comprehensible-input/youtube-comprehensible-input/"+file)

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

        #Check if string is not a number/special characters
        if (v.isalpha() == True):
            if v not in vocab_arr_distinct:
                vocab_arr_distinct.append(v.lower())

    #loop through video words, add distinct words to array
    for w in wordDF['Words']:
        w = str(w)
        w = w.replace(',', '')
        w = w.replace('¡', '')
        w = w.replace('!', '')
        w = w.replace('.', '')
        w = w.replace(' ', '')
        #Check if string is not a number/special characters
        if (w.isalpha() == True):
            if w not in word_arr_distinct:
                word_arr_distinct.append(w.lower())
    counter = 0

    #loop and if video word in vocab, add 1 to beginner words
    for word in word_arr_distinct:
        if word in vocab_arr_distinct:
            beginner_words_distinct += 1
        
    for v in vocabDF['Beginner']:
        v = v.replace(',', '')
        v = v.replace('¡', '')
        v = v.replace('.', '')
        v = v.replace('!', '')
        v = v.replace(' ', '')
        #Check if string is not a number/special characters
        if (v.isalpha() == True):
            vocab_arr.append(v)

    for w in wordDF['Words']:
        w = str(w)

        w = w.replace(',', '')
        w = w.replace('¡', '')
        w = w.replace('!', '')
        w = w.replace('.', '')
        w = w.replace(' ', '')
        #Check if string is not a number/special characters
        if (w.isalpha() == True):
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
        #Check if string is not a number/special characters
        if (word.isalpha() == True):
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




    hard_words = []

    for word in word_arr:
        word = word.replace(' ', '')
        if word in vocab_arr:
            beginner_words += 1
        else:
            hard_words.append(word)
        
        

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

    if (score > 150) and (score<180):
        level = 'Advanced'
    elif(score>180):
        level = 'Very Advanced'
    elif (score<150) and (score>120):
        level = 'Upper intermediate'
    elif (score > 80) and (score<100):
        level = 'Upper Beginner'
    elif (score < 80) and (score>40):
        level = 'Beginner'
    elif(score<40):
        level = 'Super Beginner'
    elif (score>100) and (score<120):
        level = 'Intermediate'

    return_arr = [score, level,file, wordDF['URL'][0],wordDF['Channel'][0], vocab_percent,distinct_vocab_percent,wpm,hard_sentences_percent,word_arr, hard_words,wordDF['Title'][0]]
    


    return return_arr


scores = []

##Remove word list .xlsx from list
if 'WordList.xlsx' in files: files.remove('WordList.xlsx')
if '~$WordList.xlsx' in files: files.remove('~$WordList.xlsx')

tot_results = []


for file in files:

    if file == 'WordList.xlsx':
        continue
    elif file == '~$WordList.xlsx':
        continue
    else:
        return_arr = loop(file)
        scores.append(return_arr[1])
        tot_results.append(return_arr)

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
data = {'Video Title':[],'Score': [],'Level': [], 'File':[], 'URL': [], 'Channel': [],'Beginner Vocab %':[],
'Distinct Beginner Vocab %':[],'WPM':[],'Hard Sentences %':[],'Words': [], 'Hard Words': []}

for x in tot_results:
    print('\n')
    data['Video Title'].append(x[11])
    data['Score'].append(x[0])
    data['Level'].append(x[1])
    data['URL'].append(x[3])
    data['Channel'].append(x[4])
    data['Words'].append(x[9])
    data['Hard Words'].append(x[10])
    data['File'].append(x[2])
    data['Beginner Vocab %'].append(x[5])
    data['Distinct Beginner Vocab %'].append(x[6])
    data['WPM'].append(x[7])
    data['Hard Sentences %'].append(x[8])



    ##print(x)

df = pd.DataFrame(data, columns=['Video Title','Level','Score','Channel','URL', 'Beginner Vocab %',
'Distinct Beginner Vocab %','WPM', 'Hard Sentences %','Words', 'Hard Words'])
df.to_excel("/Users/mac/Desktop/repos/youtube-comprehensible-input/sfs.xlsx", index = False, header = True)



                    
                    
