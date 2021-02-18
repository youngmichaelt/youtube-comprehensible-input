from youtube_transcript_api import YouTubeTranscriptApi
import xlsxwriter
import xlrd
import pandas as pd
import time
import random

video_code = 'j6FNRpraWwM'
##video_arr = [['j6FNRpraWwM', 'The Plateau Is a Myth - Intermediate Spanish - Language Learning #12','DreamingSpanish'],
##['s1rXXMDSJvg','Questions & Answers 2 - Intermediate Spanish - Q&A #3','DreamingSpanish']]
video_arr = []
path = '/Users/mac/Desktop/repos/youtube-comprehensible-input/youtube-comprehensible-input/'

f = open (path+'final.txt', 'r')

new = ''

t = pd.read_csv(path+'new.csv')
print(t)
counter = 0

while counter <len(t):
    video_arr.append([t['link'][counter],t['title'][counter],t['channel'][counter]])
    counter += 1





def download_captions(code, title, channel):

    #preprocess 
    code = code.split("v=")[1]
    code = code.strip('\"')
    code = code.replace('"', '')
    code = code.replace('"', '')


    title = title.strip('\"')
    title = title.replace('"','')
    title = title.replace('"','')


    channel = channel.strip('\"')
    channel = channel.replace('"', '')
    channel = channel.replace('"', '')


    print(code, title, channel)

    if code == 'mGgnhpZ8d5g':
        srt = YouTubeTranscriptApi.get_transcript(code, languages=['es-419'])
    else:
    ##Get SRT Captions
        srt = YouTubeTranscriptApi.get_transcript(code, languages=['es'])
    
    sentences = []
    ytWords = []

    time = 0
    duration=0
    counter = 1
    ##calculate time of video
    for x in srt:
        duration += x['duration']
        
        if counter == len(srt):
            time = x['start']/60
                
        counter += 1
        
        
    print("Duration")
    print(duration/60)
    print("Time")
    print(time)
    ##for each caption in srt add to sentences array
    for x in srt:
        sentences.append(x['text'])


    ##for each sentence split by word and add to ytwords
    for sentence in sentences:
        for x in sentence.split():
            ytWords.append(x)

    #print(ytWords)

    # Create a test workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(path+code+'.xlsx')
    worksheet = workbook.add_worksheet()
    ##
    counter = 0

    worksheet.write(0,0, 'Words')
    ##add words to test xlsx
    for row in range(len(ytWords)):
        worksheet.write(row+1, 0, ytWords[counter])
        counter+=1
    worksheet.write(0,1, 'Time')
    worksheet.write(1,1, time)
    worksheet.write(0,2, 'URL')
    worksheet.write(1,2, 'www.youtube.com/watch?v='+code)
    worksheet.write(0,3, 'Channel')
    worksheet.write(1,3, channel)
    worksheet.write(0,4, 'Title')
    worksheet.write(1,4, title)


        
    workbook.close()


    ##for x in srt:
    ##    print(x['text'])

    #Open word list
    ##print(wordDF['Beginner'][0])

    ##data = {'test': ['hi', 'e'], 'test': ['hi', 'e']}

    #Write Excel file
    ##df = pd.DataFrame(data, columns= ['test','test'])
    ##df.to_excel('t.xlsx', index = False, header=True)

for x in video_arr:

    try:
        download_captions(x[0], x[1], x[2])
        time.sleep(random.randint(4,10))
    except:
        print('Error occured')
        print(x)
