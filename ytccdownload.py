from youtube_transcript_api import YouTubeTranscriptApi
import xlsxwriter
import xlrd
import pandas as pd
import time

video_code = 'j6FNRpraWwM'
##video_arr = [['j6FNRpraWwM', 'The Plateau Is a Myth - Intermediate Spanish - Language Learning #12','DreamingSpanish'],
##['s1rXXMDSJvg','Questions & Answers 2 - Intermediate Spanish - Q&A #3','DreamingSpanish']]
video_arr = []
path = '/Users/mac/Desktop/repos/youtube-comprehensible-input/youtube-comprehensible-input/'

f = open (path+'videos-adv.txt', 'r')
for x in f:
    video_arr.append(x.split(','))

def download_captions(code, title, channel):

    ##Get SRT Captions
    srt = YouTubeTranscriptApi.get_transcript(code, languages=['es'])
    ##print(srt[0])
    ##print(srt[1])

    sentences = []
    ytWords = []

    time = 0
    counter = 1
    ##calculate time of video
    for x in srt:
        time += x['duration']
        
        if counter == len(srt):
            time = x['start']/60
                
        counter += 1
        
        
    print("Duration")
    print(time)
    ##for each caption in srt add to sentences array
    for x in srt:
        sentences.append(x['text'])


    ##for each sentence split by word and add to ytwords
    for sentence in sentences:
        for x in sentence.split():
            ytWords.append(x)

    print(ytWords)

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
    worksheet.write(1,2, 'www.youtube.com/watch?v='+video_code)
    worksheet.write(0,3, 'Channel')
    worksheet.write(1,3, channel)
    worksheet.write(0,4, 'Title')
    worksheet.write(1,4, title)


        
    workbook.close()


    ##for x in srt:
    ##    print(x['text'])

    #Open word list
    wordDF = pd.read_excel(path+'WordList.xlsx')
    ##print(wordDF['Beginner'][0])

    ##data = {'test': ['hi', 'e'], 'test': ['hi', 'e']}

    #Write Excel file
    ##df = pd.DataFrame(data, columns= ['test','test'])
    ##df.to_excel('t.xlsx', index = False, header=True)

for x in video_arr:
    download_captions(x[0], x[1], x[2])
    time.sleep(5)
