from youtube_transcript_api import YouTubeTranscriptApi
import xlsxwriter
import xlrd
import pandas as pd

##Get SRT Captions
srt = YouTubeTranscriptApi.get_transcript('j6FNRpraWwM', languages=['es'])
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
workbook = xlsxwriter.Workbook('intermediate3.xlsx')
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


    
workbook.close()


##for x in srt:
##    print(x['text'])

#Open word list
wordDF = pd.read_excel('WordList.xlsx')
print(wordDF['Beginner'][0])

data = {'test': ['hi', 'e'], 'test': ['hi', 'e']}

#Write Excel file
##df = pd.DataFrame(data, columns= ['test','test'])
##df.to_excel('t.xlsx', index = False, header=True)
