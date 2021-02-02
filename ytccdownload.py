from youtube_transcript_api import YouTubeTranscriptApi

srt = YouTubeTranscriptApi.get_transcript('gt8FQO9oxz4', languages=['es'])
print(srt[0])
