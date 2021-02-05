import urllib.request
import simplejson

id = 'j6FNRpraWwM'
url = 'https://youtube.com/w?=' + id

json = simplejson.load(urllib.request.urlopen(url))
print(json)
title = json['entry']['title']['$t']
author = json['entry']['author'][0]['name']

print(id, author, title)