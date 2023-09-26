import win32com.client
import speech_recognition as sr
import webbrowser
import datetime
import wolframalpha
import wikipedia
import requests
speaker = win32com.client.Dispatch("SAPI.SpVoice")

appId = "QJ99Y4-LUXXL3H7AJ"
client = wolframalpha.Client(appId)

def resolveListOrDict(variable):
  if isinstance(variable, list):
    return variable[0]["plaintext"]
  else:
    return variable["plaintext"]

def removeBrackets(variable):
  return variable.split("(")[0]

def primaryImage(title=''):
    url = 'http://en.wikipedia.org/w/api.php'
    data = {'action':'query', 'prop':'pageimages','format':'json','piprop':'original','titles':title}
    try:
        res = requests.get(url, params=data)
        key = res.json()['query']['pages'].keys()[0]
        imageUrl = res.json()['query']['pages'][key]['original']['source']
        print(imageUrl)
    except Exception as err:
        print('Exception while finding image:= '+str(err))
# method that search wikipedia...
def search_wiki(keyword=''):
  # running the query
  searchResults = wikipedia.search(keyword)
  # If there is no result, print no result
  if not searchResults:
    speaker.Speak("No result from Wikipedia")
    print("No result from Wikipedia")
    return
  # Search for page... try block
  try:
    page = wikipedia.page(searchResults[0])
  except wikipedia.DisambiguationError as err:
    # Select the first item in the list
    page = wikipedia.page(err.options[0])
  #encoding the response to utf-8
  wikiTitle = str(page.title.encode('utf-8'))
  wikiSummary = str(page.summary.encode('utf-8'))
  # printing the result
  speaker.Speak(wikiSummary)
  print(wikiSummary)


def search(text=''):
  res = client.query(text)
  # Wolfram cannot resolve the question
  if res['@success'] == 'false':
     print('Question cannot be resolved')
  # Wolfram was able to resolve question
  else:
    result = ''
    # pod[0] is the question
    pod0 = res['pod'][0]
    # pod[1] may contains the answer
    pod1 = res['pod'][1]
    # checking if pod1 has primary=true or title=result|definition
    if (('definition' in pod1['@title'].lower()) or ('result' in  pod1['@title'].lower()) or (pod1.get('@primary','false') == 'true')):
      # extracting result from pod1
      result = resolveListOrDict(pod1['subpod'])
      speaker.Speak(result)
      print(result)
    else:
      # extracting wolfram question interpretation from pod0
      question = resolveListOrDict(pod0['subpod'])
      # removing unnecessary parenthesis
      question = removeBrackets(question)
      # searching for response from wikipedia
      search_wiki(question)

def takeCommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold =0.5
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry"

if __name__== '__main__':
    print("HI I'm your Smart Assistant")
    speaker.Speak("hello, tell me something")
    while True:
        print("Listening..")
        query= takeCommand()
        sites = [["B.M.S.C.E.","https://bmsce.ac.in"],["YOUTUBE","https://youtube.com"],["google","https://google.com"],["D.H.L","https://bmsgroup.dhi-edu.com/bmsgroup_bmsce"]]
        for site in sites:
            if f"open {site[0]}".lower() in query.lower():
                    speaker.Speak(f"Opening {site[0]} website")
                    webbrowser.open(site[1])


        if "the time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            speaker.Speak(f"The time is {strfTime}")

        search(query)
