import wikipedia
import wolframalpha
import speech_recognition as sr
import win32com.client
import sys
def recognize_speech_from_mic(recognizer, microphone):
    # check that recognizer and microphone arguments are appropriate type
    if not isinstance(recognizer, sr.Recognizer):
        raise TypeError("`recognizer` must be `Recognizer` instance")

    if not isinstance(microphone, sr.Microphone):
        raise TypeError("`microphone` must be `Microphone` instance")

    # adjust the recognizer sensitivity to ambient noise and record audio
    # from the microphone
    with microphone as source:
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
        
    # set up the response object
    # try recognizing the speech in the recording
     # if a RequestError or UnknownValueError exception is caught,
    #     update the response object according
        

 # try recognizing the speech in the recording
     # if a RequestError or UnknownValueError exception is caught,
    #update the response object accordingly
        response["transcription"] = recognizer.recognize_google(audio)

    return response

    # create recognizer and mic instances
while True:
    print("your key pls...")
    try:
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        recognizer = sr.Recognizer()
        microphone = sr.Microphone()     
        preinp =recognize_speech_from_mic(recognizer, microphone)
        print("You said: {}".format(preinp["transcription"]))
        my_input=(preinp["transcription"])
    except sr.UnknownValueError:
        print("sry I can't get you")
        speaker.speak("sry I can't get you")
        continue
    #key=["your key here"]
    if(my_input==key[0]):
        while True:  
            greet=("How do I help u :")
            print(greet)
            speaker.Speak(greet)
            try:
                preinp =recognize_speech_from_mic(recognizer, microphone)
                print("You said: {}".format(preinp["transcription"]))
                my_input1=(preinp["transcription"])
                key1=["shutdown"]
                if(my_input1==key1[0]):
                    print("thank u")
                    sys.exit()
            except sr.UnknownValueError:
                print("sry I can't get you")
                speaker.speak("sry I can't get you")
                continue
            #wolframalpha code here
            for i in range (2):
                if((i%2)==0):
                    try:
                        print("wolframalpha answer:")
                        app_id = "XXXXXXXXXXXXXXXXXX"
                        client = wolframalpha.Client(app_id)
                        res = client.query(my_input1)#input is passed into the wolframealpha
                        answer = next(res.results).text
                        if(answer==None):
                            print("No response")
                            speaker.Speak("no response")
                        else:
                            #text-2-speech output by the win32 pckg is done here
                            print(answer)
                            speaker.Speak(answer)
                    except Exception as e :
                        print("No response")
                        speaker.Speak("no response")
                        continue 
                if((i%2)==1):
                    print("wikipedia answer:")
                    try:
                        if(wikipedia.summary(my_input1,1)==None):
                            print("No response")
                            speaker.Speak("no response")
                        else:
                            print(wikipedia.summary(my_input1,1))
                            speaker.Speak(wikipedia.summary(my_input1,1))
                    except Exception as e:
                        print("No actual Input")
                        speaker.Speak("no response")
        print("***********************************")   
    else:
        print("you are not he recognized person to access me...")
        sys.exit()

