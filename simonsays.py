# Sample code for speech recognition using the MS Speech API borrowed from Inigo Surguy (inigosurguy@hotmail.com)
# Modified for tv watching/media center by Rob Cherry simonsays@lxrb.com
# 
# Project hosted at http://code.google.com/p/simonsays/
#
# Next 2 lines should avoid having users run makepy on their own
from win32com.client import gencache
gencache.EnsureModule('{C866CA3A-32F7-11D2-9602-00C04F8EE628}', 0, 5, 4)
from win32com.client import constants
import win32com.client
import win32con,win32api,win32gui
import pythoncom

# For 3 word phrasing we need the last 2 words stored
backone = ""
backtwo = ""

class SpeechRecognition:
    """ Initialize the speech recognition with the passed in list of words """
    def __init__(self, wordsToAdd):
        # For text-to-speech
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.listener = win32com.client.Dispatch("SAPI.SpInProcRecognizer")
        self.listener.AudioInputStream = win32com.client.Dispatch("SAPI.SpMMAudioIn")
        self.context = self.listener.CreateRecoContext()
        self.grammar = self.context.CreateGrammar()
        # Do not allow free word recognition - only command and control
        # recognizing the words in the grammar only
        self.grammar.DictationSetState(1)
        # Create a new rule for the grammar, that is top level (so it begins
        # a recognition) and dynamic (ie we can change it at runtime)
        self.wordsRule = self.grammar.Rules.Add("wordsRule",
                       constants.SRATopLevel + constants.SRADynamic, 0)
        # Clear the rule (not necessary first time, but if we're changing it
        # dynamically then it's useful)
        self.wordsRule.Clear()
        # And go through the list of words, adding each to the rule
        [ self.wordsRule.InitialState.AddWordTransition(None, word) for word in wordsToAdd ]
        # Set the wordsRule to bse active
        self.grammar.Rules.Commit()
        self.grammar.CmdSetRuleState("wordsRule", 1)
        # Commit the changes to the grammar
        self.grammar.Rules.Commit()
        # And add an event handler that's called back when recognition occurs
        self.eventHandler = ContextEvents(self.context)
        # Announce we've started
        self.say("Started successfully")
    """Speak a word or phrase"""
    def say(self, phrase):
        self.speaker.Speak(phrase)


"""The callback class that handles the events raised by the speech object.
    See "Automation | SpSharedRecoContext (Events)" in the MS Speech SDK
    online help for documentation of the other events supported. """
class ContextEvents(win32com.client.getevents("SAPI.SpSharedRecoContext")):
    # Note that being here is no guarantee.  Any loud noise will force the speech recognition
    # into thinking that one of the words in the grammar has occured
    def OnRecognition(self, StreamNumber, StreamPosition, RecognitionType, Result):
        # newResult will be used to get recognised word text
        newResult = win32com.client.Dispatch(Result)
        word = newResult.PhraseInfo.GetText().lower()

        print("You said: ",word)


if __name__=='__main__':
    wordsToAdd = [ "code", "Hello code","Hello"]
    speechReco = SpeechRecognition(wordsToAdd)
    while 1:
        pythoncom.PumpWaitingMessages()



