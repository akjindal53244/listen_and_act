import nltk
import win32com.client
import speech_recognition as sr
import time
import pywinauto
import pythoncom
from win32com.client import gencache
gencache.EnsureModule('{C866CA3A-32F7-11D2-9602-00C04F8EE628}', 0, 5, 4)
from win32com.client import Dispatch
from win32com.client import constants
import win32con,win32api,win32gui
from subprocess import call

class SpeechRecognition:
	def __init__(self):
		self.sleeping=0
		self.wordsToAdd=["code wake up",
			"shut down",
			"what is time",
			"start Chrome",
			"close Dis",
			"close everything",
			"new tab",
			"close tab",
			"switch tab",
			"reload tab",
			"start incognito",
			"start Notepad",
			"hide Dis",
			"go to sleep",
			"show Chrome",
			"switch windows",
			"start explorer"]
		self.listener = win32com.client.Dispatch("SAPI.SpInProcRecognizer")
		self.listener.AudioInputStream = win32com.client.Dispatch("SAPI.SpMMAudioIn")
		self.context = self.listener.CreateRecoContext()
		self.grammar = self.context.CreateGrammar()
		self.grammar.DictationSetState(0)
		self.wordsRule = self.grammar.Rules.Add("wordsRule",
			constants.SRATopLevel + constants.SRADynamic, 0)
		self.wordsRule.Clear()
		[ self.wordsRule.InitialState.AddWordTransition(None, word) for word in self.wordsToAdd ]
		self.grammar.Rules.Commit()
		self.grammar.CmdSetRuleState("wordsRule", 1)
		self.grammar.Rules.Commit()
		self.eventHandler = ContextEvents(self.context)

class ContextEvents(win32com.client.getevents("SAPI.SpSharedRecoContext")):

	def OnRecognition(self, StreamNumber, StreamPosition, RecognitionType, Result):

		newResult = win32com.client.Dispatch(Result)
		word = newResult.PhraseInfo.GetText()
		if word=="code wake up":
			speechReco.sleeping=0
			mycode.speak_code("greet")
			return
		if word=="go to sleep":
			speechReco.sleeping=1
			mycode.speak_code("seeya")
			return

		if not speechReco.sleeping:
			mycode.process(word)
		
class Code:

	Ncodes=[]
	Nouns={}
	verbcodes=[]
	verbs={}
	strings={}
	treecode={}
	speakcodes=[]
	treecode['Q']=1
	treecode['S']=1
	treecode['NP']={}
	treecode['NP']['NP PNP']=0
	treecode['NP']['NP NN']=2
	treecode['NP']['DT NP']=2
	treecode['NP']['PNP']=0
	treecode['NP']['J NP']=2
	treecode['NP']['NN']=0
	treecode['PNP']=2
	treecode['DT']=0
	treecode['IN']=0
	treecode['VP']=1
	treecode['NN']=0
	treecode['J']=0	
	treecode['V']=0
	treecode['W']=0
	grammertext="""
			Q -> W V S | W S | S
			S -> NP VP | NP
			VP -> V NP | V
			NP -> NP PNP | NP NN | DT NP | PNP | J NP | NN

			PNP -> IN NP
			DT -> DT_L
			IN -> TO_L | IN_L | RB_L | RP_L
			NN -> PRP_L | NN_L | NNP_L | NNPS_L | NNS_L | VBG_L | CD_L

			J -> JJ_L | JJR_L | JJS_L


			V -> VB_L | VBN_L | VBD_L | VBP_L | VBZ_L | VBG_L | MD_L

			W -> WP_L | WDT_L | WRB_L
		"""
	wordsToAdd=["Code","Hello","Hello Code"]


	def __init__(self):
		#setting up parameters---------------------------------------------------------------
		self.range=6
		self.humor=int(self.range/2)
		self.mood=int(self.range/2)
		self.sarcasm=0

		#loading data-----------------------------------------------------------------------
		fin=open('data_speakcode','r')
		for code in fin:
			code=code.strip()
			self.speakcodes.append(code)
		fin.close()

		for code in self.speakcodes:
			self.strings[code]=[]
			fin=open("speakcode_lines/lines_"+code,'r')
			for line in fin:
				self.strings[code].append(line)
			fin.close()


		fin=open('data_Nouns','r')
		ctr=0
		for line in fin:
			line=line.split()
			self.Ncodes.append(line[0])
			for noun in line:
				self.Nouns[noun]=ctr
			ctr+=1
		fin.close()

		fin=open('data_verbs','r')
		ctr=0
		for line in fin:
			line=line.split()
			self.verbcodes.append(line[0])
			for verb in line:
				self.verbs[verb]=ctr
			ctr+=1
		fin.close()



		#initialising speech features-----------------------------------------------------
		
		self.pwa_app=pywinauto.application.Application()
		self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
		self.wsh=win32com.client.Dispatch("WScript.Shell")
		self.speak_code('greet')
		


	def process(self,sentence):
		print("started processing")
		#tagging the sentence-----------------------------------------------------------
		print(sentence)
		inputstring="you "+sentence
		tokenized_sent=nltk.word_tokenize(inputstring)
		tagged_sent=nltk.pos_tag(tokenized_sent)
		print(tagged_sent)


		wordsforgrammer={}
		for (word,tag) in tagged_sent:
			if tag in wordsforgrammer:
				wordsforgrammer[tag].append(word)
			else:
				wordsforgrammer[tag]=[word]
		for tag in wordsforgrammer:
			tagline=tag+"_L -> "
			for word in wordsforgrammer[tag]:
				tagline=tagline+" '"+word+"' | "
			self.grammertext=self.grammertext+"\n"+tagline


		#parsing input string-------------------------------------------------------
		grammer=nltk.CFG.fromstring(self.grammertext)
		parser=nltk.ChartParser(grammer)
		trees=parser.parse(inputstring.split())
		strtrees=[tree.label() for tree in trees]
		strtrees=str(strtrees)
		if strtrees=='[]':
			self.speak_code('didntget')
			return
		labels=[]
		newtrees=parser.parse(inputstring.split())
		for tree in newtrees:
			maintree=tree
			labels=[subtree.label() for subtree in tree.subtrees()]
			break
		
		#maintree
		command=self.trim_recurse(maintree)
		print(command)
		command=command.lower()
		commandtray=command.split()
		commandlen=len(commandtray)
		sflag=0
		if commandtray[0]=='you':
			sflag=1
		if sflag==1:
			if commandlen==2 and commandtray[1]=='code':
				self.speak_code('greet')
			else:
				self.speak_code('nod')
				if commandlen==3:
					try:
						nouncode=self.Ncodes[self.Nouns[commandtray[2]]]
					except KeyError:
						self.speak_code('cantdo')
						return
					fin=open("nouns/noun_"+nouncode+"_verbs",'r')
					try:
						verbcode=self.verbcodes[self.verbs[commandtray[1]]]
					except KeyError:
						self.speak_code('cantdo')
						return
					if verbcode+"\n" not in fin:
						self.speak_code('cantdo')
						return
					fin.close()
					fin=open("nouns/noun_"+nouncode+"_code",'r')
					while True:
						line = fin.readline()
						if not line: break
						if line==verbcode+"\n":
							codestrin=""
							codeline=fin.readline()
							while codeline!="---\n":
								codestrin+=codeline
								codeline=fin.readline()
							break
					fin.close()
					exec(codestrin)
		print("finished processing")
			
	def trim_recurse(self,tree):
		label=tree.label()
		if label[-1:]=="L":
			return ''.join(tree.leaves())
		returnstrin=''
		temptreecode=self.treecode[label]
		if label=='NP':
			subtreestrin=''
			for subtree in tree:
				subtreestrin=subtreestrin+" "+subtree.label()
			subtreestrin=subtreestrin.strip()
			temptreecode=self.treecode['NP'][subtreestrin]
		if temptreecode==0:
			for subtree in tree:
				returnstrin=self.trim_recurse(subtree)
				break
		if temptreecode==1:
			for subtree in tree:
				returnstrin=returnstrin+" "+self.trim_recurse(subtree)
		if temptreecode==2:
			for subtree in tree:
				testtree=subtree
			returnstrin=self.trim_recurse(testtree)
		return returnstrin.strip()

	def speak_string(self,line):
		self.speaker.Speak(line)

	def speak_code(self,speakcode):
		self.speak_string(self.strings[speakcode][self.mood])





	

global mycode
global speechReco
mycode=Code()
if __name__=='__main__':
	speechReco = SpeechRecognition()
	while 1:
		print("*")
		time.sleep(.1)
		pythoncom.PumpMessages()
