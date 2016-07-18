# -*- coding: utf-8 -*-

__author__ = 'AR'

'''
******************************** SPAM Response Automation ********************************

Please follow the following guidelines to successfully use the SAPM Response Automation

1. Please make sure the spam@nbcuni.com mailbox is set as default mailbox in Outlook.exe 
	(if not the script will fetch unread emails from your default mailbox; leading to 
	 unexpected results)

2. Run the Spam_Response_Automation.py script just by double clicking it. Enjoy!!!!
3. To stop the program simply close the program window or kill python process (careful)


Technical Details:

Requirements:
Windows OS, python 2.x with packages win32com.client, BeautifulSoup

Functions:
1. sendResponse() - Function to send the responses to the user or to malware lab.
	Args:
	1. recipient -  Recipient of the message
	2. Subject -    Subject of the email
	3. attachment - Added .msg file if email is sent to malware lab

2. urlVoidRatingChecker() - Function to find URLvoid rating for domains
	Args:
	1. domainName - Domain name extracted from URL

	Returns 
	1 - If the domain is malicious 
	0 - If domain is not malicious

3. domainFormatter() - 	Function to find domains from the email body using 
						regular expression and calls urlVoidRatingChecker() to 
						vet domain.
	Args:
	1. msgPath - Path to .msg file locally stored on the system

	Returns 
	2 - If mail contains any suspicious attachment
	1 - Mail body contains malicious domains
	0 - Nothing malicious found

4. main() - Function to go through unread emails in SPAM mailbox and call 
			domainFormatter().
			It determines if the original email is sent as an attachment or 
			just forwarded. Based on the case it process the email and sends 
			response to user/ malware lab (with original email attached)

	Args/ Returns None


Please contact Amit Raut (amit.raut[at]nbcuni.com/ amitraut007[at]me.com) for help/ 
improvement. Thank you!!!!


******************************** SPAM Response Automation ********************************
'''

import win32com.client, os, re, urllib2, urllib, sys, time
from BeautifulSoup import BeautifulSoup
import datetime as dt


# Class to log stdout to logfile
class Tee(object):
    def __init__(self, *files):
        self.files = files
    def write(self, obj):
        for f in self.files:
            f.write(obj)


# Function to send the responses to the users or to malware lab
def sendResponse(recipient, recipient_name, subject, attachment):
	outlook = win32com.client.Dispatch("Outlook.Application")

	message = outlook.CreateItem(0)
	message.To = recipient
	message.Subject = 'RE: ' + subject

	# Sending the email for further analysis to the Malware Lab
	if recipient == 'spam.analysis@company-website.com':

		message.Body = '''
Please perform the analysis of the attached email message. 

Thank you!
Response Operation Center
		'''
	elif "Urgent Attention" in subject:
		message.body = '''
Hello,

Thank you for identifying this as a suspicious email. Your awareness and actions are important. You are helping to protect personal and company data while preventing malware from being installed on your computer. 
 
<Conpany_Name> Technology SAFE sent this reported email to you and your colleagues as a "Phishing teaching exercise."  Phishing emails may look legitimate or appear to come from well-known companies but can be dangerous because they are designed to gather your private information or gain access to company data. You did the right thing by sending this email to SPAM@company.com. 
 
Please continue to be diligent in reporting suspicious messages.

For more information on phishing and other tips to keep you cyber safe, visit our website at http://company-website.com/SAFE.

If you have any questions, please contact us.

Technology SAFE
http://company-website.com/SAFE


'''

	# Sending the SPAM resonse to the users
	else: 					
		message.body = '''
Hi,

Thank you for reporting this spam email to the Technology SAFE team. We understand the annoyance associated with spam emails like these, and our analysts are working on having the email blocked against future attempts.

If you have any questions, please contact the Technology SAFE Response Team at SAFE@company-website.com 

Please continue to report such emails to us. We thank you for your diligence and awareness.

Thanks,
SAFE Team
		'''

	if attachment != None:
		message.Attachments.Add(attachment)

	message.Send()

	if recipient == 'spam.analysis@nbcuni.com':
		print '\n{0:90} ==> {1:10}\n'.format('Response sent to ' + str(recipient_name), 'Analysis')
	else:
		print '\n{0:90} ==> {1:10}\n'.format('Response sent to ' + str(recipient_name), 'SPAM')



# URLVoid rating function
def urlVoidRatingChecker(domainName):
	
	try:
		opener = urllib2.build_opener( urllib2.HTTPHandler(), urllib2.HTTPSHandler(), urllib2.ProxyHandler({'http': 'http://proxy.inbcu.com:80'})) 
		urllib2.install_opener(opener)
		
		# Getting the updated report from URLvoid
		urllib2.urlopen('http://www.urlvoid.com/update-report/' + domainName)
		response = urllib2.urlopen('http://www.urlvoid.com/scan/' + domainName)
		html = response.read()

		# Processing HTML response from urlvoid
		parsed_html = BeautifulSoup(html)
		success = parsed_html.body.find('span', attrs={'class':'label label-success'})
		warning = parsed_html.body.find('span', attrs={'class':'label label-warning'}) 
		danger =  parsed_html.body.find('span', attrs={'class':'label label-danger'})

		result = success if success is not None else (warning if warning is not None else danger)
		print '{0:90} ==> {1:10.5}'.format('http://www.urlvoid.com/scan/' + domainName, result.text)
		#print parsed_html.body.find('span', attrs={'class':'label label-success'}).text
		if result == danger:
			return 1
		else: 
			return 0

	except Exception as e:
		# print "Error occurred", e

		print '{0:90} ==> {1:10}'.format('ERROR <URLVoid>:', e)

# Domain Formatter function
def domainFormatter(msgPath):

	outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
	message = outlook.OpenSharedItem(msgPath)

	count_attachments = message.Attachments.Count
	if count_attachments > 0:
	    for item in xrange(count_attachments):
	        
	        att_Name = message.Attachments.Item(item + 1).Filename.lower()

	        # If attachment contains original email
	        # Save mail temporarily in \AppData and vett the domains and attachments present 

	        # if mail contains any of the attachments with following extensions; Send the mail for analysis directly
	        if att_Name.endswith('.txt') or att_Name.endswith('.docm') or att_Name.endswith('.rtf') or \
	         att_Name.endswith('.dotm') or att_Name.endswith('.doc') or att_Name.endswith('.xlsx') or \
	         att_Name.endswith('.xltm') or att_Name.endswith('.xlsb') or att_Name.endswith('.xlam') or \
	         att_Name.endswith('.pptx') or att_Name.endswith('.pptm') or att_Name.endswith('.potm') or \
	         att_Name.endswith('.ppam') or att_Name.endswith('.ppsm') or att_Name.endswith('.sldm') or \
	         att_Name.endswith('.pdf') or att_Name.endswith('.html') or att_Name.endswith('.zip') or \
	         att_Name.endswith('.htm') or att_Name.endswith('.bin') or att_Name.endswith('.xls'):

	        	return 2
	        	# sendResponse() # Send the message for analysis
	        	break

	# Regular expression to find the URLs in the body of the email
	urlRegEx = re.compile(r'http[s]?://(?:[a-zA-Z-]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+')
	# urlRegEx = re.compile(r'((?<=[^a-zA-Z0-9])(?:https?\:\/\/|[a-zA-Z0-9-]{1,}\.{1}|\b)(?:\w{1,}\.{1}){1,5}(?:com|org|edu|gov|uk|net|ca|de|jp|fr|au|us|ru|ch|it|nl|se|no|es|mil|iq|io|ac|ly|sm|ar|io|in){1}(?:\/[a-zA-Z0-9]{1,})*)')
	urlList = urlRegEx.findall(message.Body.encode("utf-8"))

	# Getting valid domain names from URL for urlvoid analysis
	domainNameList = []
	for url in urlList:
		try:
			domainName = url.replace('www.', '').replace('http://', '').replace('https://', '')
			if '/' in domainName:
				domainName = domainName[:domainName.index('/')]
			
			if len(domainName) and domainName.lower() not in domainNameList:
				domainNameList.append(domainName.lower())
		
		except Exception as e:
			print '{0:90} ==> {1:10}'.format('ERROR: <Domain Formatter>', e)

	# print urlList, domainNameList
	if len(domainNameList):
		for domain in domainNameList:
			if urlVoidRatingChecker(domain) == 1:
				return 1
				break

	return 0
			# return urlVoidRatingChecker(domain)


def main():
	outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

	# "6" refers to the index of a folder
	inbox = outlook.GetDefaultFolder(6) 

	# Only processing the UNREAD emails from the SPAM mailbox Inbox folder
	messages = inbox.Items.Restrict("[UnRead]=True")
	responseCode = 0

	for message in reversed(messages):
		message.UnRead = False
		
		# Removing responses to Automatic replies and mails from SAFE, SPAM and ROC (Possible infinite loop)
		if 'Automatic reply'.lower() in str(message.subject.encode("ascii", "ignore")).lower() or \
		   'Response'.lower() in str(message.sender).lower() or \
		   'Spam'.lower() in str(message.sender).lower() or \
		   'SAFE'.lower() in str(message.sender).lower():
			break

		# print '{0:*^119s}'.format("*")
		responseCode = 0
		print 'Email Sender: ', str(message.sender) #,message.SenderEmailAddress
		print 'Email Subject: ', str(message.subject.encode("ascii", "ignore"))
		print '\n{0:90} ==> {1:10}'.format('URLvoid scanned URL', 'URLvoid Rating') 

		print

		# Check if the message is with original email attachment or forwarded email
		count_attachments = message.Attachments.Count
		analysis_done = False
		if count_attachments > 0:
			print 'INFO: Attachments found. Verifying now....\n'
			for item in xrange(count_attachments):
		        
				att_Name = message.Attachments.Item(item + 1).Filename.lower()

				# If attachment contains original email
				# Save mail temporarily in current working directory and vett the domains and attachments present 

				# if mail contains any of the attachments with following extensions; Send the mail for analysis directly
				if att_Name.endswith('.txt') or att_Name.endswith('.docm') or att_Name.endswith('.rtf') or \
				 att_Name.endswith('.dotm') or att_Name.endswith('.doc') or att_Name.endswith('.xlsx') or \
				 att_Name.endswith('.xltm') or att_Name.endswith('.xlsb') or att_Name.endswith('.xlam') or \
				 att_Name.endswith('.pptx') or att_Name.endswith('.pptm') or att_Name.endswith('.potm') or \
				 att_Name.endswith('.pdf') or att_Name.endswith('.ppsm') or att_Name.endswith('.sldm') or \
				 att_Name.endswith('.ppam') or att_Name.endswith('.html') or att_Name.endswith('.zip') or \
				 att_Name.endswith('.htm') or att_Name.endswith('.bin') or att_Name.endswith('.xls'):

				 	analysis_done = True
				 	responseCode = 1
					break

				# Original email sent for the analysis
				elif att_Name.endswith('.msg'): 
					message.Attachments.Item(item + 1).SaveASFile(os.getcwd() + '\\' + att_Name)
					# print 'Attachment saved at ' + os.getcwd()
					analysis_done = True

					if domainFormatter(os.getcwd() + '\\' + att_Name): # Based on result from URLVoid. Take Action
						responseCode = 1
					else:
						responseCode = 0
						
					# Removing the temporary saved email
					if os.path.isfile(os.getcwd() + '\\' + att_Name):
						os.remove(os.getcwd() + '\\' + att_Name)
					
					if responseCode != 0:
						break

		# Processing the forwarded emails
		if not analysis_done:
			urlRegEx = re.compile(r'http[s]?://(?:[a-zA-Z-]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+')
			# urlRegEx = re.compile(r'((?<=[^a-zA-Z0-9])(?:https?\:\/\/|[a-zA-Z0-9-]{1,}\.{1}|\b)(?:\w{1,}\.{1}){1,5}(?:com|org|edu|gov|uk|net|ca|de|jp|fr|au|us|ru|ch|it|nl|se|no|es|mil|iq|io|ac|ly|sm){1}(?:\/[a-zA-Z0-9]{1,})*)')
			urlList = urlRegEx.findall(message.Body.encode("utf-8"))

			# Getting valid domain names from URL for urlvoid analysis
			domainNameList = []
			for url in urlList:
				try:
					domainName = url.lower().replace('www.', '').replace('http://', '').replace('https://', '')
					if '/' in domainName:
						domainName = domainName[:domainName.index('/')]
					
					if len(domainName) and domainName.lower() not in domainNameList:
						domainNameList.append(domainName.lower())
			
				except Exception as e:
					print '{0:90} ==> {1:10}'.format('ERROR: <Domain Formatter>', e)

			# print urlList, domainNameList
			if len(domainNameList):
				for domain in domainNameList:
					if urlVoidRatingChecker(domain) == 1:
						responseCode = 1
						break

					else:
						responseCode = 0
		
		if responseCode:
			message.categories = 'Yellow Category (Sent for Analysis)'
			sendResponse('spam.analysis@company-website.com', 'spam.analysis@company-webiste.com', message.Subject, message) # Send the message for analysis	
		else:
			message.categories = 'Spam/Legit'
			sendResponse(message.SenderEmailAddress, message.sender, message.Subject, None) # message.SenderEmailAddress # Send standard response to the user

		print '{0:*^120s}'.format("*")
		time.sleep(0.1)

# Python Main Function
if __name__ == '__main__':

	f = open('logfile.txt', 'a')
	backup = sys.stdout
	sys.stdout = Tee(sys.stdout, f)

	banner = ' SPAM Response Automation ' + str(dt.datetime.now().date()) + ' '
	print '{0:*^120s}'.format(banner)
	print '\nThe program will continue to run forever unless you close this window or \nkill the python process (CAUTION: You may stop other python programs). \n\nThank you!\n\n'
	print '{0:*^120s}'.format("*")
	# Continue to run the program forever
	while True:
		main()

