# Spam Response Automation

'''
******************************** SPAM Response Automation ********************************

Please follow the following guidelines to successfully use the SAPM Response Automation

1. Please make sure the spam@nbcuni.com mailbox is set as default mailbox in Outlook.exe 
	(if not the script will fetch unread emails from your default mailbox; leading to 
	 unexpected results)

2. Run the Spam_Response_Automation.py script just by double clicking it. Enjoy!!!!
3. To stop the program simply close the program window or kill python process (careful)


#Technical Details:

#Requirements:
Windows OS, python 2.x with packages win32com.client, BeautifulSoup

#Functions:
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


Please contact Amit Raut amitraut007[at]me.com for help/ 
improvement. Thank you!!!!


******************************** SPAM Response Automation ********************************
'''
