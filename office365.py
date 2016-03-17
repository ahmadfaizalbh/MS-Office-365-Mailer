import datetime
import logging
import urllib2
import urllib
import json
import time
import os


Error_file = "DB_Error" + str(datetime.datetime.now()) + ".log"
logging.basicConfig(filename=Error_file,
                    level=logging.DEBUG,
                    format='%(asctime)s %(levelname)s %(name)s %(message)s')
logger=logging.getLogger(__name__)


class MSOFileHandler:
	'''Attachment File handler'''

	def __init__(self, def_read_dir="", def_write_dir=""):
		'''Default Initializer function'''

		self.default_read_dir = def_read_dir + ("" if def_read_dir.\
		                        endswith("/") else "/")
		self.default_write_dir = def_write_dir + ("" if def_write_dir.\
		                         endswith("/") else "/")

	def create_file(self, MSO_dict, Dir=None):
		'''
		Attachment name and content is read and create's new file in local 
		system
		''' 

		if ("Name" in MSO_dict) and ("ContentBytes" in MSO_dict):
			f = open(( (Dir + ("" if Dir.endswith("/") else "/")) if Dir else \
			self.default_write_dir)  +MSO_dict["Name"],"wb")
			f.write(MSO_dict["ContentBytes"].decode('base64'))
			f.close()
			
	def Create_Attachment(self,File_name):
		'''
		Takes File name (with absolute or complete path) and returns 
		a dictionary of attributes needed for attachment of file in Microsoft
		mail
		'''

		MSO_dict = {"@odata.type": "#Microsoft.OutlookServices.FileAttachment"}
		F_path = ("" if File_name.startswith("/") else self.default_read_dir)\
		         + File_name
		f = open(F_path,"rb")
		MSO_dict["Name"] = f.name.split("/")[-1]
		MSO_dict["ContentBytes"] = f.read().encode('base64')
		MSO_dict["ContentType"] = mimetypes.types_map['.' \
	                              + MSO_dict["Name"].split(".")[-1]]
		(mode, ino, dev, nlink, uid, gid, size, atime, mtime, ctime) = \
	    os.stat(f.name)
		f.close()
		MSO_dict["DateTimeLastModified"] = time.strftime("%Y-%m-%dT%H:%M:%SZ",
		                                                 time.gmtime(mtime))
		return MSO_dict

class MSOffice365:
	'''
	A microsoft 365 Outlook access class
	'''	

	def __init__(self,mail_box=None):
		'''
		Default Initializer function
		'''
		
		from Settings import MSO365_Credentials as MSO365
		from Settings import Default_File_locations as DFL
		
		password_mgr = urllib2.HTTPPasswordMgrWithDefaultRealm()
		self.MailBox_id = MSO365['mail_box_ID'] if not mail_box else mail_box
		self.top_level_url = \
	    "https://outlook.office365.com/api/v1.0/Users('%s')" % self.MailBox_id
		password_mgr.add_password(None,
		                          self.top_level_url,
		                          MSO365['username'],
		                          MSO365['password'])
		handler = urllib2.HTTPBasicAuthHandler(password_mgr)
		self.opener = urllib2.build_opener(handler)
		self.FileHandler = MSOFileHandler(def_read_dir=DFL['Read_From_Dir'],
		                                  def_write_dir=DFL['Create_in_Dir'])

	@property
	def DisplayName(self):
	    """
	    """
	    
		if "profile" not in self.__dict__:
			try:
				self.profile = self.open("/")
			except Exception, e:
				self.profile = {}
		try:
			return self.profile["DisplayName"]	
		except Exception, e:
			return self.MailBox_id

	@property
	def Alias(self):
	    """
	    """
	    
		if "profile" not in self.__dict__:
			try:
				self.profile = self.open("/")
			except Exception, e:
				self.profile = {}
		try:
			return self.profile["Alias"]	
		except Exception, e:
			return self.MailBox_id.split("@")[0]

	def open(self,url):
	    """
	    """
	    
		response=json.load(self.opener.open(self.top_level_url + url))
		self.next_url = response['@odata.nextLink'] if '@odata.nextLink' in \
                        response else ""
		return response

	def next(self):
	    """
	    """
	    
		if self.next_url:
			response = json.load(self.opener.open(self.next_url))
			self.next_url = response['@odata.nextLink'] if '@odata.nextLink' \
			                in response else ""
			return response
		return {"error":{
		            "code":"ErrorInvalidUrlfield",
		            "message":"Invalid Url."
		            }
	           }

	def buildQuery(self, url="", q=None):
	    """
	    """
	    
		if q:
			fieldSep = "?"
			if type(q) == dict:
				for i in q:
					url += fieldSep + "$" + i + "="
					if type(q[i]) == list:
						sep=""
						for j in q[i]:
							url += sep + urllib.quote_plus(unicode(j).\
							       encode('utf8'), safe='/')
							sep = ","
					elif type(q[i]) == str:
						url += urllib.quote_plus(unicode(q[i]).encode('utf8'),
						                                 safe='/')
					else:
						raise ValueError("Invalid Argument Syntax")
					fieldSep = "&"
			elif type(q) == str:
				url += fieldSep + q
			else:
				raise ValueError("Invalid Argument Syntax")
		return url

	def Messages(self,q=None,mail_id=None, Folder_id=None):
	    """
	    """
	    
		url = (("/Folders('" + Folder_id + "')") if Folder_id else "") \
		      + "/Messages" + (("""('""" + mail_id + """')/""") if mail_id \
		      else "/")
		url = self.buildQuery(url, q)
		return self.open(url)

	def Post(self, url, json_data, fullurl=False):
	    """
	    """
	    
		headers = { 'X_REQUESTED_WITH' :'XMLHttpRequest',
        		    'ACCEPT': 'application/json, text/javascript, */*; q=0.01',
                    'Contentlength':len(json_data)}

		request  = urllib2.Request(url if fullurl else (self.top_level_url + \
		                           url),
		                           data=json_data,
		                           headers=headers)
		request.add_header('Content-Type', 'application/json')
		request.get_method = lambda: "POST"
		try:
			connection = self.opener.open(request)
		except urllib2.HTTPError, e:
			connection = e
		return connection.read()

	def Sendmail(self, Subject="Have you seen this new Mail REST API?",
	             Importance="High", Body=None, ToRecipients=None, Attachments=[],
	             SaveToSentItems=True):
		''' sendmail(
				Subject="Have you seen this new Mail REST API?",    
  				Importance="High",    
  				Body={"ContentType": "HTML",
  				      "Content": "It looks awesome!<br/> This is test mail" },    
  				ToRecipients=[{	"EmailAddress": {
  				                            "Name": "Your Name", 
  				                            "Address": "username@company.com"
  				                            }
                              }],
				Attachments=[list of file names],
				SaveToSentItems=True
			)
		'''	
			
		message_data={    
  				"Subject": Subject,    
  				"Importance": Importance,    
  				"Body": Body if Body else { 
  				    "ContentType": "HTML",
                    "Content": "It looks awesome!<br/> This is test mail" 
                    },    
  				"ToRecipients": ToRecipients if ToRecipients else [{
  				    "EmailAddress": {
  				        "Name": self.DisplayName, 
  				        "Address": self.MailBox_id	
  				        }
			        }]				
				}
		if Attachments:
			message_data["Attachments"] = []
			for File_name in Attachments:
				message_data["Attachments"].append(self.FileHandler.\
				Create_Attachment(File_name))
		json_data = json.dumps({  
			"Message": message_data,
			"SaveToSentItems": SaveToSentItems
		})
		return self.Post("/sendmail", json_data)

	def CreateDraftMessage(self, Folder_id='inbox',
	                       Subject="Have you seen this new Mail REST API?",
	                       Importance="High", Body=None, ToRecipients=None,
	                       Attachments=[]):
        """
        """
        
		url = "/folders('" + Folder_id + "')/messages"
		message_data = {    
  				"Subject": Subject,    
  				"Importance": Importance,    
  				"Body": Body if Body else {
  				        "ContentType": "HTML",
                        "Content": "It looks awesome!<br/> This is test mail"
                        },    
  				"ToRecipients":ToRecipients if ToRecipients else [{
  				    "EmailAddress": {
  				        "Name": self.DisplayName,
  				        "Address":  self.MailBox_id	
  				        }
			        }]
				}
		if Attachments:
			message_data["Attachments"] = []
			for File_name in Attachments:
				message_data["Attachments"].append(self.FileHandler.\
				Create_Attachment(File_name))
		json_data=json.dumps(message_data)
		return self.Post(url, json_data)

	def CreateFolder(self, Folder_id, DisplayName):
	    """
	    """

		url = "/Folders('" + Folder_id + "')/childfolders"
		json_data = json.dumps({
		  "DisplayName": DisplayName
		})
		return self.Post(url, json_data)

	def CreateContact(self, GivenName="Your Name", EmailAddresses=[],
	                  BusinessPhones=[]):
		"""CreateContact(
			GivenName = "Your Name",
			EmailAddresses = [{
			                    "Address":"username@company.com",
			                    "Name":"Your Name"
		                      }],
			BusinessPhones = ["123-456-7890"])
		"""

		json_data = json.dumps({
			"GivenName": GivenName,
			"EmailAddresses": EmailAddresses ,
			"BusinessPhones": BusinessPhones
		}) 
		return self.Post("/Contacts", json_data)

	def Folders(self, Folder_id=None, q=None):
	    """
	    """

		url = "/Folders" + (("""('""" + Folder_id + """')/""") if Folder_id \
              else "/")
		url = self.buildQuery(url, q)
		return self.open(url)
	
	def Calendars(self, Calender_id=None, q=None):
	    """
	    """

		url = "/Calendars" + (("""('""" + Calender_id + """')/""") \
		      if Calender_id else "/")	
		url = self.buildQuery(url, q)
		return self.open(url)		

	def CalendarGroups(self, CalGroup_id=None, q=None):
	    """
	    """

		url = "/CalendarGroups" + (("""('""" + CalGroup_id + """')/""") \
		      if CalGroup_id else "/")
		url = self.buildQuery(url, q)	
		return self.open(url)	

	def Events(self, Event_id=None, q=None):
	    """
	    """

		url = "/Events"+(("""('""" + Event_id + """')/""") if Event_id \
		      else "/")
		url = self.buildQuery(url, q)	
		return self.open(url)	

	def Contacts(self, Contact_id=None, Folder_id=None, q=None):
	    """
	    """

		url = (("/Contactfolders('" + Folder_id + "')") if Folder_id else "") \
		      + "/Contacts" + (("""('""" + Contact_id + """')/""") \
		      if Contact_id else "/")	
		url = self.buildQuery(url, q)
		return self.open(url)

	def ContactFolders(self, Contact_id=None, q=None):
	    ""
	    ""

		url = "/Contactfolders" + (("""('""" + Contact_id + """')/""") \
		      if Contact_id else "/")	
		url = buildQuery(url, q)
		return self.open(url)

	def Attachments(self, mail_id, Attachment_id=None, q=None, Dir=None,
	                Create_file=False):
        """
        """

		url = "/Messages('" + mail_id + "')/attachments" + (("('" + \
		      Attachment_id + "')/") if Attachment_id else "/" )
		url = self.buildQuery(url,q)
		MSO_dict = self.open(url)
		if Create_file:
			if Attachment_id:
				self.FileHandler.create_file(MSO_dict, Dir=Dir)
			else:
				for i in MSO_dict["value"]:
					self.FileHandler.create_file(i, Dir=Dir)
		return MSO_dict


