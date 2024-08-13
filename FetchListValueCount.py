#import ota api for microfocus alm to use the tdconnection object
from comtypes.client import CreateObject
import pandas as pd
import xml.etree.ElementTree as ET
from collections import defaultdict
from openpyxl.workbook import Workbook

#credenttials to login to alm
url = "http://dev-testalm.cytiva.net/qcbin"
username ="saira.banu1"
password ="Welcome@72472_SR"
#ota connection object
ota_connection = CreateObject("TDAPIOLE80.TDConnection")
ota_sa_connection = CreateObject("SAClient.SaApi.9")
#Template and Field details, this can also be used as input parameters if you want to resuse the script for multiple templates/Fields
Template_name = 'eSignature11_CR21'
filename_xml = 'data_' + Template_name + '.xml'
filename_xlsx = 'data_' + Template_name + '.xlsx'
list_name = ['Severity','Priority'] #ALM List name, saved in a list to use for multiple lists
project_list = defaultdict(list)
list_details = [] #Final list with dict values  for 'Project','Domain','List Name','List Items','Count/Size'

#function to login to alm as Siteadmin to fetch the projects linked to a template
def Connect_to_almSA():
    try:
        ota_sa_connection.login(url,username, password)
        print("Logged into to SA ")
    except Exception as e:
        print("Failed to login to SA ")
        print(e)  
    #get all projects linked to a template   
    data = ota_sa_connection.getlinkedProjects('ALMESIGNATURE',Template_name,'Template')
    ota_sa_connection.logout()   
    print("logged out from SA")    
    #saving the data to xml file
    with open(filename_xml, 'w') as f:
        f.write(data)
    
#function to login to alm
def login_to_alm():
    ota_connection.InitConnectionEx(url)
    ota_connection.Login(username, password)
    return ota_connection.LoggedIn

def Connect_to_all_templateProjects():
    tree = ET.parse(filename_xml) 
    root = tree.getroot()
   
    for item in root.findall('TDXItem'): 
        project_list[item.find('DOMAIN_NAME').text].append(item.find('PROJECT_NAME').text)
        ota_connection.Connect(item.find('DOMAIN_NAME').text, item.find('PROJECT_NAME').text)        
        get_list_details(item.find('DOMAIN_NAME').text, item.find('PROJECT_NAME').text)
        #list_details.append(details)
        
    save_to_excel(list_details)               

def get_list_details(Domain, Project):
    print('Connection to Project ' + Project + ' in Domain ' + Domain)   
    
     #List to store the list items/values
    for item in list_name:
        print(item)
        List_items = []
        Project_details_dict = {} #to store the project details that will be passed to list_details list
        customizations = ota_connection.Customization   
        custlists1 = customizations.Lists
        custlist1 = custlists1.List(item)
        listrootnode1 = custlist1.RootNode
        childnode = listrootnode1.Children            
        for values in childnode:
            List_items.append(values.Name)              
        Project_details_dict['Domain Name'] = Domain
        Project_details_dict['Project Name'] = Project
        Project_details_dict['List Name'] = item
        Project_details_dict['Count/Size'] = childnode.count
        Project_details_dict['List Items'] = List_items               
        list_details.append(Project_details_dict) 
    
def save_to_excel(list_details):
    df = pd.DataFrame(list_details)
    df.to_excel(filename_xlsx,index=False) 
    #print(df.head())
    print("List details saved to excel")

#function to logout from alm
def logout_from_alm():
    ota_connection.Logout()
    ota_connection.Disconnect()
    print("Logged out from ALM") 

if __name__ == "__main__":   
    Connect_to_almSA()
    login_result = login_to_alm()
    if login_result == True:
        print("Connected to ALM")        
    else:
        print("Not Connected to ALM")
    Connect_to_all_templateProjects()
    logout_from_alm()
