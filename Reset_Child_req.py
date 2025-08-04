#1. Connect to ALM, Login to project by choosing what project to login to (user input)
#2. Waits for user prompt to enter the requirement Id. 
#3. Script navigate to all the child requirement & reset the requirement. (status, revision, Signature etc)
#4. The script also capture the value of the fields Pre & Post resetting the fields in an excel file. (using Pandas)


#Revision number 0
#Status to Design

from comtypes.client import CreateObject
import pandas as pd
import xml.etree.ElementTree as ET
from collections import defaultdict
from openpyxl.workbook import Workbook

#credenttials to login to alm
url = "https://<host>/qcbin"
username =""
password =""
ota_connection = CreateObject("TDAPIOLE80.TDConnection")
domain = "<domain_name>"
#project = "NPI_Systems"
projects = {1: "Project1", 2: "Project2", 3: "project3", 4:"Project4", 5: "Project5", 6: "Project6"}
list_details = [] #Final list with dict values  for 'Project','Domain','List Name','List Items','Count/Size'


#function to login to alm
def login_to_alm():
    ota_connection.InitConnectionEx(url)
    ota_connection.Login(username, password)
    return ota_connection.LoggedIn

#function to connect to project
def connect_to_project(domain, project):
    try:
        ota_connection.Connect(domain, project)
        print("Logged into to Project " + project)                
    except Exception as e:
        print("Failed to login to Project " + project)
        print(e)
    Update_Req()

def Update_Req():
    Req_id = int(input("Please enter the Requirement ID, whose child req has to reset:") )
    Rec_fac = ota_connection.ReqFactory
    oChild = Rec_fac.GetChildrenList(Req_id)
    parentreqfact = ota_connection.ReqFactory.item(Req_id)
    print("Total child Requirement are:", oChild.count)
    parent_req_type = parentreqfact.Field( "RQ_TYPE_ID" ) 
    status = get_field_name("Status")     
    Revision_number =  get_field_name("Revision Number") 
    eSign = get_field_name("Reviewers/Signatures")
    Rejection_Reason = get_field_name("Rejection Reason")    
    Project_details_dict = {}          
    if parent_req_type.upper() != "Document".upper() :  
        old_status = parentreqfact.Field( status )  
        old_revision = parentreqfact.Field(Revision_number)  
        old_RR = parentreqfact.Field(Rejection_Reason)
        old_signatures = parentreqfact.Field(eSign) 
        parentreqfact.Field[status] = "Design"
        parentreqfact.Field[Revision_number] = 0
        parentreqfact.Field[Rejection_Reason] = " "

        parentreqfact.Field[eSign] = " "        
        parentreqfact.Post()
        Project_details_dict["Requirement Name"] = parentreqfact.Name
        Project_details_dict["Requirement ID"] = parentreqfact.ID
        Project_details_dict["Immediate Parent Req ID"] = "NA"
        Project_details_dict["Old Status Value"] = old_status
        Project_details_dict["Old Revision Value"] = old_revision
        Project_details_dict["Old Rejection_Reason Value"] = old_RR
        Project_details_dict["Old eSign Value"] = old_signatures
        Project_details_dict["New Status Value"] = parentreqfact.Field( status ) 
        Project_details_dict["New Revision Value"] = parentreqfact.Field(Revision_number)
        Project_details_dict["New Rejection_Reason Value"] = parentreqfact.Field(Rejection_Reason)
        Project_details_dict["New eSign Value"] = parentreqfact.Field(eSign) 
        list_details.append(Project_details_dict)    
      
    for child in oChild:
        update_child_Req(child.id,parentreqfact.ID)
    #print("Final list:",list_details) 
    filename_xlsx = project + '_' + str(parentreqfact.ID) + '_1' + parentreqfact.Name + '.xlsx'
    print("Filename:",filename_xlsx)
    save_to_excel(list_details,filename_xlsx) 

def update_child_Req(Child_req_id,father_ID):    
    temprFact = ota_connection.ReqFactory
    child_req = temprFact.Item(Child_req_id)
    status = get_field_name("Status") 
    Revision_number =  get_field_name("Revision Number") 
    eSign = get_field_name("Reviewers/Signatures")
    Rejection_Reason = get_field_name("Rejection Reason")    
    Project_details_dict = {}
    if child_req.Field( status ) not in ["Design", "Obsolete", "Postponed"]:
        old_status = child_req.Field( status )  
        old_revision = child_req.Field(Revision_number)  
        old_RR = child_req.Field(Rejection_Reason)
        old_signatures = child_req.Field(eSign)      
        child_req.Field[status] = "Design"
        child_req.Field[Revision_number] = 0
        child_req.Field[Rejection_Reason] = " "
        child_req.Field[eSign] = " "
        child_req.Post()
        Project_details_dict["Requirement Name"] = child_req.Name
        Project_details_dict["Requirement ID"] = child_req.ID
        Project_details_dict["Immediate Parent Req ID"] = father_ID
        Project_details_dict["Old Status Value"] = old_status
        Project_details_dict["Old Revision Value"] = old_revision
        Project_details_dict["Old Rejection_Reason Value"] = old_RR
        Project_details_dict["Old eSign Value"] = old_signatures
        Project_details_dict["New Status Value"] = child_req.Field(status)
        Project_details_dict["New Revision Value"] = child_req.Field(Revision_number)
        Project_details_dict["New Rejection_Reason Value"] = child_req.Field(Rejection_Reason)
        Project_details_dict["New eSign Value"] = child_req.Field(eSign) 
        list_details.append(Project_details_dict)
    sub_child = temprFact.GetChildrenList(child_req.ID)          
    if sub_child.count > 0: 
        for  child in sub_child:
            update_child_Req(child.id,child_req.ID)

def save_to_excel(list_details,filename_xlsx):
    df = pd.DataFrame(list_details)
    df.to_excel(filename_xlsx,index=False) 
    print(df.head())
    print("List details saved to excel")

#function to logout from alm
def logout_from_alm():
    ota_connection.Logout()
    ota_connection.Disconnect()
    print("Logged out from ALM")    

#function to get physical name of the field
def get_field_name(Fieldname):
    fieldlist = ota_connection.fields("Req")
    for field in fieldlist:
        fieldprop = field.Property
        if fieldprop.userlabel == Fieldname:
            return fieldprop.dbcolumnname
        
if __name__ == "__main__":   
    login_result = login_to_alm()    
    if login_result == True:
        print("Connected to ALM") 
        print("Project List", projects)
        key_id = int(input("Enter the Id of the project you want to login to from the above list of projects: "))
        if key_id > 7 or key_id < 1 :
            print("Invalid input, Please select the project from the given range")
            print("Project List", projects)
            key_id = int(input("Enter the Id of the project you want to login to from the above list of projects: "))
        project = projects[key_id]   
        print(project)            
        connect_to_project(domain, project)
    else:
        print("Not Connected to ALM")
    logout_from_alm()
