
#import ota api for microfocus alm to use the tdconnection object
from comtypes.client import CreateObject
from datetime import date
#import response

#credenttials to login to alm
url = "http://<host>/qcbin"
username =""
password =""
#ota connection object
#ota_connection = CreateObject("TDApiOle80.TDConnection")
ota_connection = CreateObject("TDAPIOLE80.TDConnection")
#project list
project_list =['project1','project2','...']
domain = "QUALITY"

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
        create_defect()        
    except Exception as e:
        print("Failed to login to Project " + project)
        print(e)

#function to create defect in ALM
def create_defect():
    i = 0
    bug_factory = ota_connection.BugFactory
    #new_bug = bug_factory.AddItem(Null)
    thebug = bug_factory.item(1) 
    Discription = get_field_name("Description")     
    status = get_field_name("Approval Status")
    detected_date = get_field_name("Detected on Date")
    detected_by = get_field_name("Detected By")
    environment = get_field_name("Environment")
    summary = get_field_name("Summary")    
    detected_Verion = get_field_name("Detected in Version/Release")
    workaround = get_field_name("Workaround/Corrective Action")
    RootCause = get_field_name("Root Cause")
    Impact = get_field_name("Impact")
    
    while i <= 4:
        thenew_bug = bug_factory.AddItem(None)
        thenew_bug.field[summary] = "This is a test defect"
        thenew_bug.field[Discription] = "This is a test defect"
        thenew_bug.field[status] = "New"
        thenew_bug.field[detected_date] = date.today().strftime("%m/%d/%Y")
        thenew_bug.field[detected_Verion] = "Test"
        thenew_bug.field[environment] = "Other"
        thenew_bug.field[detected_by] = "<username>"
        thenew_bug.field[Impact] = "No impact"
        #I used the below code as value to impact field. as I was testing the field size with 1k charater
        #response.fetch_text_from_url("https://admhelp.microfocus.com/alm/en/24.1/online_help/Content/Tutorial/sa_defect_add.htm")

        thenew_bug.Post()
        i += 1
    print("5 new defect is created")
    thenew_bug = None

#function to logout from alm
def logout_from_alm():
    ota_connection.Logout()
    ota_connection.Disconnect()
    print("Logged out from ALM")    

#function to get physical name of the field
def get_field_name(Fieldname):
    fieldlist = ota_connection.fields("Bug")
    for field in fieldlist:
        fieldprop = field.Property
        if fieldprop.userlabel == Fieldname:
            return fieldprop.dbcolumnname


if __name__ == "__main__":   
    login_result = login_to_alm()
    if login_result == True:
        print("Connected to ALM")
        for project in project_list:
            connect_to_project(domain, project)
    else:
        print("Not Connected to ALM")
    logout_from_alm()
