import pyexcel as pe
import json

def readExcelFile():
    myActiveContact = {}
    my_array = pe.get_array(file_name="../activeContact.xlsx")
    
    #Remove the header
    del my_array[0]
    
    for record in my_array:
        myActiveContact[record[1]] = record[2]
        
    return myActiveContact

def getCustomers():
    myCustomers = [{
        "name" : "Acme Inc",
        "icn" : 50,
        "supportClientId" : 12345
        }, 
                   {
        "name" : "Jones",
        "icn" : 100,
        "supportClientId" : 23456
        }, 
                   {
        "name" : "Price",
        "icn" : 200,
        "supportClientId" : 34567
        }, 
                   ]
    return myCustomers

def __main__():
    print "This is main method"
    #Make an API call and get a list of Customers
    #ICN and SupportClientId
    customers = getCustomers()
    
    #Read from active contact spreadsheet
    activeContacts = readExcelFile()
    
    print customers
    print activeContacts
    
    for customer in customers:
        if customer["icn"] in activeContacts:
            print activeContacts[customer["icn"]] + " has support client id " + str(customer["supportClientId"]) + " for customer " + customer["name"]
        else:
            print customer["name"] + " has ICN " + str(customer["icn"]) + " not matching any customer ICN"

__main__()