#!/usr/bin/env python
# coding: utf-8

#Bulk ILL uploader, Places resource sharing requests via API, requests sumbmitted as a reference manager XML file.
#Script outputs xlsx and csv files.
#Author Rachel Merrick.
#Originally written and executed with Anaconda/Jupyter:
#Some libraries may need to be installed before running the script in other environments.

#Output to indicate where the script is up to.
print('Importing libraries.')
#Import required libraries/modules.import requests as req #Used to to access Alma API.
import requests as req #Used to to access Alma API.
from requests.structures import CaseInsensitiveDict #used to create and send JSON request data object.
import xml.etree.ElementTree as ET #Used to parse XML file.
import pandas as pd #Used to create data frame.
import re #Allows use of regular expresions which have been used to extract error messages.
#from ipywidgets import IntProgress #Used for progress bar
#from IPython.display import display #Used for progress bar

#Function that tests if an element in an XML path is pressent and adds element or empty string to a list.
#List is returned to the main program.
def get_element(path, root):
    element_list = []
    for record in root.iter('record'):
        element = record.find(path)
        if element is not None:
            element = record.find(path).text
            element = element.replace('"','\\\"') #Add JSON escape character infront of double quotes.
        else:
            element = ""
        element_list.append(element)
    return(element_list)

#Function that returns a list of an XML element's attribute to the main program. 
def get_attribute(path, attribute_type, root):
    attribute_list = []
    for element in root.iter(path):
        attribute = element.get(attribute_type)
        attribute = attribute.replace('"','\\\"') #Add JSON escape character infront of double quotes.
        attribute_list.append(attribute)
    return(attribute_list)

#Main program.
def main():
    #Output to indicate where the script is up to.
    print('Script started.')
    
    #The following variables will need to be checked and added/edited before running the script.
    user_id = 'Enter primary ID' #Requester's Alma primary identifier
    file_name = 'Enter XML file name'#File name containing citation being requested
    
    #Variables used to send API request.
    key = ' ' #Alma production API key. add your requests key
    headers = CaseInsensitiveDict() #Used to send JSON portfolio object.
    headers["Content-Type"] = "application/json; charset=utf-8" #Used to send JSON portfolio object.
    url = ("https://api-ap.hosted.exlibrisgroup.com/almaws/v1/users/"+user_id + 
                "/resource-sharing-requests?override_blocks=false&apikey="+ key) #API request URL.

    #Variables used to detect and report errors and other messages.
    url_error = '<errorsExist>true</errorsExist>' #Test to indicate when an error is pressent in API request url.
    error_message = '<errorMessage>(.+)</errorMessage>' #Extracting the error message from API response.
    data_object_error = '"errorsExist":true' #Test to indicate when an error is pressent in JSON data object.
    #Test to indicate when an request ID is pressent in JSON data object.
    internal_request_id = '<request_id>(.+)</request_id>'

    #Output to indicate where the script is up to.
    print('Cleaning data.')
    
    #Read request file
    with open(file_name, 'r', encoding="utf8") as file :
        filedata = file.read()

    #Remove any XML style tags from EndNote libraries.
    filedata = filedata.replace('<style face="normal" font="default" size="100%">', '')
    filedata = filedata.replace('</style>', '')
    #Remove line breaks, carriage returns, spacing which shouldn't be sent in the JSON data object.
    filedata = filedata.replace('\n', '')
    filedata = filedata.replace('&#xD', '')
    filedata = filedata.replace('&#xA', '')
    filedata = filedata.replace('\r', '')
    filedata = filedata.replace('\t', '')
    filedata = filedata.replace('            ', '')

    # Save over request file with version with style tages removed.
    with open(file_name, 'w', encoding="utf8") as file:
        file.write(filedata)
        file.close()

    #Parse request file, assign tree and root variables.
    tree = ET.parse(file_name)
    root = tree.getroot()
    
    #Output to indicate where the script is up to.
    print('Parsing XML.')
    
    #Call function for required fields sending XML paths. Variables will contain a list of the relevant elements or
    #attributes.
    all_ref_types = get_attribute('ref-type', 'name', root)
    all_primary_titles = get_element('titles/title', root)
    all_volumes = get_element('volume', root)
    all_secondary_titles = get_element('titles/secondary-title', root)
    all_authors = get_element('contributors/authors/author', root)
    all_issues = get_element('number', root)
    all_issns = get_element('isbn', root)
    all_dates = get_element('dates/year', root)
    all_publisher = get_element('publisher', root)
    all_pages = get_element('pages', root)
    all_dois = get_element('electronic-resource-num', root)
    all_editions = get_element('edition', root)

    #Create data frame which will be used as output showing the API responses.
    df_data = {'Article_chapter_title' : all_primary_titles, 'Journal_book_title' : all_secondary_titles,
               'Author': all_authors ,'Volume': all_volumes, 'Issue': all_issues, 'Date': all_dates, 
               'Edition': all_editions ,'ISSN' : all_issns, 'Publisher' :all_publisher, 'Pages': all_pages, 
               'DOIs': all_dois,'Refernce_type': all_ref_types,'Responses': all_ref_types}
    df = pd.DataFrame(df_data, columns=['Article_chapter_title','Journal_book_title','Author','Volume', 'Issue','Date',
                                        'Edition','ISSN','Publisher','Pages','DOIs','Refernce_type','Responses'])

    #Output to indicate where the script is up to.
    print('Sending requests to Alma:')
    
    #Progress bar.
    #johns_progress_bar = IntProgress(min=0, max=len(all_ref_types)) # intantiate the bar
    #display(johns_progress_bar) # display the bar

    #Loop to iterate over each citation in the request file by using the attribute lists.
    for i in range(len(all_ref_types)):
        #If citation type is journal article, conference proceeding or book section.
        if (all_ref_types[i] == 'Journal Article' or all_ref_types[i] == 'Conference Proceedings'
           or all_ref_types[i] == 'Book Section' or all_ref_types[i] == 'book_section') :
            #assign citation being looped over's attributes to variable.
            primary_title = all_primary_titles[i]
            issn = all_issns[i]
            author = all_authors[i]
            year = all_dates[i]
            publisher = all_publisher[i]
            volume = all_volumes[i]
            secondary_title = all_secondary_titles[i]
            issue = all_issues[i]
            pages= all_pages[i]
            doi = all_dois[i]
            edition = all_editions[i]

            #Construct portion of JSON data object for journal articles and conference proceeings using variables
            #created by iterating over lists in this loop.
            if (all_ref_types[i] == 'Journal Article' or all_ref_types[i] == 'Conference Proceedings'):
                excerpt_data = '''"citation_type": {
        "value": "CR"
      },
      "level_of_service": {
        "value": "WHEN_CONVINIENT"
      },
      "title": "'''+ primary_title + '''",
      "issn": "''' + issn + '''",
      "author": "''' + author + '''",
      "year": "''' + year + '''",
      "publisher": "''' + publisher + '''",
      "volume": "''' + volume + '''",
      "journal_title": "''' + secondary_title + '''",
      "issue": "''' + issue + '''",
      "pages": "''' + pages + '''",
      "doi": "''' + doi + '''"
    }
    '''
            #Construct portion of JSON data object for book sections using variables created by iterating over 
            #lists in this loop.
            else:
                    excerpt_data = '''"citation_type": {
        "value": "BK"
      },
      "level_of_service": {
        "value": "WHEN_CONVINIENT"
      },
      "title": "'''+ secondary_title + '''",
      "isbn": "''' + issn + '''",
      "author": "''' + author + '''",
      "year": "''' + year + '''",
      "publisher": "''' + publisher + '''",
      "edition": "''' + edition + '''",
      "chapter_title": "''' + primary_title + '''",
      "pages": "''' + pages + '''",
      "doi": "''' + doi + '''"
    }
    '''
            #Create full JSON data object which includes a string with generic ILL request with the users ID and the
            #book or article portions contained in variables above.
            json_data = '''{
      "requester": {
        "value": "''' +user_id + '''"
      },
      "requested_media": "7",
      "format": {
        "value": "DIGITAL"
      },
      "preferred_send_method": {
        "value": "EMAIL"
      },

      "pickup_location_type": "LIBRARY",
      "pickup_location": {
        "value": "nsyd"
      },
      "copyright_status": {
        "value": "APPROVED"
      },
      "agree_to_copyright_terms": true,
      "willing_to_pay": false,
      ''' + excerpt_data

            #Send request using Post URL and data object to Alma. Reponse will be contained in the variable.
            response = req.post(url, headers=headers, data=json_data.encode('utf-8'))

            #Selection statement if there is an error add error message to notes column of data frame, 
            #else No errors detected.
            #Testing for API url error, if error pressent add error to responses column.
            if url_error in response.text:
                note = re.search(error_message, response.text)
                df.at[i, 'Responses'] = note.group(1)

            #Testing for data object error, if error pressent add error to responses column.
            elif data_object_error in response.text:
                note = response.text
                df.at[i, 'Responses'] = 'Error in data sent to API, error message: '+note

            #Testing for internal request ID (on all succesful requests) and adding it to responses column.
            elif bool(re.search(internal_request_id, response.text)) == True:
                note = re.search(internal_request_id, response.text)
                df.at[i, 'Responses'] = 'Request created. Internal request id: ' + note.group(1)

            #Else request not successfully placed and there is no error message, add uknown response to 
            #responses column   
            else:
                note = 'Unkown response'
                df.at[i, 'Responses'] = note

        #Not a supported reference type skip placing request and any other actions, add note to responses column.
        else:
            note = 'Unsupported reference type'
            df.at[i, 'Responses'] = note

        #Save to csv as progresses incase script is interupted.
        df.to_csv('Bulk_ill_request_response.csv')

        #Increment to progress bar.
        #johns_progress_bar.value += 1

    #Print message of how many records were processed and to wait for xlsx to be written to.
    print(str(i + 1) + " records processed, please wait for responses to be saved to xlsx file...")

    #Create dataframe with count for each response.
    df_updated = df.replace(to_replace ='Request created. Internal request id: (.+)', value = 'Request created'
                            , regex = True)
    counts = df_updated['Responses'].value_counts()

    # Create a Pandas Excel writer using XlsxWriter as the engine- imports another library.
    writer = pd.ExcelWriter('Bulk_ill_request_response.xlsx', engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    df.to_excel(writer, sheet_name='Detailed')
    counts.to_excel(writer, sheet_name='Counts')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    #print file save completeion messages.
    print("Response file saved as: Bulk_ill_request_response.xlsx.")

#Run main program.
if __name__ == "__main__":
    main()

