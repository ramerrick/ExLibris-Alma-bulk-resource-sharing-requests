# ExLibris-Alma-bulk-resource-sharing-requests

<b>Brief Description:</b> The script places resource sharing (ILL) requests in bulk in Alma from an exported XML file from a reference manager.

<b>Long description/background:</b>
The script parses over an XML file exported from a reference manager (Zotero, EndNote and Mendeley files have been tested and work, others may also work). Elements and attributes that are required  to request a interlibrary loan are added to lists. A loop is then used to add the elements from the lists into a JSON object that is then sent with a POST call to the Alma user API to place the requests. The API response is then added to a dataframe which contains each citation's details. The dataframe is then saved as an excel file to be sent to the requester.

A pdf with local instructions on running the script has also been included in the repository. 

<b>API:</b> Alma, users, read/write</br>
Originally written and executed with Anaconda/Jupyter</br>
<b>Output files:</b> two output files are created Bulk_ill_request_response.xlsx and Bulk_ill_request_response.csv. The csv file is written as each request is placed. This file is created in case the script is suddenly interrupted and halted.

<b>Modules/Libraries used:</b>
<ul>
<li> Requests (for API communication)</li>
<li>Requests.structures import CaseInsensitiveDict (used to create and send JSON data objects)</li>
<li>Pandas (using data frames)</li>
<li>Re (regular expressions)</li>
<li>Datetime (current day/time, used to time)</li>
<li>xml.etree.ElementTree (used to parse XML file)</li>
<li>Xlsxwriter (used to create xlsx file)</li>
<li>ipywidgets import IntProgress (used for progress bar)</li>
<li>Python.display import display (used for progress bar)</li>
</ul>	
<b>Includes:</b>
<ul>
<li>Timing the script from start to finish</li>
<li>Importing an XML file</li>
<li>Parsing XML required elements into lists</li>
<li>Parsing XML required attributes into lists</li>
<li>Creating a data frame from lists</li>
<li>Iterating over the list using a for loop, values contained in the lists used to build individual JSON data objects for each fine to be sent to the API</li>
<li>Post API requests with JSON object and write response to variable
<li>Selection statement, look for error notifier with regular expressions:</li>
		-If there is error: Look for error message in response using regular expressions and contain the message, add it to variable, If there is no error add success message to variable
<li>Add either success or error message (contained in variable) to the notes column of the data frame</li>
<li>Write new data frame with updated notes column with error and success messages to xlsx and csv files</li>
<li>Create a data frame containing the number of each response and add it to sheet two of the spreadsheet.</li>
<li>Display records processed.</li>
</ul>		
	
<b>Resources Used:</b></br>
API console https://developers.exlibrisgroup.com/console/</br>
Rest User Resource Sharing Request: https://developers.exlibrisgroup.com/alma/apis/docs/xsd/rest_user_resource_sharing_request.xsd/?tags=POST (What can be added in JSON object for Resource sharing requests)</br>
Post method with JSON data object https://stackoverflow.com/questions/9733638/how-to-post-json-data-with-python-requests</br>
Pandas cheat sheet: http://datacamp-community-prod.s3.amazonaws.com/f04456d7-8e61-482f-9cc9-da6f7f25fc9b</br>
Regular Expressions Python Capture groups: https://stackoverflow.com/questions/48719537/capture-groups-with-regular-expression-python
Find substring in string http://net-informations.com/python/basics/contains.htm#:~:text=Using%20Python's%20%22in%22%20operator,%2C%20otherwise%2C%20it%20returns%20false%20.</br>
Update dataframes: https://www.askpython.com/python-modules/pandas/update-the-value-of-a-row-dataframe?fbclid=IwAR3o7WAnpPYk6q7PmyMcpmpW8F9F4EVzV3GR-KoUAdmGPOUaIkELYoZOAhA</br>
Create dataframes: https://www.geeksforgeeks.org/create-a-pandas-dataframe-from-lists/</br>
Parse XML in python: https://docs.python.org/3/library/xml.etree.elementtree.html![image](https://user-images.githubusercontent.com/119052194/207497990-07b012dc-3e1c-4784-b413-638e18db99ba.png)
