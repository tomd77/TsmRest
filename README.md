## TsmRest
Python class to call the REST API of IBM Spectrum Protect (TSM)      
The goal is to automate daily operations and give the TSM administrator a rest ^^

### Concept
The Spectrum Protect REST API has a method [issueConfirmedCommand](https://www.ibm.com/support/pages/ibm-spectrum-protect-operations-center-v810-rest-api-commands) 
that allows us to execute any TSM command.  
This TsmRest class uses that method.      
How do we use this class:

* Reactive monitoring and alerting  
* Backup job success reports, capacity invoicing, inventory auditing, ...
* Auto restart failed backup jobs for known return codes

### Requires
* Python 3 (tested on >=3.9.1)
* Spectrum Protect (SP) server and Operations Center (OC) >= version 8.1
* [openpyxl](https://openpyxl.readthedocs.io/en/stable/#installation) to generate Excel reports createReport()
 
### Example
We want to get the list of all TSM administrators (query admin || select * from admins)
```python
 from tsmrest import TsmRest

 # Set IP and PORT of your Spectrum Protect Operations Center
 my_api = TsmRest("tsmOC.company.com", "11090")

 # Run TSM command
 my_api.run_command("TSM01", "admin", "password", "query admin")
```
**run_command()** performs the REST call and parses the result. You end up with two class variables:
```python
 pprint(my_api.raw_result)    # = original REST response, structured by IBM
 pprint(my_api.parsed_result) # = parsed REST response, result of calling parse_raw_data()
```
You can target multiple TSM servers by using a list:
```python
 tsm_srv_list = ['TSM01', 'TSM02']
 my_api.run_command(tsm_srv_list, "admin", "password", "query admin")
```
Optional: The **parsed_result** can be used to create a report.
```python
 # Create a report
 my_api.create_report("XLSX", 'report1.xlsx', "System Admins Sheet", "349DCA")  # Excel
 my_api.create_report("CSV", 'report1.csv')                                     # CSV
 my_api.create_report("HTML", 'report1.html')                                   # HTML
``` 

**FYI**: **raw_result** vs **parsed_result**:
```python
 my_api.raw_result

 [[{  'hdr': [{'def': 'Administrator Name', 'id': '23296'},
             {'def': 'Days Since Last Access', 'id': '23294'},
             {'def': 'Days Since Password Set', 'id': '23309'},
             {'def': 'Locked?', 'id': '23297'},
             {'def': 'Privilege Classes', 'id': '23311'}],
    'items': [{'23294': 155,
              '23296': 'ADMIN',
              '23297': {'def': 'No', 'id': '23402'},
              '23309': 986,
              '23311': [{'val': {'def': 'System', 'id': '23440'}}]},
             {'23294': 928,
              '23296': 'ADMIN_CENTER',
              '23297': {'def': 'No', 'id': '23402'},
              '23309': 2919,
              '23311': []
             }]
 }]]


 my_api.parsed_result

 {  'hdr': ['TSM SERVER',
           'Administrator Name',
           'Days Since Last Access',
           'Days Since Password Set',
           'Locked?',
           'Privilege Classes'],
  'items': [{'Administrator Name': 'ADMIN',
            'Days Since Last Access': 155,
            'Days Since Password Set': 986,
            'Locked?': 'No',
            'Privilege Classes': 'System',
            'TSM SERVER': 'SERVER1'},
           {'Administrator Name': 'ADMIN_CENTER',
            'Days Since Last Access': 928,
            'Days Since Password Set': 2919,
            'Locked?': 'No',
            'Privilege Classes': [],
            'TSM SERVER': 'SERVER1'}],
 'srv': 'SERVER1'},
 'cmd': 'query admin a*'
```

### Project Roadmap
* Working on web app (frontend Javascript/Fetch, backend WSGI/Python) 
* Excel sheet 'Intro' with helpful links

### Ways to contribute
* Review code: Make more Pythonic. I'm self-learning Python with this project.  
* Bring new ideas to the table
* Jump in and help develop features that are on the roadmap

### License
Just want to give something back to the internet after having received so much.
Distributed under the MIT License.  Please do what you want with it.  

### Contact
Tom Desert  
tomdesert@hotmail.com

### References
- [Choose an open source license](https://choosealicense.com/)
- [Writing and formatting on Github](https://help.github.com/articles/getting-started-with-writing-and-formatting-on-github)
- [Markdown here cheatsheet](https://github.com/adam-p/markdown-here/wiki/Markdown-Here-Cheatsheet)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/#installation) and PHPExcel
- [TSM REST API doc](https://www.ibm.com/support/pages/ibm-spectrum-protect-operations-center-v810-rest-api-commands)
