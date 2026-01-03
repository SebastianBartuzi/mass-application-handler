# mass-application-handler

This application allows to automate the process of sending question requests to local governments in Poland, as well as
creating analysis reports.

## Requests sending
- With the request template as .docx provided, the app creates 2,500 copies of the template each filled with a
local authority contact data.
- The app sends 2,500 e-mails via Microsoft Outlook to each local authority.

## Report creation
- The app goes through the Microsoft Outlook inbox e-mails until it finds an e-mail with a STOP flag.
- The app extracts e-mail text and attachments, sends them with a proper prompt to Gemini LLM model in order to extract
desired information. The authorities are differentiated by a unique TERYT code, which unless provided by the authority,
is extracted by the model as well.
- After extracting the information, the app will send an e-mail with a STOP flag to the responses inbox.
- With extracted information, the app creates a report and sends it to the inboxes of users who should receive them.
Additionally, an operation summary is sent - number of responses received, problems encountered, responses worth human's
attention (e.g. the authority refusing to give the information).

## Manual setup
Currently the application is unfriendly for non-developer users and MUST be manually adjusted by a developer in order 
to work for a new application template.

### Requests sending
1. Install dependencies from pyproject.toml
2. Check assets folder:
* * replace application.docx (leave under the same name) with the application template desired;
* * update any local authorities data in data.xlsx (governors change throughout the term, so it is advisable to have this
data up-to-date)
3. Ensure addressees in .env are correct
4. Ensure Outlook open with the main account set as the inbox from which applications will be sent
4. Run send.py to send e-mails

### Report creation
1. Install dependencies from pyproject.toml
2. Replace assets/report_template.xlsx with the desired report template
3. Replace prompt.txt and response_schema.json in assets/mail_analysis in order to extract desired information
4. Edit code in models/response_mail_model.py, services/excel_service.py and services/response_service.py to cooperate
with information defined in response_schema
5. Ensure report columns, LLM parameters and addressees in .env are correct
6. Ensure Outlook open with the main account set as the inbox where responses are received
7. Run read.py to perform an analysis

## Future works
Tasks to do in order to adjust the app to be friendly for non-developers:
* Build a frontend application that allows to send a .docx template of a request and define other necessary parameters
(such as LLM model, e-mail addresses where operation summary should be sent, etc.).
* Introduce a validator checking if the template has all required placeholders.
* Build a tool that automatically looks up governors changes using Gemini with grounding with Google Search.
* Build a tool to define searched information, automatically updating prompt and response_schema.
* Adjust code in models/response_mail_model.py, services/excel_service.py and services/response_service.py to cooperate
with the tool above.
* Remove report columns and LLM parameters (ensuring secrecy of API key provision) from .env and integrate them with
the tool mentioned above.
* Build a CRON to automatically perform an analysis of new responses received every 24hrs, append Celery queueing to
the task.
* Some authorities refuse to respond to the request saying that the sender is not authorised by the organisation to
send requests in its name - consult the organisation in order how to solve the problem.

## Additional information
In case of any technical issues, feel free to contact me.
