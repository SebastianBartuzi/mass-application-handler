# mass-application-handler

This application allows Młoda Lewica members to automate the process of sending question requests to local governments in Poland. The program creates PDF files and sends e-mails with them attached, saving a hell of time.

You must ensure you have your @mlodalewica.pl e-mail address enabled in Microsoft Outlook in order to use this programme.

1. If you don't have Python installed in your system, make sure you download it from here: https://www.python.org/ftp/python/3.12.5/python-3.12.5-amd64.exe
2. IMPORTANT! During installation make sure you select the option to add Python to PATH environmental variables
3. Download this programme by going to the top of the page, select Code and Download ZIP. Extract the ZIP file.
4. Run python install.py
5. Fill local government data in debug/input/data.xlsx
6. Run python main.py
7. Files are ready in debug/output
8. To automatically send files, run python sendNoAuthorisation.py
9. To automatically send files and double check all messages, run python sendWithAuthorisation.py

In case of any technical issues, feel free to contact me.
