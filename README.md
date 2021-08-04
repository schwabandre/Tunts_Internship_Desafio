# Tunts_Internship_Desafio - Andre Schwab

This program called "Grade-Calculator", reads input from spreadsheet from Google Sheets and calculates the students grades,
and it updates the spreadsheet evaluating if the student was either apprvoed or not. It was implemented in python language using Google Sheet API v4. 

### Credentials

You must configure a Google OAuth 2.0 client ID in order to the program work correctly.

Access [Google API Console](https://console.developers.google.com), select or create a project and **select Create credentials**, then **OAuth client ID**.
Then, copy the `credentials.json.example` file to a new file named `credentials.json` and add the following information from the OAuth client ID you just created:

```
{ 
    "client_id": "",
    "project_id": "",
    "client_secret": "",
}
```

### How to Execute 

It's required to install google package with the following command:
`pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib`

Than execute the program with the command:
`python calculo_notas.py`

The program will attempt to open a new window or tab in your default browser. If this fails, copy the URL from the console and manually open it in your browser.

If you are not already logged into your Google account, you will be prompted to log in. If you are logged into multiple Google accounts, you will be asked to select one account to use for the authorization.

Click the Accept button.
The program will proceed automatically, and you may close the window/tab.