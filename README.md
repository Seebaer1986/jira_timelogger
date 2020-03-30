# Jira Timelogger
Read appointments from Microsoft Outlook and use them to log work in an Atlassian JIRA instance.

## Dependencies
* jira-python
* pywin32

## Installation
* download the repo
* navigate to the repo on your disk with a command prompt
* install dependencies: `pip install .`
* run the script: `python jira_timelogger\__init__.py`

## Usage
The script will guide you through and ask different questions:

### Steps of the script
#### 1. Ask for the path to your Outlook calendar
You can just hit ENTER to use the default calendar. If you set up a custom calendar in Outlook to track your work, then you need to specify the path:
`\\Outlook-file\Calendar\MyLogCalendar`

To find the path, go to Outlook and right click your calendar. Choose `Properties...` and then the patch is your `Location:` followed by the name of the calendar itself.

#### 2. Ask for a category to use to mark logged appointments
Under the hood the script will assign an additional category to each appointment which was successfully logged in Jira.
You can specify your own category name or go with the default `jira_logged`. 
In both cases the script will create the category in your Outlook if neccessary.

#### 3. Ask for a date to start looking for appointments to log
The script will get all appointments from the specified calendar starting between the date specified here (00:00am) and today 23:59pm. If you just hit ENTER it will only get appointments for the current day.
Please enter the date in the `YYYY-MM-DD` format.

#### 4. Ask for the URL of your JIRA instance
You can specify Cloud and Server instances. Please provide the full adress including `https://`.

#### 5. Ask for the username to use to authenticate in JIRA
Your Atlassian JIRA `username`.

#### 6. Ask for your password or API Token
If you use a _server instance_ the script will ask for your _password_, which you can enter securely (is not shown).
If you want to connect to a _Cloud instance_ of JIRA then you will need to specify a generated _API token_, since basic ath with a password is no longer possible.
To obtain a token go to https://id.atlassian.com/manage/api-tokens.

### Saving your settings
All of your answers (except your password in a server instance) will be saved in `config.conf`, saving you time when you run the script again.

### How are the issues identified?
The script parses your appointments subjects and searches for a pattern matching JIRA issue ID's: `[A-Z0-9]*-[0-9]*`
If your subject contains more than one match, it will use the first match.
If your subject does not contain a match the appointment will be ignored.
If your subject does contain a match but it is not a valid issue in your JIRA instance, the appointment will be ignored.
