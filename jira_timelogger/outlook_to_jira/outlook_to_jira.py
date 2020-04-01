from win32com.client import Dispatch
from jira import JIRA
from getpass import getpass
import datetime
import pdb
import sys
import configparser
import io
import re
import os
from pathlib import Path

def main():
    OUTLOOK_FORMAT = '%d.%m.%Y %H:%M'
    outlook = Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    config_filename = 'config.conf'
    path_to_config = os.path.join(Path(os.path.dirname(__file__)).parent, config_filename)

    # check for existence of config file
    if not os.path.exists(path_to_config):
        print('Config.conf does not exist. Generating default config.')
        generate_default_config(path_to_config=path_to_config)

    # open config
    config = configparser.RawConfigParser(allow_no_value=True)
    config.read(path_to_config)

    # check for a folderPath in config
    folder_path = config.get('Outlook', 'folder_path')
    if folder_path != '':
        print(f'Last time you used the following calendar: {folder_path}.')
        print('Press enter to use it again.')
        print('Type default to use the Outlook default calender')
        print('Enter a new if you want to use a new calendar.')
        user_input = input()

        if user_input != '':
            folder_path = user_input

    else:
        print('Press enter to use the Outlook default calender')
        print('Enter a new if you want to use a new calendar.')
        folder_path = input()
        if folder_path == '':
            folder_path = 'default'
             # clear folder_path from config
            write_config(config, path_to_config, section='Outlook', key='folder_path', value='')
        else:
            write_config(config, path_to_config, section='Outlook', key='folder_path', value=folder_path)

    # read category name to mark all processed items in outlook
    processed_category = config.get('Outlook', 'processed_category')
    if processed_category == '':
        processed_category_default = 'jira_logged'
        print('')
        print('Please input a name for the category which should be used to mark processed appointments.')
        user_input = input(f'Leave blank to use the programs default: "{processed_category_default}"')
        if user_input == '':
            processed_category = processed_category_default
        else:
            processed_category = user_input

    # write new category name to config
    write_config(config, path_to_config, section='Outlook', key='processed_category', value=processed_category)

    # check if the category is already present in outlook
    category_found = False
    for category in ns.Categories:
        if str(category) == processed_category:
            category_found = True

    if category_found == False:
        print('')
        try:
            ns.Categories.Add(processed_category)
            print(f'Info: Added category to Outlook: {processed_category}')
        except:
            print('Could not add category to Outlook.')
            sys.exit(1)

    #get all appointments from outlook
    print('')
    print('Please enter the date (format: YYYY-MM-DD) were you want to start processing items.')
    begin = input('Press ENTER to use default (today): ')    
    appointments = get_outlook_appointments(config, path_to_config=path_to_config, ns=ns, begin=begin)
    
    # ask for JIRA instance
    jira_url = config.get('Jira', 'url')
    if jira_url == '':
        while jira_url == '':
            print('')
            jira_url = input('Please specify the JIRA instance (URL): ')
    else:
        jira_input = input(f'Last time you connected to {jira_url}. Press ENTER to use it again or enter a new JIRA instance: ')
        if jira_input != '':
            jira_url = jira_input

    # write new JIRA URL to config
    write_config(config, path_to_config, section='Jira', key='url', value=jira_url)

    # ask for username
    jira_user = config.get('Jira', 'username')
    if jira_user == '':
        while jira_user == '':
            print('')
            jira_user = input('Please specify your username for JIRA: ')
    else:
        jira_input = input(f'Last time you connected to JIRA using {jira_user}. Press ENTER to use it again or enter a new JIRA username: ')
        if jira_input != '':
            jira_user = jira_input

    # write new JIRA username to config
    write_config(config, path_to_config, section='Jira', key='username', value=jira_user)

    # init some variables
    jira_password = ''
    use_token = ''

    # if JIRA cloud -> API Token needed
    if '.atlassian.net' in jira_url:
        print('')
        print('=====================================================')
        print('You are trying to connect to a JIRA Cloud instance.')
        print('Please make sure to generate an API Token at:')
        print('')
        print('https://id.atlassian.com/manage/api-tokens')
        print('')
        print('That API Token is then used instead of your Password.')
        print('=====================================================')
        print('')

        api_token = config.get('Jira', 'api_token')
        if api_token != '':
            use_token = input(f'You already saved an API Token. Do you want to use it again? y = yes, n = no: ')
            if use_token.upper() == 'Y':
                jira_password = api_token

    # ask for password/ token and try to authenticate
    if jira_password == '':
        jira_password = getpass()

    # ask if api_token should be stored
    if '.atlassian.net' in jira_url and use_token.upper() != 'Y':
        print('')
        save_token = input('Do you want to save the API Token for the next session? y = yes, n = no: ')
        
        if save_token.upper() == 'Y':
            write_config(config, path_to_config, section='Jira', key='api_token', value=jira_password)

    try:
        auth_jira = JIRA(jira_url, basic_auth=(jira_user, jira_password))
    except:
        print(f'Could not authenticate with the given credentials.')
        sys.exit(1)

    print('')
    print('Processing Outlook Appointments...')

    for appointmentItem in appointments:
        # check if it was already processed before
        if processed_category in appointmentItem.Categories:
            print(f'[Info] Appointment "{appointmentItem.Subject}" is already logged in Jira')
            continue
        
        # check if the subject contains valid JIRA Issue ID, if there are multiple matches, pick the first. REGEX: [A-Z0-9]*-[0-9]*
        m = re.search('[A-Z0-9]*-[0-9]*', appointmentItem.Subject.upper())
        try:
            ticket = m.group(0)
        except:
            # no valid Ticket ID String found
            print(f'[Info] Appointment "{appointmentItem.Subject}" does not contain a valid jira ticket id.')
            continue
        
        # check if there is a jira issue for the extracted ticket id
        try:
            issue = auth_jira.issue(ticket)
        except:
            print(f'[Info] Appointment "{appointmentItem.Subject}" could not be logged. There is no issue with the ID {ticket}')
            continue

        print(f'[Info] Processing outlook item "{appointmentItem.Subject}":')

        # log in jira jira.add_worklog("issue number", timeSpent="2h", comment="comment", started="")
        auth_jira.add_worklog(ticket,timeSpent=appointmentItem.Duration,comment=appointmentItem.Subject, started=appointmentItem.Start)
        print(f'    Logged {appointmentItem.Duration} minutes on {ticket}.')

        # add processed category to outlook item and save it in outlook
        appointmentItem.Categories = appointmentItem.Categories + ', ' + processed_category
        appointmentItem.Save()
        print(f'    Added the category {processed_category} to the Outlook item.')

def write_config(config, path_to_config, section, key, value):
    config.set(section, key, value)
    cfgfile = open(path_to_config, 'w')
    config.write(cfgfile)
    cfgfile.close()

def generate_default_config(path_to_config):
    # create config object
    config = configparser.ConfigParser()
    config['Outlook'] = {'folder_path':'',
                        'processed_category':''}
    config['Jira'] = {'url':'',
                    'username':'',
                    'api_token':''}
    # write config file
    cfgfile = open(path_to_config, 'w')
    config.write(cfgfile)
    cfgfile.close()

def get_outlook_appointments(config, path_to_config, ns, begin, end=''):
    # get folder_path from config
    folder_path = config.get('Outlook', 'folder_path')

    # if you should use the default calendar
    if folder_path == '':
        appointments = ns.GetDefaultFolder(9).Items
        
    else:
    # if folderPath was given, check if it is a legit path
        # cut away the first \\ if there are any
        if folder_path[:2] == '\\\\':
            folder_path = folder_path[2:]

        # split folder path
        folders = folder_path.split('\\')

        # try to set folder to root of the path
        try:
            folder = ns.Folders.Item(folders[0])

        except:
            print(f'Given path to calendar "{folder_path}" is not correct, please check.')
            sys.exit(1)

        # for the rest of the path check also if the path is legit
        for sub_folder in folders[1:]:
            try:
                folder = folder.Folders.Item(sub_folder)
            except:
                print(f'Given path to calendar "{folder_path}" is not correct, please check.')
                sys.exit(1)

        appointments = folder.Items

    if begin == '':
        begin = datetime.date.today().strftime('%d/%m/%Y')
    else:
        begin = begin[8:10]+'/'+begin[5:7]+'/'+begin[:4]

    if end == '':
        end = datetime.date.today().strftime('%d/%m/%Y')
    else:
        end = end[8:10]+'/'+end[5:7]+'/'+end[:4]

    # for debugging
    #print(f'Begin: {begin} | End: {end}')

    # restrict outlook items to the items in this timeframe
    restriction = f'[Start] >= "{begin} 0:00am" AND [Start] <= "{end} 11:59pm"'
    restricted_items = appointments.Restrict(restriction)
    restricted_items.IncludeRecurrences = 'True'
    restricted_items.Sort('[Start]')

    ## for debugging -> output restricted items and end the program
    #for item in restricted_items:
    #    print(f'{item.Subject} ({item.Start})')
    #sys.exit(1)

    return restricted_items

if __name__ == "__main__":
    main()