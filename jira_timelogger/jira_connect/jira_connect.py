from jira import JIRA
from jira_timelogger.timelogger_config import timelogger_config as tc
from getpass import getpass
import sys

def connect(config):
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
    tc.write_config(config, section='Jira', key='url', value=jira_url)

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
    tc.write_config(config, section='Jira', key='username', value=jira_user)

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
            tc.write_config(config, section='Jira', key='api_token', value=jira_password)

    try:
        auth_jira = JIRA(jira_url, basic_auth=(jira_user, jira_password))

        return auth_jira
    except:
        print(f'Could not authenticate with the given credentials.')
        sys.exit(1)