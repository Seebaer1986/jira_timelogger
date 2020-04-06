from win32com.client import Dispatch
from jira import JIRA
from getpass import getpass
import datetime
import sys
from jira_timelogger.timelogger_config import timelogger_config as tc

def generate_report():
    config = tc.init_config()
    print(f'{config.get("Outlook", "folder_path")}')

    tc.write_config(config, section='Outlook', key='folder_path', value='Outlook-Datendatei\Calendar\Zeiterfassung')

if __name__ == "__main__":
    generate_report()