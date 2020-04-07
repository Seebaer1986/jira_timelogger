import datetime
import calendar
import sys
from jira_timelogger.timelogger_config import timelogger_config as tc
from jira_timelogger.jira_connect import jira_connect as jc

def generate_report():
    # get config file
    config = tc.init_config()

    # connect to jira instance
    jira = jc.connect(config)

    # ask for username or user from config
    username_config = config.get('Jira', 'username')
    print('')
    print(f'Please press ENTER to use "{username_config}"" as the user for the worklog report.')
    username = input(f'If you want to generate the report for another user, please specify the users mail address: ')
    if username == '':
        username = username_config 

    # ask for the start date
    begin_default = datetime.date.today().replace(day=1).strftime('%Y-%m-%d')
    print('')
    print(f'Press ENTER to generate the worklog report starting at "{begin_default}"."')
    begin = input('If you want to generate the report starting at another date, please enter the date (format: "YYYY-MM-DD"): ')
    if begin == '':
        begin = begin_default

    # ask for the end date
    end_default = datetime.date.today().replace(day=calendar.monthrange(datetime.date.today().year,datetime.date.today().month)[1]).strftime('%Y-%m-%d')
    print('')
    print(f'Press ENTER to generate the worklog report until "{end_default}"."')
    end = input('If you want to generate the report ending at another date, please enter the date (format: "YYYY-MM-DD"): ')
    if end == '':
        end = end_default

    # get all issues were there is a worklog by the questioned user
    issues = jira.search_issues(f'worklogAuthor = "{username}" and worklogDate >= "{begin}" AND worklogDate <= "{end}"')

    # iterate through the issues and their worklogs and them to a dictionary if it matches the search criteria
    report_logs = []
    print('')

    # print header of CSV
    print('Project|Ticket|date|time worked')

    for issue in issues:
        # print(f'Processing worklog of {issue}')
        worklogs = jira.worklogs(issue)
        for worklog in worklogs:
            if worklog.started >= begin and worklog.started <= end and worklog.author.emailAddress == username:
                print(f'{issue.fields.project}|{issue.key}|{worklog.started[:10]}|{worklog.timeSpent}')
                # print('adding worklog to report')
                # report_row = { 
                #     'issue': issue.key, 
                #     'date': worklog.started,
                #     'time_spent': worklog.timeSpent
                # }
                # report_logs += report_row
            # else:
            #     print('worklog ignored')

    # print(report_logs)  


if __name__ == "__main__":
    generate_report()