from outlook_to_jira import outlook_to_jira as o
from worklog_report import worklog_report as r


print('What do you want to do?')
print('Press "1" to get Outlook appointments and log them to JIRA tickets.')
print('Press "2" to get a worklog report for a specific user and timespan.')
sub_module = input()

if sub_module == '1':
    o.post_outlook_to_jira()
elif sub_module == '2':
    r.generate_report()
else:
    print(f'{sub_module} is not recognized.')