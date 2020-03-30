from setuptools import setup

setup(name='jira_timelogger',
      version='0.1',
      description='Module to get appointment information from outlook and create worklog entries in a jira instance.',
      url='',
      author='MaMS',
      author_email='m.mueller-splett@fmmail.de',
      license='MIT',
      packages=['jira_timelogger'],
      install_requires=[
          'jira',
          'pywin32'
      ],
      zip_safe=False)