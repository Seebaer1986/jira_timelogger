from setuptools import setup, find_packages

setup(name='jira_timelogger',
      version='0.2',
      description='Module to get appointment information from outlook and create worklog entries in a jira instance. Also it can create a worklog report based on logs in JIRA.',
      url='',
      author='MaMS',
      author_email='m.mueller-splett@fmmail.de',
      license='MIT',
      packages=find_packages(),
      install_requires=[
          'jira',
          'pywin32'
      ],
      zip_safe=False
      )