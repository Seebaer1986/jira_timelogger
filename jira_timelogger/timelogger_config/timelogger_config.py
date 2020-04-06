import os
from pathlib import Path
import configparser

config_filename = 'config.conf'
path_to_config = os.path.join(Path(os.path.dirname(__file__)).parent, config_filename)

def init_config():
    # check for existence of config file
    if not os.path.exists(path_to_config):
        print('Config.conf does not exist. Generating default config.')
        generate_default_config(path_to_config=path_to_config)

    # open config
    config = configparser.RawConfigParser(allow_no_value=True)
    config.read(path_to_config)

    return config

def generate_default_config():
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

def write_config(config, section, key, value):
    config.set(section, key, value)
    cfgfile = open(path_to_config, 'w')
    config.write(cfgfile)
    cfgfile.close()