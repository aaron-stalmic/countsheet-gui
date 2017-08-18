import sys
from os import path, getcwd
import re


def get_config():
    if getattr(sys, 'frozen', False):
        application_path = path.dirname(sys.executable)
    elif __file__:
        application_path = path.dirname(__file__)

    config_path = path.join(application_path, 'config.cfg')

    try:
        with open(config_path, 'r') as file:
            contents = file.read()
            server = re.search('server = (.*)', contents).group(1)
            database = re.search('database = (.*)', contents).group(1)
            username = re.search('username = (.*)', contents).group(1)
            password = re.search('password = (.*)', contents).group(1)
    except FileNotFoundError:
        error = "There was no configuration file found in the root folder."
        error += " Please make sure a config file is included that includes"
        error += " the server, database, username, and password."
        tk.messagebox.showerror("Pick Ups To Notes", error)
    except AttributeError:
        error = "There was a problem with the configuration file."
        error += " Please make sure it includes a server, database, username, and password."
        tk.messagebox.showerror("Pick Ups To Notes", error)
    return [server, database, username, password]

config = get_config()
