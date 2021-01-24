
import json

CONFIG = {}

with open('./config/config.json') as json_file:
    CONFIG = json.load(json_file)