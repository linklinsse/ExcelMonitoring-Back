
from flask import Flask
from flask_cors import CORS, cross_origin


from src import config

app = Flask(__name__)
if config.CONFIG['CORS']:
    cors = CORS(app)
    app.config['CORS_HEADERS'] = 'Content-Type'

from src import excel

if __name__ == '__main__':
    app.run(debug=config.CONFIG['debug'], host=config.CONFIG['host'], port=config.CONFIG['port'])
