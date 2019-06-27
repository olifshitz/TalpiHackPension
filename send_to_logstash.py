from elasticsearch import *
import json
import pickle

ES_HOST_IP = 'localhost'
ES_HOST_PORT = 9200

es = Elasticsearch([{'host': ES_HOST_IP, 'port': ES_HOST_PORT}])

content = pickle.load(open('data_json.json', 'rb'))

for i,data in enumerate(content):
    es.index(index='pensions', id=i+1, body=json.loads(data))
