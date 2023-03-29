import requests
import json
import pprint

url= 'https://api.cafci.org.ar/fondo/285/clase/1711/ficha'

response=requests.get(url).content

json=json.loads(response)

pprint.pprint(json)

print(json)

#for i in range(50000):
#    for j in range(50000):
#        try:
#            url= f'https://api.cafci.org.ar/fondo/{i}/clase/{j}/ficha'
#            response = requests.get(url)
#            if response.status_code == 200 and json.loads(response.content)!="":
#                json = json.loads(response.content)
#                print(json["data"]["model"]["nombre"],i,j)
#        except:
#            continue

1613242