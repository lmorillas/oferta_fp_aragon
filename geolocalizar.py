#coding: utf-8

from unicodecsv import DictReader
from geopy.geocoders import GoogleV3
import json

dr = DictReader(open('centros.csv'))
centros = [d for d in dr]

geocoder = GoogleV3()

for c in centros:
    #if not c.get('latlon'):
        direccion = u'{}, {}, Aragón, Spain'.format(c.get(U'DIRECCIÓN'), c.get('LOCALIDAD'))
        x = geocoder.geocode(direccion)
        if x:
            c['latlon'] = '{},{}'.format(x.latitude, x.longitude)


print 'No encontrados', [c for c in centros if not c.get('latlon')]

json.dump(centros, open('centros2.json', 'w'))
