
#THE BEST RESULTS WITH INCREASING VALUE OF POOl
import csv
import time
from multiprocessing import Pool
import requests

start=time.time()
def hit_api(object):
    address, city, pincode, referenceId = [object[k] for k in ('streetAddress', 'city', 'pinCode', 'referenceId')]
    row = {
        "streetAddress": address,
        "city": city,
        "pinCode": pincode,
        "referenceId": referenceId
    }
    response = requests.post("API_name", json=row,
                             headers={"x-api-key": "API_KEY"})
    json_response = response.json()
    return json_response


def read_from_csv(path):
    csv.register_dialect('myDialect',
                         delimiter=',',
                         quoting=csv.QUOTE_ALL,
                         skipinitialspace=True)
    with open(path, 'r') as f:
        reader = csv.DictReader(f, dialect='myDialect')
        p = Pool(100)
        data = p.map(hit_api, reader)
        p.close()
    print(data)
    return data

def write_to_csv(path,response_array):
    with open(path, 'w') as csvFile:
        fields = ['total', 'Id', 'lat', 'no', 'pop', 'is', 'Area', 'lon',
                  'as', 'distance', 'total_']   #Required Headers in CSV

        writer = csv.DictWriter(csvFile, fieldnames=fields,extrasaction='ignore', delimiter = ',')
        writer.writeheader()
        writer.writerows(response_array)

if __name__ =='__main__':
    response = read_from_csv('file to be read')
    write_to_csv('file to write',response)
    end=time.time()
print(end-start)