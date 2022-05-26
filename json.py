import json
f = open('data.json')
data = json.load(f)
index =0
while(index<77):
    print(data[index]['blockDetails']['dateCreated'])
    index+=1
f.close()
