import json

def load_data():
  with open('data.json') as f:
    return json.load(f)

def save_data(data):
  with open('data.json', 'w') as f:
     json.dump(data, f)
