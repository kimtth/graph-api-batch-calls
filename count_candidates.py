import json

with open('api_result.json') as f:
    data = json.load(f)
    count = len(data['data']['meetingTimeSuggestions'])
    print(f'Total meeting time suggestions (candidates): {count}')
