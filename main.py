import requests
import pandas as pd

df = pd.DataFrame({'ReviewId': pd.Series([], dtype='string'),
                             'ProductId': pd.Series([], dtype='string'),
                             'SubmissionTime': pd.Series([], dtype='string'),
                             'Rating': pd.Series([], dtype='string'),
                             'ReviewText': pd.Series([], dtype='string'), 
                             'UserNickname': pd.Series([], dtype='string'),
                             'BadgesOrder': pd.Series([], dtype='string'),
                             'skinTone': pd.Series([], dtype='string'),
                             'skinType': pd.Series([], dtype='string'),
                             'eyeColor': pd.Series([], dtype='string'),
                             'hairColor': pd.Series([], dtype='string'), 
                             })


for x in range(32):

    headers = {
        'Accept': '*/*',
        'Accept-Language': 'ja,en;q=0.9',
        'Connection': 'keep-alive',
        'Origin': 'https://www.sephora.com',
        'Referer': 'https://www.sephora.com/',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'cross-site',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'sec-ch-ua': '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"macOS"',
    }

    response = requests.get(
        "https://api.bazaarvoice.com/data/reviews.json?Filter=contentlocale%3Aen*&Filter=ProductId%3AP510914&Sort=SubmissionTime%3Adesc&Limit=6&Offset="+ str(x*6) +"&Include=Products%2CComments&Stats=Reviews&passkey=calXm2DyQVjcCy9agq85vmTJv5ELuuBCF2sdg4BnJzJus&apiversion=5.4&Locale=en_US",
        headers=headers,
    )

    # print(response.content)

    import json
    import csv

    # Parse JSON response
    response_data = json.loads(response.content)
    reviews = response_data['Results']

    # Define CSV file path
    csv_file = 'reviews.csv'

    # Define CSV header
    header = ['ReviewId', 'ProductId', 'SubmissionTime', 'Rating', 'ReviewText', 'UserNickname', 'BadgesOrder', 'skinTone', 'skinType', 'eyeColor', 'hairColor']

    for review in reviews:
        collected_data = pd.DataFrame({
            'ReviewId': review['Id'],
            'ProductId': review['ProductId'],
            'SubmissionTime': review['SubmissionTime'],
            'Rating': review['Rating'],
            'ReviewText': review['ReviewText'],
            'UserNickname': review['UserNickname'],
            'BadgesOrder': ', '.join(review.get('BadgesOrder', [])),
            'skinTone': review.get('ContextDataValues', {}).get('skinTone',{}).get('Value', ''),
            'skinType': review.get('ContextDataValues', {}).get('skinType',{}).get('Value', '') ,
            'eyeColor': review.get('ContextDataValues', {}).get('eyeColor',{}).get('Value', ''),
            'hairColor': review.get('ContextDataValues', {}).get('hairColor',{}).get('Value', '') 
        },index =[0])
        df = pd.concat([collected_data, df]).reset_index(drop = True)
            
    print("fetched page: "+str(x))
    # print(f'Reviews saved to {csv_file}')


df.to_excel("ExtractedReviews.xlsx")
