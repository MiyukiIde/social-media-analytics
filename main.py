import requests
import pandas as pd

# Define the pages list before the loop
pages = ['3AP510914', '3AP510911', '3AP511095', '3AP454389', '3AP506377', '3AP506192', '3AP511097', '3AP395723', '3AP470058', '3AP504510', '3AP454387', '3AP505211', '3AP510384', '3AP482257', '3AP456407', '3AP511275', '3AP504986', '3AP500399', '3AP504911', '3AP482747', '3AP504987', '3AP500287', '3AP504513', '3AP505483', '3AP505208', '3AP510551', '3AP510386', '3AP504519', '3AP505875', '3AP405944', '3AP419222', '3AP419221']

df = pd.DataFrame({
    'ReviewId': pd.Series([], dtype='string'),
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

for page in pages:
    offset = 0  # Initialize offset to 0 for each page
    while True:
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
            f"https://api.bazaarvoice.com/data/reviews.json?Filter=contentlocale%3Aen*&Filter=ProductId%{page}&Sort=SubmissionTime%3Adesc&Limit=6&Offset={offset}&Include=Products%2CComments&Stats=Reviews&passkey=calXm2DyQVjcCy9agq85vmTJv5ELuuBCF2sdg4BnJzJus&apiversion=5.4&Locale=en_US",
            headers=headers,
        )

        # Parse JSON response
        response_data = response.json()
        reviews = response_data['Results']

        if not reviews:  # If no more reviews are returned, break the loop
            break

        for review in reviews:
            collected_data = pd.DataFrame({
                'ReviewId': [review['Id']],
                'ProductId': [review['ProductId']],
                'SubmissionTime': [review['SubmissionTime']],
                'Rating': [review['Rating']],
                'ReviewText': [review['ReviewText']],
                'UserNickname': [review['UserNickname']],
                'BadgesOrder': [', '.join(review.get('BadgesOrder', []))],
                'skinTone': [review.get('ContextDataValues', {}).get('skinTone', {}).get('Value', '')],
                'skinType': [review.get('ContextDataValues', {}).get('skinType', {}).get('Value', '')],
                'eyeColor': [review.get('ContextDataValues', {}).get('eyeColor', {}).get('Value', '')],
                'hairColor': [review.get('ContextDataValues', {}).get('hairColor', {}).get('Value', '')],
            })
            df = pd.concat([collected_data, df], ignore_index=True)

        print(f"Fetched {len(reviews)} reviews for page {page} at offset {offset}")
        offset += 6  # Increment offset for the next batch of reviews

df.to_excel("ExtractedReviews.xlsx")

