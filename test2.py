import requests
from bs4 import BeautifulSoup


import openpyxl

api_key = "136badc2a5cb66709725678516c4de04bd7cde806f97ba2c893de7a6a7e3a209"  # Fill in with your API key
# temp = ""
# output = ""
currentdata = [
    {
        "JTA_ID": "JTA152131",
        "COMPANY": "Toll Brothers Commercial",
        "NAME": "Haverly",
        "ADDRESS": "31 East Thomas Road",
        "CITY": "Phoenix",
        "ZIP": 85012,
    },
    {
        "JTA_ID": "JTA153540",
        "COMPANY": "Pangea Real Estate",
        "NAME": "8101 S Justine",
        "ADDRESS": "8101 S Justine",
        "CITY": "Chicago",
        "ZIP": 60620,
    },
    {
        "JTA_ID": "JTA155522",
        "NAME": "Catalyst Midtown",
        "ADDRESS": "1011 Northside Dr NW",
        "CITY": "Atlanta",
        "ZIP": 30318,
    },
]


def serp_search(query):
    url = "https://serpapi.com/search"
    params = {
        "api_key": api_key,
        "q": query
    }
    response = requests.get(url, params=params)
    data = response.json()["organic_results"]
    from_apartment = None
    for item in data:
        print("data")  
        if item["source"] == "Apartments.com":
            from_apartment = item
            break
    
    if not from_apartment:
        result = {
            "rating": "No",
            "review": 0
        }
        return result
    
    rating = from_apartment.get("rich_snippet", {}).get("top", {}).get("detected_extensions", {}).get("rating")
    votes = from_apartment.get("rich_snippet", {}).get("top", {}).get("detected_extensions", {}).get("votes")
    apartmentURL = from_apartment["link"]
    
    if rating and votes:
        output = {
            "rating": rating,
            "review": votes
        }
        return output
    else:
        return search_apartment(apartmentURL)

def search_apartment(url):
    try:
        response = requests.get(f"{url}/#reviewsSection")
        soup = BeautifulSoup(response.text, "html.parser")
        ratingText = soup.select_one(".averageRating").get_text()
        reviewText = soup.select(".ratingReviewsWrapper p:last-of-type")[0].get_text()
        review_result = reviewText.split(" ")[0]
        review = int(review_result) if review_result != "No" else 0
        ratings = int(ratingText) if ratingText else 0
        
        output = {
            "rating": ratings if ratingText else 0,
            "review": review if reviewText else 0
        }
        
        return output
    except Exception as error:
        raise error

def print_data(data):
    arr = []
    for obj in data[2:]:
        try:
            Q = f"{obj['NAME']}, {obj['ADDRESS']}, {obj['CITY']}, {obj['ZIP']}"
            print(Q)
            output = serp_search(Q)
            print(output)
            ratings = output["rating"]
            reviews = output["review"]
            table = {
                "JTA_ID": obj["JTA_ID"],
                "NAME": obj["NAME"],
                "ADDRESS": obj["ADDRESS"],
                "CITY": obj["CITY"],
                "ZIP": obj["ZIP"],
                "QUERY": Q,
                "APARTMENT_RATINGS": ratings if ratings != "No" else None,
                "APARTMENT_REVIEWS": reviews if reviews else 0
            }
            arr.append(table)
        except Exception as error:
            print(error)
    
    print(len(arr))
    print(arr)
    
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    
    for i, row in enumerate(arr, start=1):
        for j, value in enumerate(row.values(), start=1):
            worksheet.cell(row=i, column=j, value=value)
    
    workbook.save("output3.xlsm")
    print("XLSM file created successfully.")

print_data(currentdata)
