import requests
import json
import xlsxwriter
import time

# Enter your API key here
API_KEY = 'wwsdftt7pw6a2scvd8hd3qcf'

# Get the desired number of records from the user
num_records = int(input("Enter the number of records you want to extract: "))

# Set the base URL for the IEEE Explore API
base_url = 'http://ieeexploreapi.ieee.org/api/v1/search/articles?apikey={}&format=json&max_records={}&start_record=1&sort_order=asc&sort_field=article_title&querytext=blockchain'.format(
    API_KEY, num_records)

# Send a GET request to the API and retrieve the response
response = requests.get(base_url)

# Parse the response as JSON
data = json.loads(response.text)

# Create a new Excel workbook and worksheet
workbook = xlsxwriter.Workbook('blockchain_publications.xlsx')
worksheet = workbook.add_worksheet()

# Write the headers to the worksheet
worksheet.write(0, 0, 'Title')
worksheet.write(0, 1, 'Abstract')
worksheet.write(0, 2, 'Authors')
worksheet.write(0, 3, 'Year')
worksheet.write(0, 4, 'Publication type')

# Loop through the articles in the response and write the publication information to the worksheet
for i, article in enumerate(data['articles']):
    title = article['title']
    abstract = article['abstract']

    # Check the type of the authors field
    if isinstance(article['authors'], list):
        # If the authors field is a list of dictionaries, extract the full name of each author
        authors = [author['full_name'] for author in article['authors']]
    else:
        # If the authors field is not a list of dictionaries, set the authors field to an empty list
        authors = []

    year = article['publication_year']
    pub_type = article['content_type']

    # Write the publication information to the worksheet
    worksheet.write(i + 1, 0, title)
    worksheet.write(i + 1, 1, abstract)
    worksheet.write(i + 1, 2, ', '.join(authors))
    worksheet.write(i + 1, 3, year)
    worksheet.write(i + 1, 4, pub_type)

    # Print status indicator after every 100 records extracted
    if (i+1) % 100 == 0:
        print(f"{i+1} records extracted.")

    # Add a 1-second delay after every 8 records
    if (i+1) % 8 == 0:
        time.sleep(1)

# Close the workbook
workbook.close()

# Print a message to confirm the data has been saved to the Excel sheet
print(f'{num_records} publication data saved to blockchain_publications.xlsx')
