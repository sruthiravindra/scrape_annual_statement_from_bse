import requests
from bs4 import BeautifulSoup
import pandas as pd 

# url of the page to scrape
# scrape_url = 'https://www.bseindia.com/corporates/results.aspx?Code=500875&Company=ITC%20LTD.&qtr=117.50&RType=D&Typ=A&Qname=2023'
scrape_url = 'https://www.bseindia.com/corporates/results.aspx?Code=500875&Company=ITC%20LTD.&qtr=117.50&RType=c&Typ=A&Qname=2023'
target_excel_file_name = 'bse_itc_consolidated.xlsx'

try:
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
    response = requests.get(scrape_url, headers=headers)

    # raise an HTTPError if the HTTP request returned an unsuccessful status code
    response.raise_for_status()

    beaut_sp = BeautifulSoup(response.text, 'html.parser')

    # the table that we need to parse is ContentPlaceHolder1_tbl_typeID
    table_id = 'ContentPlaceHolder1_tbl_typeID'
    table_html_content = beaut_sp.find('table', {'id': table_id})

    # case: table exists
    if table_html_content:

        # get all rows in the table
        table_html_rows = table_html_content.find_all('tr')
        data = [] # data is a multi dimensional datastructure and will contain each rows column content

        # loop all rows
        for row in table_html_rows:

            # get list of all all text from each column/cell
            row_content = [cell.text.strip() for cell in row.find_all(['td', 'th'])]
            data.append(row_content)

        df = pd.DataFrame(data[1:], columns=data[0])

        df.to_excel(target_excel_file_name, index=False)
        print('success')
    else:
        print('empty html content')

except requests.exceptions.HTTPError as errh:
    print(f"HTTP Error: {errh}")
except requests.exceptions.ConnectionError as errc:
    print(f"Error Connecting: {errc}")
except requests.exceptions.Timeout as errt:
    print(f"Timeout Error: {errt}")
except requests.exceptions.RequestException as err:
    print(f"Request Error: {err}")