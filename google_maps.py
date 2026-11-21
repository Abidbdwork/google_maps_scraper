import json
import logging
import os
import random
from datetime import datetime

import requests
import pandas as pd

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')


class google_maps:
    def __init__(self):
        root = os.path.dirname(__file__)
        self.output = os.path.join(root, 'output')
        os.makedirs(self.output, exist_ok=True)
        env = dict(os.environ)
        self.map_api_key = env.get("map_api_key")  # AIzaSyA6my HzS10YXdcaz AFalmXvDkr YCp5cLc8
        self.user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/157.0.0.0 Safari/537.36"
        self.headers = {
            'accept': '*/*',
            'accept-language': 'en-US,en;q=0.9,bn;q=0.8,de;q=0.7',
            'content-type': 'application/json+protobuf',
            'origin': 'https://maps-docs-team.web.app',
            'priority': 'u=1, i',
            'referer': 'https://maps-docs-team.web.app/',
            'sec-ch-ua': '"Google Chrome";v="155", "Chromium";v="155", "Not(A:Brand";v="24"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'cross-site',
            'user-agent': self.user_agent,
            'x-browser-channel': 'stable',
            'x-goog-api-key': self.map_api_key,
            'x-user-agent': 'grpc-web-javascript/0.1'
        }
        self.column_mapping = {"company_name": "Business name", "phone": "Business phone number", "full_address": "Complete address",
                               "website": "Website", "city": "City", "state": "State", "map_url": "Source URL"}
        self.all_working_proxies = [i.strip() for i in open("proxies.txt", encoding="utf-8").read().splitlines() if i.strip() and i.strip().startswith("http")]
        # temp input

        if env.get('JENKINS_URL'):
            queries = env.get("Search_queries")
            if queries:
                queries = [i.strip() for i in queries.splitlines() if i.strip()]
                logging.info(f"Total input queries found: {len(queries)}")
            else:
                raise Exception("No input queries found")
        else:
            # temp use
            # queries = ["cafe in barishal"]
            queries = [input("Please enter your search query: ")]
        self.queries = queries
        self.concurrency = 1
        self.main_records = []
        self.start_process()

    def start_process(self):
        for ind, query in enumerate(self.queries, start=1):
            logging.info(f"Currently running {ind}/{len(self.queries)} || Query: {query}")
            main_records = self.search_places(query)
            self.main_records.extend(main_records)

        main_df = pd.DataFrame(self.main_records)
        main_df.rename(self.column_mapping, axis=1, inplace=True)
        final_df = main_df[list(self.column_mapping.values())]
        final_df['Collection date'] = datetime.now().date()
        final_df.to_excel(os.path.join(self.output, 'google_maps.xlsx'), index=False)

    def search_places(self, Query: str, Region_filter=None, Page_size=20, Retry=10):
        url = "https://places.googleapis.com/$rpc/google.maps.places.v1.Places/SearchText?alt=json"
        fields = 'places.name,places.id,places.nationalPhoneNumber,places.internationalPhoneNumber,places.formattedAddress,places.addressComponents,places.plusCode,places.location,places.viewport,places.rating,places.websiteUri,places.googleMapsUri,places.adrFormatAddress,places.displayName,nextPageToken'
        next_page_token = None
        payload = [Query, "en-US", Region_filter, None, None, None, None, None, None, Page_size,
                   None, False, None, None, None, None, None, None, next_page_token]
        headers = self.headers.copy()
        headers['x-goog-fieldmask'] = fields
        main_records = []
        while True:
            for i in range(Retry):
                if self.all_working_proxies:
                    proxy = random.choice(self.all_working_proxies)
                    proxy_dict = {"https": proxy, "http": proxy}
                else:
                    proxy_dict = None
                try:
                    response = requests.post(url, headers=headers, data=json.dumps(payload), proxies=proxy_dict, timeout=10)
                    if response.status_code == 200:
                        data = response.json()
                        next_page_token = data.get("nextPageToken")
                        main_records += self.place_parser(data.get('places', []))
                        break
                    else:
                        logging.warning(f"Retrying - {i + 1}. Unable to get correct response -- {response.status_code} {response.text}")
                except Exception as e:
                    logging.error(f"Retrying - {i + 1}. Error during sending requests. - {proxy_dict} - Error:", exc_info=True)
            if next_page_token:
                payload[-1] = next_page_token
                logging.info(f"Next page token: {next_page_token}")
            else:
                logging.info("No more pages")
                break
        return main_records

    @staticmethod
    def place_parser(Data_records):
        address_mapping = {"street_number": "street_number", "route": "street_address", "locality": "city", "sublocality": "town",
                           "administrative_area_level_1": "state", "postal_code": "zip_code", "country": "country"}
        new_records = []
        for record in Data_records:
            try:
                address_components = record.get('addressComponents', [])
                phone = record.get('internationalPhoneNumber')
                website = record.get('websiteUri')
                display_name = record.get('displayName', {}).get('text')
                full_address = record.get('formattedAddress')
                map_url = record.get('googleMapsUri')
                temp = {'company_name': display_name, "phone": phone, "website": website, "full_address": full_address,
                        "map_url": map_url}
                location = record.get('location')
                address_dict = {}
                for component in address_components:
                    types = component.get('types', [])
                    long_value = component['longText']
                    short_value = component.get('shortText')

                    for address_type in types:
                        if address_type in address_mapping:
                            value = long_value
                            if address_type in ["country"]:
                                value = short_value
                            address_dict[address_mapping[address_type]] = value
                            break

                address = f"{address_dict.get('street_address', '')} {address_dict.get('street_number', '')}".strip()
                new_dict = {'address': address, **{i: address_dict[i] for i in address_dict if
                                                   i in ['city', 'zip_code', 'state', 'country']}}
                if not new_dict.get('city') and address_dict.get('town'):
                    new_dict['city'] = address_dict['town']
                new_dict.update(temp)
                new_dict.update(location)
                new_records.append(new_dict)
            except Exception:
                logging.error(f"Error during parsing -- {record} - Error:", exc_info=True)
        return new_records


if __name__ == '__main__':
    google_maps()
