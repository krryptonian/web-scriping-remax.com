from time import sleep
import pandas as pd
from bs4 import BeautifulSoup
import requests


def make_hyperlink(link, text):
    return '=HYPERLINK("%s", "%s")' % (link, text)


def getDetails(profile_link):
    response = requests.get(profile_link, timeout=5)
    soup = BeautifulSoup(response.content, 'html.parser')
    container = soup.find(
        attrs={'data-test': 'agent-bio'}).find(attrs={'class': 'noprint'})
    # get agent name
    agent_name = container.find(
        'h1', attrs={'class': 'h2 mt-6'}).get_text(strip=True)
    # get profile picture

    # profileContainer = container.find(attrs={'class': 'col md:max-w-1/2 lg:max-w-full'})
    # profile_picture = profileContainer.find('img')
    # print(profile_picture['data-src'])

    # old profile picture code

    # profile_picture = container.find('img',attrs={'class':'mb-4 photo-max-height'}).get('src') if container.find('img',attrs={'class':'mb-4 photo-max-height'}) else ""
    # get phone numbers
    phones_container = container.find(
        'div', attrs={'class': 'bio-phone mb-8 lg:mb-0'})
    agent_phones = ';'.join([phone.get_text(
        strip=True) for phone in phones_container.find_all('h4')] if phones_container else [])
    # Get agent website
    website_link = container.find('a', class_='my-website').get(
        'href', '') if container.find('a', class_='my-website') else ""
    # get languages
    agent_language = container.find('p', attrs={'data-test': 'bio-languages'}).find('span').get_text(
        strip=True) if container.find('p', attrs={'data-test': 'bio-languages'}) else ""

    # Find address container
    address_container = container.find(
        'a', class_='inline-block directions-link')
    if not address_container:
        print("Address not found")
        return
    # Get address
    address = address_container.get_text(strip=True)
    # Get location link
    location_link = address_container.get('href', '')

    # Find social network container
    social_links = ''
    social_container = container.find('div', class_='mb-12').find_all(
        'a', class_='social-icon') if container.find('div', class_='mb-12') else []
    for social in social_container:
        social_links += f"{social.get('href')}\n"

    return [agent_name, agent_phones, website_link, social_links, agent_language, make_hyperlink(location_link, address)]


# url = 'https://www.remax.com/real-estate-agents/-ca?filters={"page":1,"count":"96","sortBy":"lastName"}'

def getPageData(url):
    response = requests.get(url, timeout=5)
    soup = BeautifulSoup(response.content, 'html.parser')
    container = soup.find(attrs={'class': 'roster-container'})
    data = []
    i = 1
    for div in container.find_all('div', recursive=False):
        print(i)
        i += 1
        agent_profile_link = f"https://www.remax.com{div.find(attrs=    {'data-test':'agent-card-name'}).get('href')}"
        data.append(getDetails(agent_profile_link))

    return data


response = requests.get(
    "https://www.remax.com/real-estate-agents/-ca", timeout=5)
soup = BeautifulSoup(response.content, 'html.parser')
# container = soup.find(attrs={'data-test': 'agent-card-img'}).find('a')
# profile_picture = container.find('img')
# print(profile_picture['data-src'])


totalRecords = soup.find(attrs={'class': 'roster-sort-container'}
                         ).find('h4', attrs={'class': 'mr-3'}).get_text(strip=True)
totalRecords = int(totalRecords[0:5].replace(',', ""))
totalPages = totalRecords//96+1


allData = []
for i in range(1, totalPages):
    print('url: '+str(i))
    url = f'https://www.remax.com/real-estate-agents/-ca?filters={{"page":{i},"count":"96","sortBy":"lastName"}}'
    tempData = getPageData(url)
    allData.extend(tempData)

df = pd.DataFrame(allData, columns=["AGENT NAME", "PHONE NUMBERS",
                  "WEBSITE LINK", "SOCIAL MEDIA LINKS", "LANGUAGE", "ADDRESS"])
df.to_excel("agent_data.xlsx", index=False)
