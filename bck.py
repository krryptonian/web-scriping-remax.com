from time import sleep
import pandas as pd
from bs4 import BeautifulSoup
import requests
from openpyxl import load_workbook

def make_hyperlink(link,text):
    return '=HYPERLINK("%s", "%s")' % (link, text)

def getDetails(profile_link):
    response = requests.get(profile_link, timeout=10)
    soup = BeautifulSoup(response.content, 'html.parser')
    try:
        main = soup.find(attrs={'data-test': 'agent-bio'})
        container = main.find('div',attrs={'class':'noprint'})

        if main:
            # get agent name
            agent_name = container.find('h1',attrs={'class':'h2 mt-6'}).get_text(strip=True)
            print(agent_name)
            # get profile picture

            # profileContainer = container.find(attrs={'class': 'col md:max-w-1/2 lg:max-w-full'})
            # profile_picture = profileContainer.find('img')
            # print(profile_picture['data-src'])

            #old profile picture code
            # profile_picture = container.find('img',attrs={'class':'mb-4 photo-max-height'}).get('src') if container.find('img',attrs={'class':'mb-4   photo-max-height'}) else ""
            # get phone numbers

            #phones_container = container.find('div', attrs={'class': 'bio-phone mb-8 lg:mb-0'})
            #agent_phones = ';'.join([phone.get_text(strip=True) for phone in phones_container.find_all('h4')] if phones_container else [])

            # phones_container.find_next('a').get_text(strip=True)

            phones_container = container.find('div', attrs={'class': 'bio-phone mb-8 lg:mb-0'})
            phoneNo = ""
            telephoneNo = ""
            
            if(phones_container):
                i = 0
                for phone in phones_container.find_all('span'):
                        
                    data = phone.find('path', attrs={'d': 'M6.62 10.79c1.44 2.83 3.76 5.14 6.59 6.59l2.2-2.2c.27-.27.67-.36 1.02-.24 1.12.37 2.33.57 3.57.57.55 0 1 .45 1 1V20c0 .55-.45 1-1 1-9.39 0-17-7.61-17-17 0-.55.45-1 1-1h3.5c.55 0 1 .45 1 1 0 1.25.2 2.45.57 3.57.11.35.03.74-.25 1.02l-2.2 2.2z'})
                    if(data):
                        #telephoneNo = ';'.join([phone.get_text(strip=True) for phone in phones_container.find('a')] if phones_container else [])
                        if phones_container:    
                            telephoneNo += ", "+str(phones_container.find_all('a')[i].get_text(strip=True))
                        i+=1
                    else:
                        if phones_container:    
                            phoneNo += ", "+str(phones_container.find_all('a')[i].get_text(strip=True))                           
                        i+=1
                        #phoneNo = ', '.join([phone.get_text(strip=True) for phone in phones_container.find('a')] if phones_container else [])

            print(profile_link)
            print('Phone no')
            print(phoneNo)
            print("telephoneNo")
            print(telephoneNo)
            
            # agent_phones = ';'.join([phone.get_text(strip=True) for phone in phones_container.find_all('h4')] if phones_container else [])
            # Get agent website
            website_link = container.find('a', class_='my-website').get('href', '') if container.find('a', class_='my-website') else ""
            # get languages
            agent_language = container.find('p', attrs={'data-test': 'bio-languages'}).find('span').get_text(strip=True) if container.find('p', attrs=  {'data-test': 'bio-languages'}) else ""

            

            # Find address container
            address_container = container.find('a', class_='inline-block directions-link')
            if not address_container:
                print("Address not found")
                return
            # Get address

          

            address = address_container.get_text(strip=True)
            # Get location link
            location_link = address_container.get('href', '')

            # Find social network container
            social_links = ''
            social_container = container.find('div', class_='mb-12').find_all('a', class_='social-icon') if container.find('div', class_='mb-12') else []
            for social in social_container:
                social_links += f"{social.get('href')}\n"
            
            #print( [agent_name, phoneNo,telephoneNo, website_link, social_links, agent_language, make_hyperlink(location_link, address)])
            return [agent_name, phoneNo,telephoneNo, website_link, social_links, agent_language, make_hyperlink(location_link, address)]
        else:
            return ["Agent not found", "", "","", "", "", ""]
        
    # except requests.exceptions.RequestException as e:
    except:
        #print('container not found!')
        return ["Agent not found", "", "", "","", "", ""]

def getPageData(url):
    response = requests.get(url,timeout=5)
    soup = BeautifulSoup(response.content, 'html.parser')
    container = soup.find(attrs={'class': 'roster-container'})

    i=1
    data = []
    for div in container.find_all('div', recursive=False):
        print(i)
        i+=1
        agent_profile_link = f"https://www.remax.com{div.find(attrs={'data-test':'agent-card-name'}).get('href')}"
        data.append(getDetails(agent_profile_link))
    
    wb_append = load_workbook("agent_data.xlsx")
    sheet = wb_append.active
    print(data)
    for i in data:
        sheet.append(i)
    wb_append.save('agent_data.xlsx')
    print('done')
    sleep(5)

    # df = pd.DataFrame(data, columns=["AGENT NAME", "PHONE NUMBERS", "WEBSITE LINK", "SOCIAL MEDIA LINKS", "LANGUAGE", "ADDRESS"])
    # df.to_excel("agent_temp_data.xlsx", index=False)

    # agent_data = pd.read_excel('agent_data2.xlsx')
    # agent_temp_data = pd.read_excel('agent_temp_data.xlsx')
   
    # combine_data = pd.concat([agent_data,agent_temp_data], index=False)
    # combine_data.to_excel("agent_data2.xlsx", index=False)

response = requests.get("https://www.remax.com/real-estate-agents/-ca", timeout=10)
soup = BeautifulSoup(response.content, 'html.parser')    
totalRecords = soup.find(attrs={'class': 'roster-sort-container'}).find('h4',attrs={'class':'mr-3'}).get_text(strip=True)
totalRecords = int(totalRecords[0:5].replace(',',""))
totalPages = totalRecords//96+1


allData = []
# try:
#     df = pd.DataFrame(allData, columns=["AGENT NAME", "PHONE NUMBERS","TELEPHONES NUMBERS", "WEBSITE LINK", "SOCIAL MEDIA LINKS", "LANGUAGE", "ADDRESS"])
#     df.to_excel("agent_data.xlsx", index=False)
# except:
#     print("File is already created")

for i in range(1,totalPages+1):
    print("page: "+str(i))
    url = f'https://www.remax.com/real-estate-agents/-ca?filters={{"page":{i},"count":"96","sortBy":"lastName"}}'
    getPageData(url)
    