from bs4 import BeautifulSoup
import urllib.request

import json


def get_databases(server_database, security_protocol_database):
    """Scraps databases for the SMTP-Server, security protocols and ports and stores the information in two files, whose names are given by Args. Gets called by easymail.client, if the database files are not found.

    Args:
        server_database (str): Path to the location to save the scraped SMTP-Server database.
        security_protocol_database (str): Path to the location to save the scraped protocol and port database.
    """
    user_agent = "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.0.7) Gecko/2009021910 Firefox/3.0.7"
    url = "https://www.smtpsoftware.com/smtp-server-list/"
    headers = {"User-Agent": user_agent}

    request = urllib.request.Request(url, None, headers)  # The assembled request
    response = urllib.request.urlopen(request)
    data = response.read()  # The data u need
    soup = BeautifulSoup(data, "lxml")
    hosts = []
    for item in soup.find_all("h4", itemprop="text"):
        hosts.append(item.string.split(": ")[1])
    domains = []
    for item in soup.find_all("h3", itemprop="name"):
        if item.string != None:
            domains.append(item.string.split("For ")[1].split(" ")[0])
        else:
            domains.append(None)
    smtp_server_list = []
    for i in range(len(domains)):
        smtp_server_list.append([domains[i], hosts[i]])

    with open(server_database, "w") as file:
        file.write(json.dumps(smtp_server_list))

    url = "https://www.arclab.com/en/kb/email/list-of-smtp-and-pop3-servers-mailserver-list.html"
    response = urllib.request.urlopen(url)
    data = response.read()  # The data u need
    soup = BeautifulSoup(data, "lxml")

    relevant_rows = []
    tables = soup.find_all("table", {"class": "t-fine"})
    for table in tables:
        rows = table.find_all("tr")
        for row in rows:
            if "SMTP" in str(row):
                relevant_rows.append(row)

    relevant_information = []
    for row in relevant_rows:
        relevant_information_element = []
        for line in row.find_all("td"):
            if not "SMTP" in str(line.string):
                relevant_information_element.append(line.string)
        relevant_information.append(relevant_information_element)

    with open(security_protocol_database, "w") as file:
        file.write(json.dumps(relevant_information))
