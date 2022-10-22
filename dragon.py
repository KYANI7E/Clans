import requests

class Dragon:

    def __init__(self, token, clanTag):
        self.token = token
        self.clanTag = clanTag
        self.headers = {
            'Accept': "application/json",
            'Authorization': "Bearer " + self.token
        }
        
    def getClanRaids(self, clanTag=None):
        if clanTag == None:
            clanTag = self.clanTag
        print("Fetching clan raid info for " + self.clanTag)
            
        uri = "/clans/" + clanTag + "/capitalraidseasons"
        api_endpoint = "https://api.clashofclans.com/v1"

        url = api_endpoint + uri

        try:
            response = requests.get(url, headers=self.headers, timeout=30)
            return (response.json(), response.status_code)
        except:
            if 400 <= response.status_code <= 599:
                return "Error {}".format(response.status_code)

    def getClanInfo(self, clanTag=None):
        if clanTag == None:
            clanTag = self.clanTag
        
        print("Fectching clan info for " + self.clanTag)

        uri = "/clans/" + clanTag
        api_endpoint = "https://api.clashofclans.com/v1"

        url = api_endpoint + uri

        try:
            response = requests.get(url, headers=self.headers, timeout=30)
            return (response.json(), response.status_code)
        except:
            if 400 <= response.status_code <= 599:
                return "Error {}".format(response.status_code)
    
    def getClanWarInfo(self, clanTag=None):
        if clanTag == None:
            clanTag = self.clanTag

        print("Fetching clan war info for " + self.clanTag)

        uri = clanTag + "/currentwar"
        api_endpoint = "https://api.clashofclans.com/v1/clans/"

        url = api_endpoint + uri

        try:
            response = requests.get(url, headers=self.headers, timeout=30)
            return (response.json(), response.status_code)
        except:
            if 400 <= response.status_code <= 599:
                return "Error {}".format(response.status_code)