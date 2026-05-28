import requests
import json
import os
from ftplib import FTP

AVAIBLE_SERVICES = ("tinyurl", "infinityfree")

class Redirect:
    """
    ### Parameters:
    - url (str): The URL to redirect to.
    - name (str): The name of the redirect. This will be used as the alias for the tiny URL and the redirect path for the InfinityFree server.
    """
    def __init__(self, url, name):
        self.url = url
        self.name = name

    def tinyUrl(self):
        """
        # Not finished!
        Returns the tiny URL for the provided URL.
        """
        url = f"https://api.tinyurl.com/create?api_token={os.getenv('TINYURL')}"
        data = {
            "url": self.url,
            "domain": "tinyurl.com",
            "alias": self.name,
            "tags": self.name,
            "expires_at": None,
            "description": None
        }
        response = requests.post(url, data=json.dumps(data))
        response.raise_for_status()
        return response.text
    def infinityfree(self, infinityfreeURL):
        """
        Updates the .htaccess file on the InfinityFree server to add a redirect rule.
        ### Parameters:
        - infinityfreeURL (str): The URL of the InfinityFree website.
        """
        ftp = FTP(os.getenv("INFINITYFREE_HOST"), os.getenv("INFINITYFREE_USERNAME"), os.getenv("INFINITYFREE_PASSWORD"))
        file_path = "./htdocs/.htaccess"
        update = f"Redirect 301 /o/{self.name} {self.url}\n" # o for OneDrive

        # Download the .htaccess file
        with open("local_htaccess", "wb") as local_file:
            ftp.retrbinary(f"RETR {file_path}", local_file.write)

        # Read the downloaded file and append the new redirect rule
        with open("local_htaccess", "a") as local_file:
            local_file.write(update)

        # Upload the modified file back to the server
        with open("local_htaccess", "rb") as local_file:
            ftp.storbinary(f"STOR {file_path}", local_file)

        # Clean up the local file
        os.remove("local_htaccess")

        ftp.quit()
        return f"{infinityfreeURL}/o/{self.name}"