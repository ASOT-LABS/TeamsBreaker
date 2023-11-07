import json
import sys

import requests
from loguru import logger


class TeamsUser:
    def __init__(self, bearer_token, username):
        self.btoken = bearer_token
        self.username = username
        self.useragent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko)"

    def get_status(self):
        headers = {
            "Authorization": "Bearer " + self.btoken,
            "X-Ms-Client-Version": "1415/1.0.0.2023031528",
            "User-Agent": self.useragent,
        }

        content = requests.get(
            "https://teams.microsoft.com/api/mt/emea/beta/users/%s/externalsearchv3?includeTFLUsers=true"
            % (self.username),
            headers=headers,
            timeout=10,
        )

        if content.status_code == 403:
            logger.warning(
                "User exists but the target tenant or your tenant disallow communication to external domains."
            )
            return None

        if content.status_code == 401:
            logger.error("Unable to enumerate user. Is the access token valid?")
            sys.exit(1)

        if content.status_code != 200 or (content.status_code == 200 and len(content.text) < 3):
            logger.warning(
                "Unable to enumerate user. User does not exist, is not Teams-enrolled, is part of senders tenant, or is"
                " configured to not appear in search results."
            )
            return None

        user_profile = json.loads(content.text)[0]
        if "sfb" in user_profile["mri"]:
            logger.warning("This user has a Skype for Business subscription and cannot be sent files.")
            return None
        else:
            return user_profile

    def check_teams_presence(self, mri):
        """
        Checks the presence of a user, using the teams.microsoft.com endpoint

        Args:
           mri (str): MRI of the user that should be checked

        Returns:
           Presence data structure (dict): Structure containing presence information about the targeted user
        """
        headers = {
            "Content-Type": "application/json",
            "Authorization": "Bearer " + self.btoken,
        }

        payload = [{"mri": mri}]

        content = requests.post(
            "https://presence.teams.microsoft.com/v1/presence/getpresence/", headers=headers, json=payload, timeout=10
        )

        if content.status_code != 200:
            logger.warning("Error: %d" % (content.status_code))
            return

        json_content = json.loads(content.text)
        return json_content
