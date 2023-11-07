import json
import os
import sys
from urllib.parse import urlparse

import requests
import teams_requests
from loguru import logger
from msal import PublicClientApplication

user_agent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko)"


def get_tenant_id(username):
    domain = username.split("@")[-1]

    response = requests.get(
        "https://login.microsoftonline.com/%s/.well-known/openid-configuration" % (domain), timeout=10
    )

    if response.status_code != 200:
        logger.error("Could not retrieve tenant id for domain %s" % (domain))

    json_content = json.loads(response.text)
    tenant_id = json_content.get("authorization_endpoint").split("/")[3]

    return tenant_id


def two_fa(username, scope):
    # Values hardcoded for corporate/part of organization users
    app = PublicClientApplication(
        "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
        authority="https://login.microsoftonline.com/%s" % get_tenant_id(username),
    )

    try:
        # Initiate the device code authentication flow and print instruction message
        flow = app.initiate_device_flow(scopes=[scope])
        if "user_code" not in flow:
            logger.error("Could not retrieve user code in authentication flow")
            sys.exit(1)
        logger.warning(flow.get("message"))
    except Exception:
        logger.error("Could not initiate device code authentication flow")
        sys.exit(1)

    # Initiates authentication based on the previously created flow. Polls the MS endpoint for entered device codes.
    try:
        result = app.acquire_token_by_device_flow(flow)
    except Exception as err:
        logger.error("Error while authenticating: %s" % (err.args[0]))
        sys.exit(1)

    return result


def get_sender_info(bearer):
    logger.info("Fetching sender info...")

    userID = None
    skipToken = None
    senderInfo = None

    headers = {"Authorization": "Bearer %s" % (bearer)}

    # First request fetches userID associated with our sender/bearer token
    response = requests.get("https://teams.microsoft.com/api/mt/emea/beta/users/tenants", headers=headers, timeout=10)

    if response.status_code != 200:
        logger.error("Could not retrieve senders userID!")
        sys.exit(1)

    # Store userID as well as the tenantName of our sending user
    userID = json.loads(response.text)[0].get("userId")
    json.loads(response.text)[0].get("tenantName")

    # Second, we need to find the display name associated with our userID
    # Enumerate users within sender's tenant and find our matching user
    while True:
        url = "https://teams.microsoft.com/api/mt/emea/beta/users"
        if skipToken:
            url += f"?skipToken={skipToken}&top=999"

        response = requests.get(url, headers=headers, timeout=10)

        if response.status_code != 200:
            logger.error("Could not retrieve senders display name!")
            sys.exit(1)

        users_response = json.loads(response.text)
        users = users_response["users"]
        skipToken = users_response.get("skipToken")

        # Iterate through retrieved users and find the one that matches our previously retrieved UserID.
        for user in users:
            if user.get("id") == userID:
                senderInfo = user
                break

        if senderInfo or not skipToken:
            break

    # Add tenantName to our senderInfo data for later
    # Populating tenantName by parsing UPN because ran into issues where peoples 'Organization Name' differed from their 'Initial Domain Name'
    if senderInfo:
        senderInfo["tenantName"] = senderInfo["userPrincipalName"].split("@")[-1].split(".")[0]
        logger.success("Obtained sender info.")
    else:
        logger.success("Could not find the sender's user information!")
        sys.exit(1)

    return senderInfo


def get_skype_token(bearer):
    logger.info("Fetching Skype token...")

    headers = {"Authorization": "Bearer " + bearer}

    # Requests a Skypetoken
    # https://digitalworkplace365.wordpress.com/2021/01/04/using-the-ms-teams-native-api-end-points/
    content = requests.post("https://authsvc.teams.microsoft.com/v1.0/authz", headers=headers, timeout=10)

    if content.status_code != 200:
        logger.error("Error fetching skype token: %d" % (content.status_code), True)

    json_content = json.loads(content.text)
    if "tokens" not in json_content:
        logger.error("Could not retrieve Skype token", True)

    logger.success("Obtained Skype Token!")
    return json_content.get("tokens").get("skypeToken")


def get_bearer_token(username, password, scope):
    result = None

    # If this string was passed in for scope, we are grabbing our initial Bearer token
    if scope == "https://api.spaces.skype.com/.default":
        logger.info("Fetching Bearer token for Teams...")

    # If scope doesn't match the above, we are fetching our Sharepoint Bearer
    else:
        logger.info("Fetching Bearer token for SharePoint...")
        # logger.debug(f"Using {scope} as tenant")

        # If scope was passed in as a dictionary, we are assembling our sharepoint domain automagically using the UPN from senderInfo
        if isinstance(scope, dict):
            scope = "https://%s-my.sharepoint.com/.default" % scope.get("tenantName")

        # Otherwise scope was passed in as a user-defined option
        else:
            scope = "https://%s-my.sharepoint.com/.default" % scope
        logger.info(f"Using {scope} as tenant. If file upload does not work, double-check this is correct.")

    # Values hardcoded for corporate/part of organization users
    app = PublicClientApplication(
        "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
        authority="https://login.microsoftonline.com/%s" % get_tenant_id(username),
    )
    try:
        # Initiates authentication based on credentials.
        result = app.acquire_token_by_username_password(username, password, scopes=[scope])
    except ValueError as err:
        if "This typically happens when attempting MSA accounts" in err.args[0]:
            logger.warning(
                "Username/Password authentication cannot be used with Microsoft accounts. Either use the device code"
                " authentication flow or try again with a user managed by an organization."
            )
        logger.error("Error while acquring token")
        sys.exit(1)

    # Login not successful
    if "access_token" not in result:
        if "Error validating credentials due to invalid username or password" in result.get("error_description"):
            logger.error("Invalid credentials entered")
            sys.exit(1)
        elif "This device code has expired" in result.get("error_description"):
            logger.error("The device code has expired. Please try again")
            sys.exit(1)
        elif "multi-factor authentication" in result.get("error_description"):
            result = two_fa(username, scope)
        else:
            logger.error(result.get("error_description"))
            sys.exit(1)

    logger.success("Bearer Token obtained succesfully.")
    return result["access_token"]


def create_closed_chat(bearer_token, skype_token, sender, target, chat_title):
    headers = {
        "Authentication": "skypetoken=" + skype_token,
        "Host": "teams.microsoft.com",
        "User-Agent": user_agent,
        "Content-Type": "application/json",
        "Origin": "https://teams.microsoft.com",
        "Referer": "https://teams.microsoft.com/",
    }

    body = {
        "members": [
            {"id": target.get("mri"), "role": "User", "isReader": "true", "isFollowing": "true"},
            {"id": sender.get("mri"), "role": "Admin", "isReader": "false", "hidden": "true"},
        ],
        "properties": {
            "threadType": "chat",
            "topic": chat_title,
            "isStickyThread": "false",
            "joiningenabled": "false",
            "templateType": "ClosedChat",
        },
    }

    content = requests.post(
        "https://emea.ng.msg.teams.microsoft.com/v1/threads",
        headers=headers,
        data=json.dumps(body),
        timeout=10,
    )

    return (content.status_code == 201, content)


def upload_file(sharepoint_token, sender_sharepoint_url, sender_drive, attachment):
    headers = {
        "Authorization": "Bearer " + sharepoint_token,
        "User-Agent": user_agent,
        "Content-Type": "application/json",
        "Origin": "https://teams.microsoft.com",
        "Referer": "https://teams.microsoft.com/",
    }

    attachment_name = os.path.basename(attachment)

    url = f"{sender_sharepoint_url}/personal/{sender_drive}/_api/v2.0/drive/root:/Microsoft%20Teams%20Chat%20Files/{attachment_name}:/content?@name.conflictBehavior=replace&$select=*,sharepointIds,webDavUrl"

    with open(attachment, mode="rb") as file:
        body = file.read()

    content = requests.put(
        url,
        headers=headers,
        data=body,
        timeout=10,
    )

    # Returns 201 if created for first time. 200 if replaced.
    return (content.status_code == 201 or content.status_code == 200, content)


def get_alt_invite_link(
    sharepoint_token, sender_sharepoint_url, sender_drive, sender_info, target_info, upload_id, secure_link
):
    url = f"{sender_sharepoint_url}/personal/{sender_drive}/_api/v2.0/sites/root/items/{upload_id}/driveItem/invite"

    headers = {
        "Authorization": "Bearer " + sharepoint_token,
        "User-Agent": user_agent,
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Application": "Teams_Web",
        "Sec-Ch-Ua-Platform": "",
        "Origin": "https://teams.microsoft.com",
        "Sec-Fetch-Site": "cross-site",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Dest": "empty",
        "Sec-Ch-Ua": "",
        "Prefer": "getDefaultLink,ReturnSpecificErrorCode",
        "Sec-Ch-Ua-Mobile": "?0",
        "Referer": "https://teams.microsoft.com/",
    }

    body = {"recipients": [{"email": target_info.get("userPrincipalName")}]}

    content = requests.post(
        url,
        headers=headers,
        data=json.dumps(body),
        timeout=10,
    )

    return (content.status_code == 200 or content.status_code == 201, content)


def get_invite_link(
    sharepoint_token, sender_sharepoint_url, sender_drive, sender_info, target_info, upload_id, secure_link
):
    # Assemble invite link request URL
    url = f"{sender_sharepoint_url}/personal/{sender_drive}/_api/web/GetFileById(@a1)/ListItemAllFields/ShareLink?@a1=guid%27{upload_id}%27"

    headers = {
        "Authorization": "Bearer " + sharepoint_token,
        "User-Agent": user_agent,
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "Origin": "https://www.odwebp.svc.ms",
        "Referer": "https://www.odwebp.svc.ms/",
    }

    # Define two different settings blocks for the request body depending on if we are sending a secure link or not.
    unsecure = """            "allowAnonymousAccess": true,
            "trackLinkUsers": false,
            "linkKind": 4,
            "expiration": null,
            "role": 1,
            "restrictShareMembership": false,
            "updatePassword": false,
            "password": "",
            "scope": 0"""

    secure = """            "linkKind": 6,
            "expiration": null,
            "role": 1,
            "restrictShareMembership": true,
            "updatePassword": false,
            "password": "",
            "scope": 2"""

    settings = secure if secure_link else unsecure

    # If sender and target info match, this is a test message. Use single recipient PPI
    if sender_info == target_info:
        # Stitch body together
        body = """
        {{
            "request": {{
            "createLink": true,
            "settings": {{
                {}
            }},
            "peoplePickerInput": "[{{\\"Key\\":\\"i:0#.f|membership|{}\\",\\"DisplayText\\":\\"{}\\",\\"IsResolved\\":true,\\"Description\\":\\"{}\\",\\"EntityType\\":\\"User\\",\\"EntityData\\":{{\\"IsAltSecIdPresent\\":\\"False\\",\\"Title\\":\\"\\",\\"Email\\":\\"{}\\",\\"MobilePhone\\":\\"\\",\\"ObjectId\\":\\"{}\\",\\"Department\\":\\"\\"}},\\"MultipleMatches\\":[],\\"ProviderName\\":\\"Tenant\\",\\"ProviderDisplayName\\":\\"Tenant\\"}}]"
            }}
        }}
        """.format(
            settings,
            sender_info.get("userPrincipalName"),
            sender_info.get("displayName"),
            sender_info.get("userPrincipalName"),
            sender_info.get("userPrincipalName"),
            sender_info.get("id"),
        )

    else:
        # Stitch body together
        body = """
        {{
            "request": {{
            "createLink": true,
            "settings": {{
                {}
            }},
            "peoplePickerInput": "[{{\\"Key\\":\\"i:0#.f|membership|{}\\",\\"DisplayText\\":\\"{}\\",\\"IsResolved\\":true,\\"Description\\":\\"{}\\",\\"EntityType\\":\\"User\\",\\"EntityData\\":{{\\"IsAltSecIdPresent\\":\\"False\\",\\"Title\\":\\"\\",\\"Email\\":\\"{}\\",\\"MobilePhone\\":\\"\\",\\"ObjectId\\":\\"{}\\",\\"Department\\":\\"\\"}},\\"MultipleMatches\\":[],\\"ProviderName\\":\\"Tenant\\",\\"ProviderDisplayName\\":\\"Tenant\\"}},{{\\"Key\\":\\"{}\\",\\"DisplayText\\":\\"{}\\",\\"IsResolved\\":true,\\"Description\\":\\"{}\\",\\"EntityType\\":\\"\\",\\"EntityData\\":{{\\"SPUserID\\":\\"{}\\",\\"Email\\":\\"{}\\",\\"IsBlocked\\":\\"False\\",\\"PrincipalType\\":\\"UNVALIDATED_EMAIL_ADDRESS\\",\\"AccountName\\":\\"{}\\",\\"SIPAddress\\":\\"{}\\",\\"IsBlockedOnODB\\":\\"False\\"}},\\"MultipleMatches\\":[],\\"ProviderName\\":\\"\\",\\"ProviderDisplayName\\":\\"\\"}}]"
            }}
        }}
        """.format(
            settings,
            sender_info.get("userPrincipalName"),
            sender_info.get("displayName"),
            sender_info.get("userPrincipalName"),
            sender_info.get("userPrincipalName"),
            sender_info.get("id"),
            target_info.get("userPrincipalName"),
            target_info.get("userPrincipalName"),
            target_info.get("userPrincipalName"),
            target_info.get("userPrincipalName"),
            target_info.get("userPrincipalName"),
            target_info.get("userPrincipalName"),
            target_info.get("userPrincipalName"),
        )

    # Send request
    content = requests.post(url, headers=headers, data=body, timeout=10)

    return (content.status_code == 200, content)


def get_invite_link_parser(invite_info, alternative=False):
    if alternative:
        file_url = invite_info.get("value")[0].get("link").get("webUrl")
        file_share_id = invite_info.get("value")[0].get("id")
        return (file_url, file_share_id)
    else:
        file_url = invite_info.get("d").get("ShareLink").get("sharingLinkInfo").get("Url")
        file_share_id = invite_info.get("d").get("ShareLink").get("sharingLinkInfo").get("ShareId")
        return (file_url, file_share_id)


def create_meeting_thread(bearer_token):
    headers = {
        "Authorization": "Bearer " + bearer_token,
        "Host": "teams.microsoft.com",
        "User-Agent": user_agent,
        "Content-Type": "application/json",
        "Origin": "https://teams.microsoft.com",
        "Referer": "https://teams.microsoft.com/",
    }

    body = {"isStreamEnabled": False, "meetingType": "Scheduled"}

    content = requests.post(
        "https://teams.microsoft.com/api/mt/part/emea-03/beta/me/calendarEvents/privateMeeting/schedulingService/create",
        headers=headers,
        data=json.dumps(body),
        timeout=10,
    )

    return (content.status_code == 201, content)


def closed_chat_thread_parser(content):
    parsed_url = urlparse(content.headers.get("Location"))
    return parsed_url.path.split("/")[-1]


def meeting_thread_parser(content):
    return (
        content.json().get("value").get("groupContext").get("threadId"),
        content.json().get("value").get("etag"),
        content.json().get("value").get("meetingUrl"),
        content.json().get("value").get("links"),
        content.json().get("value").get("teamsVtcTenantId"),
        content.json().get("value").get("views").get("html"),
    )


def upload_file_parser(content):
    return content.json().get("sharepointIds").get("listItemUniqueId")


def create_schedule(bearer_token, thread, sender_info, chat_title, etag, meeting_url, links, teams_vtc_tenant_id, html):
    headers = {
        "Authorization": "Bearer " + bearer_token,
        "Host": "teams.microsoft.com",
        "User-Agent": user_agent,
        "Content-Type": "application/json",
        "Origin": "https://teams.microsoft.com",
        "Referer": "https://teams.microsoft.com/",
    }

    create_event_body = teams_requests.build_create_schedule_request(
        chat_title,
        thread,
        etag.replace('"', ""),
        meeting_url,
        links,
        teams_vtc_tenant_id,
        sender_info.get("mri"),
        sender_info.get("userPrincipalName"),
        sender_info.get("displayName"),
        html,
    )

    js = json.dumps(create_event_body)

    content = requests.post(
        "https://teams.microsoft.com/api/mt/part/emea-03/v2.0/me/calendars/default/events?isOnlineMeeting=true&shouldDecryptData=true",
        headers=headers,
        data=js,
        timeout=10,
    )

    return content.status_code == 201


def chat_unhide(skype_token, thread):
    headers = {
        "Authentication": "skypetoken=" + skype_token,
        "User-Agent": user_agent,
        "Content-Type": "application/json",
        "Origin": "https://teams.microsoft.com",
        "Referer": "https://teams.microsoft.com/",
    }

    body = json.dumps({"hidden": "false"})

    content = requests.put(
        f"https://emea.ng.msg.teams.microsoft.com/v1/threads/{thread}/properties?name=hidden",
        headers=headers,
        data=body,
        timeout=10,
    )

    return content.status_code == 200


def debug_post(req):
    """
    At this point it is completely built and ready
    to be fired; it is "prepared".

    However pay attention at the formatting used in
    this function because it is programmed to be pretty
    printed and may differ from the actual request.
    """
    print(
        "{}\n{}\r\n{}\r\n\r\n{}".format(
            "-----------START-----------",
            req.method + " " + req.url,
            "\r\n".join(f"{k}: {v}" for k, v in req.headers.items()),
            req.body,
        )
    )


def chat_add_member(skype_token, thread, target):
    headers = {
        "Authentication": "skypetoken=" + skype_token,
        "User-Agent": user_agent,
        "Content-Type": "application/json",
        "Origin": "https://teams.microsoft.com",
        "Referer": "https://teams.microsoft.com/",
    }

    body = json.dumps({"members": [{"id": target.get("mri"), "role": "Admin", "shareHistoryTime": -1}]})

    content = requests.post(
        f"https://emea.ng.msg.teams.microsoft.com/v1/threads/{thread}/members", headers=headers, data=body, timeout=10
    )

    return content.status_code == 201


def chat_create_closed_chat(bearer_token, skype_token, sender, target, chat_title):
    r, content = create_closed_chat(bearer_token, skype_token, sender, target, chat_title)
    if not r:
        logger.error("Could not create closed chat!")
        sys.exit(1)

    thread = closed_chat_thread_parser(content)

    if not chat_unhide(skype_token, thread):
        logger.error("Could not unhide chat!")
        sys.exit(1)

    return thread


def chat_create_meeting(bearer_token, skype_token, sender, target, chat_title):
    r, content = create_meeting_thread(bearer_token)

    if not r:
        logger.error("Could not create meeting chat!")
        sys.exit(1)

    thread, *meeting_args = meeting_thread_parser(content)

    if not thread:
        logger.error("Could not read thread, from create_meeting_thread response")

    if not create_schedule(bearer_token, thread, sender, chat_title, *meeting_args):
        logger.error("Could not create schedule!")
        sys.exit(1)

    if sender != target and not chat_add_member(skype_token, thread, target):
        logger.error("Could not add member to schedule chat!")
        sys.exit(1)

    if not chat_unhide(skype_token, thread):
        logger.error("Could not unhide chat!")
        sys.exit(1)

    return thread


def chat_send_msg(skype_token, thread, msg):
    url = "https://amer.ng.msg.teams.microsoft.com/v1/users/ME/conversations/" + thread + "/messages"
    headers = {
        "Authentication": "skypetoken=" + skype_token,
        "User-Agent": user_agent,
        "Content-Type": "application/json, Charset=UTF-8",
        "Origin": "https://teams.microsoft.com",
        "Referer": "https://teams.microsoft.com/",
    }

    body = {
        "content": msg,
        "messagetype": "RichText/Html",
        "contenttype": "text",
        "amsreferences": [],
        "clientmessageid": "3529890327684204437",
        "imdisplayname": "it iam",
        "properties": {"importance": "", "subject": ""},
    }

    content = requests.post(url, headers=headers, data=json.dumps(body).encode(encoding="utf-8"), timeout=10)

    return content.status_code == 201


def chat_send_msg_with_file(
    skype_token, thread, msg, sender_sharepoint_url, sender_drive, upload_info, file_url, file_share_id
):
    url = "https://amer.ng.msg.teams.microsoft.com/v1/users/ME/conversations/" + thread + "/messages"
    headers = {
        "Authentication": "skypetoken=" + skype_token,
        "User-Agent": user_agent,
        "Content-Type": "application/json, Charset=UTF-8",
        "Origin": "https://teams.microsoft.com",
        "Referer": "https://teams.microsoft.com/",
    }

    # logger.info(invite_info)

    files = [
        {
            "@type": "http://schema.skype.com/File",
            "version": 2,
            "id": upload_info.get("sharepointIds").get("listItemUniqueId"),
            "baseUrl": f"{sender_sharepoint_url}/personal/{sender_drive}/",
            "type": upload_info.get("webUrl").split(".")[-1],
            "title": upload_info.get("webUrl").split("/")[-1],
            "state": "active",
            "objectUrl": upload_info.get("webUrl"),
            "providerData": "",
            "itemid": upload_info.get("sharepointIds").get("listItemUniqueId"),
            "fileName": upload_info.get("webUrl").split("/")[-1],
            "fileType": upload_info.get("webUrl").split(".")[-1],
            "fileInfo": {
                "itemId": None,
                "fileUrl": file_url,
                "siteUrl": f"{sender_sharepoint_url}/personal/{sender_drive}/",
                "serverRelativeUrl": "",
                "shareUrl": file_url,
                "shareId": file_share_id,
            },
            "botFileProperties": {},
            "permissionScope": "anonymous",
            "filePreview": {},
            "fileChicletState": {"serviceName": "p2p", "state": "active"},
        }
    ]

    body = {
        "content": msg,
        "messagetype": "RichText/Html",
        "contenttype": "text",
        "amsreferences": [],
        "clientmessageid": "3529890327684204437",
        "imdisplayname": "it iam",
        "properties": {"files": json.dumps(files), "importance": "", "subject": ""},
    }

    content = requests.post(url, headers=headers, data=json.dumps(body).encode(encoding="utf-8"), timeout=10)

    return content.status_code == 201, files


def file_send_invite(sharepoint_token, sender_sharepoint_url, sender_drive, sender_info, target_info, upload_info):
    upload_id = upload_info.get("sharepointIds").get("listItemUniqueId")

    alternative = False

    if not alternative:
        r, content = get_invite_link(
            sharepoint_token, sender_sharepoint_url, sender_drive, sender_info, target_info, upload_id, False
        )
    else:
        r, content = get_alt_invite_link(
            sharepoint_token, sender_sharepoint_url, sender_drive, sender_info, target_info, upload_id, False
        )

    if not r:
        logger.error("Could not send invite link!")
        sys.exit(1)

    logger.success("File invite sent succesfully!")

    return get_invite_link_parser(content.json(), alternative)


def file_upload(sharepoint_token, sender_sharepoint_url, sender_drive, attachment, sender_info):
    logger.info(f"Uploading [{attachment}] ...")
    r, content = upload_file(sharepoint_token, sender_sharepoint_url, sender_drive, attachment)

    if not r:
        logger.error("Could not upload file!")
        sys.exit(1)

    upload_info = content.json()
    logger.success("File uploaded succesfully!")

    return upload_info


def authenticate(username, password, sharepoint=None):
    bToken = get_bearer_token(username, password, "https://api.spaces.skype.com/.default")
    skypeToken = get_skype_token(bToken)
    senderInfo = get_sender_info(bToken)

    # Fetch sharepointToken passing in alternate vars for scope depending on whether specified a specific sharepoint domain to use.
    if sharepoint:
        sharepointToken = get_bearer_token(username, password, sharepoint)
    else:
        sharepointToken = get_bearer_token(username, password, senderInfo)

    return bToken, skypeToken, sharepointToken, senderInfo
