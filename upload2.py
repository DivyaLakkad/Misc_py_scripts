from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.runtime.utilities.request_options import RequestOptions

from office365.sharepoint.file_creation_information import FileCreationInformation
from settings import settings
import requests
import os
from os.path import basename


ctx_auth = AuthenticationContext(url=settings['https://grahamcanada.sharepoint.com/sites/Quinn/Shared%20Documents/Forms/AllItems.aspx?newTargetListUrl=%2Fsites%2FQuinn%2FShared%20Documents&viewpath=%2Fsites%2FQuinn%2FShared%20Documents%2FForms%2FAllItems%2Easpx&viewid=02a5b6d0%2D0cdc%2D4d4b%2D9914%2Dc35fb56e65c4&id=%2Fsites%2FQuinn%2FShared%20Documents%2FScaffold%5Ftest%2FIn%5Ftest'])
if ctx_auth.acquire_token_for_user(username=settings['basharatj@graham.ca'], password=settings['grahamindustrial']):
    upload_binary_file("C:\Users\divyal\Downloads\ilovepdf_merged.pdf",ctx_auth)

def upload_binary_file(file_path, ctx_auth):
    """Attempt to upload a binary file to SharePoint"""

    base_url = settings['url']
    folder_url = "MyFolder"
    file_name = basename(file_path)
    files_url ="{0}/_api/web/GetFolderByServerRelativeUrl('{1}')/Files/add(url='{2}', overwrite=true)"
    full_url = files_url.format(base_url, folder_url, file_name)

    options = RequestOptions(settings['url'])
    context = ClientContext(settings['url'], ctx_auth)
    context.request_form_digest()

    options.set_header('Accept', 'application/json; odata=verbose')
    options.set_header('Content-Type', 'application/octet-stream')
    options.set_header('Content-Length', str(os.path.getsize(file_path)))
    options.set_header('X-RequestDigest', context.contextWebInformation.form_digest_value)
    options.method = 'POST'

    with open(file_path, 'rb') as outfile:

        # instead of executing the query directly, we'll try to go around
        # and set the json data explicitly

        context.authenticate_request(options)

        data = requests.post(url=full_url, data=outfile, headers=options.headers, auth=options.auth)

        if data.status_code == 200:
            # our file has uploaded successfully
            # let's return the URL
            base_site = data.json()['d']['Properties']['__deferred']['uri'].split("/sites")[0]
            relative_url = data.json()['d']['ServerRelativeUrl'].replace(' ', '%20')

            return base_site + relative_url
        else:
            return data.json()['error']