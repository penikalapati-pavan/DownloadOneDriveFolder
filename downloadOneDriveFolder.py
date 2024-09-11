import asyncio
import os
import aiofiles
import argparse
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.search.query.query_post_request_body import QueryPostRequestBody
from msgraph.generated.models.search_request import SearchRequest
from msgraph.generated.models.entity_type import EntityType
from msgraph.generated.models.search_query import SearchQuery


class OneDriveDownloader:
    def __init__(self, client_id, client_secret, tenant_id, download_path):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.download_path = download_path
        self.graph_client = None
        self.authenticate()

    def authenticate(self):
        """Authenticate with Azure AD and initialize the GraphServiceClient."""
        credentials = ClientSecretCredential(self.tenant_id, self.client_id, self.client_secret)
        scopes = ['https://graph.microsoft.com/.default']
        self.graph_client = GraphServiceClient(credentials, scopes)

    async def download_file(self, drive_id, file_item_id, filename, directory_path):
        """Download a single file from OneDrive."""
        content = await self.graph_client.drives.by_drive_id(drive_id).items.by_drive_item_id(file_item_id).content.get()
        file_path = os.path.join(directory_path, filename)
        async with aiofiles.open(file_path, 'wb') as f:
            await f.write(content)
        print(f"Downloaded file: {filename}")

    async def download_folder(self, drive_id, folder_item_id, parent_directory):
        """Download all files and folders in a given OneDrive folder."""
        items = await self.graph_client.drives.by_drive_id(drive_id).items.by_drive_item_id(folder_item_id).children.get()

        for item in items.value:
            if item.folder:  # If it's a folder, recursively download its contents
                new_folder_path = os.path.join(parent_directory, item.name)
                os.makedirs(new_folder_path, exist_ok=True)
                await self.download_folder(drive_id, item.id, new_folder_path)
            elif item.file:  # If it's a file, download it
                await self.download_file(drive_id, item.id, item.name, parent_directory)

    async def search_and_download(self, folder_name, web_url):
        """Search for a folder in OneDrive, verify webUrl, and download its contents."""
        try:

            # Construct the search request body
            request_body = QueryPostRequestBody(
                requests=[
                    SearchRequest(
                        entity_types=[EntityType.DriveItem],
                        region="IND",
                        query=SearchQuery(query_string=folder_name),
                        from_=0,
                        size=25,
                    )
                ]
            )

            result = await self.graph_client.search.query.post(request_body)

            folder_item_id, drive_id = None, None

            # print(result)

            for search_response in result.value:
                for hit_container in search_response.hits_containers:
                    for hit in hit_container.hits:
                        print(hit.resource.parent_reference.drive_id, ' ', hit.resource.name)
                        if hit.resource.name == folder_name and hit.resource.web_url == web_url :
                            folder_item_id = hit.resource.id
                            drive_id = hit.resource.parent_reference.drive_id
                            print(f"Matching folder found. Folder Item ID: {folder_item_id}, Drive ID: {drive_id}")

            if folder_item_id and drive_id:
                await self.download_folder(drive_id, folder_item_id, os.path.join(self.download_path, folder_name))
            else:
                print(f"No matching folder found for query: {folder_name} and webUrl: {web_url}")

        except Exception as e:
            print(f"Error while searching for folder : {str(e)}")


class OneDriveSearchDownloadApp:
    def __init__(self):
        self.args = self.parse_arguments()

    
    def parse_arguments(self):
        """Parse command-line arguments with optional flags."""
        parser = argparse.ArgumentParser(description="Download a folder from OneDrive.")

        # Make all flags optional, but set 'required=True' for client_id, client_secret, tenant_id
        parser.add_argument('--client_id', required=True, help="Azure AD Application client ID")
        parser.add_argument('--client_secret', required=True, help="Azure AD Application client secret")
        parser.add_argument('--tenant_id', required=True, help="Azure AD tenant ID")

        # download_path is optional and defaults to the current directory if not provided
        parser.add_argument('--download_path', default=os.getcwd(), help="Local path to save the downloaded folder (defaults to current directory)")

        # Optional folder_name and username
        parser.add_argument('--folder_name', required=True, help="Name of the folder to download")
        parser.add_argument('--web_url', required=True, help="webUrl of the OneDrive folder (can get this form details of the folder in onedrive)")

        return parser.parse_args()

    def run(self):
        """Start the download process."""
        downloader = OneDriveDownloader(
            self.args.client_id,
            self.args.client_secret,
            self.args.tenant_id,
            self.args.download_path
        )
        asyncio.run(downloader.search_and_download(self.args.folder_name, self.args.web_url))


if __name__ == '__main__':
    app = OneDriveSearchDownloadApp()
    app.run()
