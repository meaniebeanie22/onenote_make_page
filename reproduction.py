from azure.identity import DeviceCodeCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.onenote_page import OnenotePage
from msgraph.generated.users.item.onenote.sections.item.pages.pages_request_builder import (
    PagesRequestBuilder,
)
import asyncio
from pprint import pprint

AZURE_SETTINGS = {
    "clientId": "6eb5c9aa-3e3b-4e84-aaf2-43e645d31fd7",
    "tenantId": "348841ce-b851-46d7-a6b2-96bba96fea69",
    "graphUserScopes": "User.Read Notes.ReadWrite.All Notes.Create",
}

SECTION_ID = "1-66e372e8-2950-460a-92cc-d0ee26dccdc7"

class Graph:
    settings: dict
    device_code_credential: DeviceCodeCredential
    graph_client: GraphServiceClient

    def __init__(self, config: dict):
        self.settings = config
        client_id = self.settings["clientId"]
        tenant_id = self.settings["tenantId"]
        graph_scopes = self.settings["graphUserScopes"].split(" ")

        self.device_code_credential = DeviceCodeCredential(
            client_id, tenant_id=tenant_id
        )
        self.graph_client = GraphServiceClient(
            self.device_code_credential, graph_scopes
        )

    async def make_page_request_info(self, contents, title):
        request_config = (
            PagesRequestBuilder.PagesRequestBuilderPostRequestConfiguration()
        )
        request_config.headers.add("Content-Type", "application/xhtml+xml")
        request_config.headers.add("boundary", "MyPartBoundary")

        page = OnenotePage(content=contents, title=title)
        page_post_request = self.graph_client.me.onenote.sections.by_onenote_section_id(
            SECTION_ID
        ).pages.to_post_request_information(page, request_config) # this points to a section in my onenote
        return page_post_request
    
    async def make_page(self, contents, title):
        request_config = (
            PagesRequestBuilder.PagesRequestBuilderPostRequestConfiguration()
        )
        request_config.headers.add("Content-Type", "application/xhtml+xml")
        request_config.headers.add("boundary", "MyPartBoundary")
        
        page = OnenotePage(content=contents, title=title)
        posted_page = self.graph_client.me.onenote.sections.by_onenote_section_id(
            "1-66e372e8-2950-460a-92cc-d0ee26dccdc7"
        ).pages.post(page, request_config) # this points to a section in my onenote
        return await posted_page

page_html = """
<!DOCTYPE html>
<html lang="en-AU">
        <head>
                <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
                <title>Bug test!</title>
        </head>
        <body data-absolute-enabled="true" style="font-family:Calibri;font-size:11pt">
                <p>Testing attempt</p>
        </body>
</html>
"""

async def main():
    graph: Graph = Graph(AZURE_SETTINGS)
    make_page_request = await graph.make_page_request_info(page_html, "Onenote make page test")
    made_page = await graph.make_page(page_html, "onenote make page test")
    print(f"Request content:\n{make_page_request.content}\n")
    pprint(made_page)

if __name__ == "__main__":
    asyncio.run(main())

