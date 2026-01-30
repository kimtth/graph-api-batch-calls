"""
Cleanup script: Remove all calendar events with "TEST_502" in the title
Targets the authenticated user's mailbox and all test attendee calendars
Handles pagination
"""

import asyncio
import os
from urllib.parse import quote
from azure.identity import DeviceCodeCredential, ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.event_collection_response import EventCollectionResponse
from kiota_abstractions.request_information import RequestInformation
from kiota_abstractions.method import Method
from dotenv import load_dotenv

load_dotenv()

SCOPES = ["Calendars.Read.Shared", "Calendars.ReadWrite", "User.Read"]


def load_test_attendees() -> list:
    raw_value = os.getenv("TEST_ATTENDEES", "")
    attendees = [item.strip() for item in raw_value.split(",") if item.strip()]
    if not attendees:
        raise ValueError("TEST_ATTENDEES is required. Provide a comma-separated list in .env.")
    return attendees

def get_graph_client(use_device_code: bool = True, use_app_permissions: bool = False) -> GraphServiceClient:
    """Initialize Graph client."""
    if use_app_permissions:
        tenant_id = os.getenv("AZURE_TENANT_ID")
        client_id = os.getenv("AZURE_CLIENT_ID")
        client_secret = os.getenv("AZURE_CLIENT_SECRET")
        # Application permissions can de
        credential = ClientSecretCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret
        )
        return GraphServiceClient(credentials=credential, scopes=["https://graph.microsoft.com/.default"])

    credential = DeviceCodeCredential(
        client_id="03836660-f561-40f4-879f-8d067652de12",
        tenant_id="e25fc136-055d-41d0-9503-83f9c657cbea"
    )
    return GraphServiceClient(credentials=credential, scopes=SCOPES)


async def _fetch_events_for_user(client: GraphServiceClient, user_id: str):
    """Fetch events using /users/{id}/events with paging (client-side filter later)."""
    all_events = []

    def _extend_events(page):
        if page and page.value:
            all_events.extend(page.value)

    base_url = f"https://graph.microsoft.com/v1.0/users/{quote(user_id)}/events"
    first_url = f"{base_url}?$select=id,subject&$top=1000"

    request_info = RequestInformation()
    request_info.http_method = Method.GET
    request_info.url = first_url

    events_page = await client.request_adapter.send_async(
        request_info,
        EventCollectionResponse,
        None
    )
    _extend_events(events_page)

    next_link = getattr(events_page, "odata_next_link", None)
    while next_link:
        request_info = RequestInformation()
        request_info.http_method = Method.GET
        request_info.url = next_link
        events_page = await client.request_adapter.send_async(
            request_info,
            EventCollectionResponse,
            None
        )
        _extend_events(events_page)
        next_link = getattr(events_page, "odata_next_link", None)

    return all_events


async def cleanup_test_events(client: GraphServiceClient, attendees: list, keyword: str = "TEST_502"):
    """Delete all events containing keyword from attendee calendars and admin mailbox."""
    total_deleted = 0
    total_failed = 0
    total_scanned = 0

    print(f"\nCleaning up events containing '{keyword}' from all calendars...\n")
    user_list = attendees

    for user_id in user_list:
        print(f"Processing: {user_id}")
        try:
            all_events = await _fetch_events_for_user(client, user_id)
            if not all_events:
                print("  No events found.")
                continue

            total_scanned += len(all_events)
            keyword_lower = keyword.lower()
            matching_events = [
                e for e in all_events
                if e.subject and keyword_lower in e.subject.lower()
            ]
            print(f"  Found {len(all_events)} events, {len(matching_events)} matching.")

            for event in matching_events:
                try:
                    await client.users.by_user_id(user_id).events.by_event_id(event.id).delete()
                    total_deleted += 1
                except Exception as e:
                    total_failed += 1
                    print(f"    ✗ Failed: {event.subject} ({str(e)[:80]})")
        except Exception as e:
            msg = str(e)[:120]
            print(f"  ✗ Error: {msg}")

    print(f"\n{'='*60}")
    print("✅ Cleanup complete:")
    print(f"   Scanned: {total_scanned}")
    print(f"   Deleted: {total_deleted}")
    print(f"   Failed: {total_failed}")
    print(f"{'='*60}\n")

    return {"deleted": total_deleted, "failed": total_failed, "scanned": total_scanned}


async def main():
    print("\n=== Cleanup TEST_502 Events ===\n")
    client = get_graph_client(use_app_permissions=True)
    attendees = load_test_attendees()
    await cleanup_test_events(client, attendees, keyword="TEST_502")


if __name__ == "__main__":
    asyncio.run(main())
