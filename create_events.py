"""
Create test events across multiple user calendars using application permissions.
Requires AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET.
"""

import asyncio
import os
from datetime import datetime, timedelta
from typing import List, Dict

from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.event import Event
from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
from dotenv import load_dotenv

load_dotenv()


def load_test_attendees() -> List[str]:
    raw_value = os.getenv("TEST_ATTENDEES", "")
    attendees = [item.strip() for item in raw_value.split(",") if item.strip()]
    if not attendees:
        raise ValueError("TEST_ATTENDEES is required. Provide a comma-separated list in .env.")
    return attendees


def get_app_client() -> GraphServiceClient:
    tenant_id = os.getenv("AZURE_TENANT_ID")
    client_id = os.getenv("AZURE_CLIENT_ID")
    client_secret = os.getenv("AZURE_CLIENT_SECRET")
    credential = ClientSecretCredential(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
    )
    return GraphServiceClient(credentials=credential, scopes=["https://graph.microsoft.com/.default"])


def generate_time_slots(start_date: datetime, num_days: int = 30, slot_minutes: int = 30) -> List[Dict]:
    """Generate test time slots covering all hours (unrestricted)."""
    slots = []
    current = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    for day in range(num_days):
        date = current + timedelta(days=day)
        for start_hour in range(0, 24, 1):
            start_dt = date.replace(hour=start_hour, minute=0)
            end_dt = start_dt + timedelta(minutes=slot_minutes)
            slots.append({
                "start": start_dt.isoformat(),
                "end": end_dt.isoformat(),
            })
    return slots


async def create_test_events(client: GraphServiceClient, attendees: List[str], slots: List[Dict], subject_prefix="TEST_502_", log_every: int = 50):
    created = 0
    errors = 0

    for attendee in attendees:
        per_user_created = 0
        per_user_errors = 0
        print(f"Creating events for {attendee} ({len(slots)} slots)...")
        for i, slot in enumerate(slots, start=1):
            try:
                event = Event()
                event.subject = f"{subject_prefix}{attendee.split('@')[0]}_{i}"

                start_time = DateTimeTimeZone()
                start_time.date_time = slot["start"]
                start_time.time_zone = "UTC"
                event.start = start_time

                end_time = DateTimeTimeZone()
                end_time.date_time = slot["end"]
                end_time.time_zone = "UTC"
                event.end = end_time

                await client.users.by_user_id(attendee).events.post(event)
                created += 1
                per_user_created += 1
            except Exception:
                errors += 1
                per_user_errors += 1

            if i % log_every == 0:
                print(f"  {attendee}: {i}/{len(slots)} processed (ok={per_user_created}, err={per_user_errors})")

        print(f"  Done {attendee}: ok={per_user_created}, err={per_user_errors}")

    return {"created": created, "errors": errors}


async def main():
    client = get_app_client()
    slots = generate_time_slots(datetime.now(), num_days=60, slot_minutes=30)
    attendees = load_test_attendees()
    result = await create_test_events(client, attendees, slots, log_every=50)
    print(f"Created: {result['created']} events, Errors: {result['errors']}")


if __name__ == "__main__":
    asyncio.run(main())
