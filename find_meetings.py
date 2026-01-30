from datetime import datetime, timedelta
from typing import List, Dict
from pathlib import Path
import asyncio
import json
import time
import os
from azure.identity import DeviceCodeCredential, InteractiveBrowserCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.item.find_meeting_times.find_meeting_times_post_request_body import FindMeetingTimesPostRequestBody
from msgraph.generated.models.attendee_base import AttendeeBase
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.time_constraint import TimeConstraint
from msgraph.generated.models.time_slot import TimeSlot
from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
from msgraph.generated.models.attendee_type import AttendeeType
from msgraph.generated.models.activity_domain import ActivityDomain
from dotenv import load_dotenv

load_dotenv()


def load_test_attendees() -> List[str]:
    raw_value = os.getenv("TEST_ATTENDEES", "")
    attendees = [item.strip() for item in raw_value.split(",") if item.strip()]
    if not attendees:
        raise ValueError("TEST_ATTENDEES is required. Provide a comma-separated list in .env.")
    return attendees


SCOPES = ["Calendars.Read.Shared", "Calendars.ReadWrite", "User.Read"]


def get_graph_client(use_device_code: bool = True) -> GraphServiceClient:
    """Initialize Graph client with delegated authentication.
    
    Args:
        use_device_code: If True, uses device code flow. If False, uses browser auth.
    """
    if use_device_code:
        credential = DeviceCodeCredential(
            client_id=os.getenv("AZURE_CLIENT_ID"),
            tenant_id=os.getenv("AZURE_TENANT_ID")
        )
    else:
        credential = InteractiveBrowserCredential(
            client_id=os.getenv("AZURE_CLIENT_ID"),
            tenant_id=os.getenv("AZURE_TENANT_ID")
        )
    
    return GraphServiceClient(credentials=credential, scopes=SCOPES)


def batch_attendees(attendees: List[str], batch_size: int = 5) -> List[List[str]]:
    """Split attendees into batches."""
    return [attendees[i:i + batch_size] for i in range(0, len(attendees), batch_size)]


def chunk_dates(start: datetime, end: datetime, days: int = 3) -> List[tuple]:
    """Split date range into chunks."""
    chunks = []
    current = start
    while current < end:
        chunk_end = min(current + timedelta(days=days), end)
        chunks.append((current, chunk_end))
        current = chunk_end
    return chunks


async def call_findmeetingtimes(client: GraphServiceClient, attendees: List[str], start: str, end: str, 
                          duration: int = 60, max_candidates: int = 50, save_debug: bool = False) -> Dict:
    """Call Graph API findMeetingTimes endpoint using SDK."""
    try:
        # Build attendee list
        attendee_list = []
        for email in attendees:
            attendee = AttendeeBase()
            attendee.type = AttendeeType.Required
            email_addr = EmailAddress()
            email_addr.address = email
            attendee.email_address = email_addr
            attendee_list.append(attendee)
        
        # Build time constraint
        time_constraint = TimeConstraint()
        time_constraint.activity_domain = ActivityDomain.Unrestricted
        
        time_slot = TimeSlot()
        start_time = DateTimeTimeZone()
        start_time.date_time = start
        start_time.time_zone = "UTC"
        time_slot.start = start_time
        
        end_time = DateTimeTimeZone()
        end_time.date_time = end
        end_time.time_zone = "UTC"
        time_slot.end = end_time
        
        time_constraint.time_slots = [time_slot]
        
        # Build request body
        request_body = FindMeetingTimesPostRequestBody()
        request_body.attendees = attendee_list
        request_body.time_constraint = time_constraint
        request_body.meeting_duration = f"PT{duration}M"
        request_body.max_candidates = max_candidates
        request_body.is_organizer_optional = True
        
        # Call API
        result = await client.me.find_meeting_times.post(request_body)

        # Store original API response for debugging (only when requested)
        if save_debug:
            # Note: SDK returns objects, so we extract all available properties
            raw_response = {
                "odata_context": getattr(result, 'odata_context', None),
                "odata_type": getattr(result, 'odata_type', None),
                "additional_data": result.additional_data if hasattr(result, 'additional_data') else {},
                "emptySuggestionsReason": result.empty_suggestions_reason,
                "meetingTimeSuggestions": [],
                "_debug_total_count": len(result.meeting_time_suggestions) if result.meeting_time_suggestions else 0
            }
            
            # Extract all suggestion data
            if result.meeting_time_suggestions:
                for item in result.meeting_time_suggestions:
                    suggestion = {
                        "odata_type": getattr(item, 'odata_type', None),
                        "confidence": item.confidence,
                        "organizerAvailability": item.organizer_availability,
                        "suggestionReason": item.suggestion_reason,
                        "additional_data": item.additional_data if hasattr(item, 'additional_data') else {},
                    }
                    
                    if item.meeting_time_slot:
                        suggestion["meetingTimeSlot"] = {
                            "start": {
                                "dateTime": item.meeting_time_slot.start.date_time,
                                "timeZone": item.meeting_time_slot.start.time_zone
                            } if item.meeting_time_slot.start else None,
                            "end": {
                                "dateTime": item.meeting_time_slot.end.date_time,
                                "timeZone": item.meeting_time_slot.end.time_zone
                            } if item.meeting_time_slot.end else None
                        }
                    
                    if item.attendee_availability:
                        suggestion["attendeeAvailability"] = [
                            {
                                "availability": avail.availability.value if avail.availability else None,
                                "attendee": {
                                    "emailAddress": {
                                        "address": avail.attendee.email_address.address
                                    } if avail.attendee.email_address else None
                                } if avail.attendee else None
                            } for avail in item.attendee_availability
                        ]
                    
                    if item.locations:
                        suggestion["locations"] = [
                            {
                                "displayName": loc.display_name if hasattr(loc, 'display_name') else None
                            } for loc in item.locations
                        ]
                    
                    raw_response["meetingTimeSuggestions"].append(suggestion)
            
            result_file_path = Path(__file__).parent / "raw_api_response.json"
            with open(result_file_path, "w", encoding="utf-8") as f:
                json.dump(raw_response, f, indent=2)
            print(f"  📄 Raw API response saved to: {result_file_path} ({len(result.meeting_time_suggestions or [])} suggestions)")
        
        return {
            "status": 200,
            "data": {
                "meetingTimeSuggestions": [
                    {
                        "confidence": item.confidence,
                        "organizerAvailability": item.organizer_availability,
                        "meetingTimeSlot": {
                            "start": {
                                "dateTime": item.meeting_time_slot.start.date_time,
                                "timeZone": item.meeting_time_slot.start.time_zone
                            },
                            "end": {
                                "dateTime": item.meeting_time_slot.end.date_time,
                                "timeZone": item.meeting_time_slot.end.time_zone
                            }
                        }
                    } for item in (result.meeting_time_suggestions or [])
                ],
                "emptySuggestionsReason": result.empty_suggestions_reason
            }
        }
    except Exception as e:
        return {"status": 500, "data": str(e)}


def generate_time_slots(start_date: datetime, num_days: int = 30, slot_minutes: int = 30) -> List[Dict]:
    """Generate test time slots covering all hours (unrestricted)."""
    slots = []
    current = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    
    for day in range(num_days):
        date = current + timedelta(days=day)
        # Include weekends too - unrestricted domain
        for start_hour in range(0, 24, 1):  # Every hour, all day
            start_dt = date.replace(hour=start_hour, minute=0)
            end_dt = start_dt + timedelta(minutes=slot_minutes)
            slots.append({
                "start": start_dt.isoformat(),
                "end": end_dt.isoformat()
            })
    return slots


def show_mitigation_math(num_attendees: int, batch_size: int = 5, days_per_chunk: int = 3):
    """Show mitigation strategy: how many API calls needed."""
    attendees = load_test_attendees()[:num_attendees]
    slots = generate_time_slots(datetime.now(), 30)
    
    batches = batch_attendees(attendees, batch_size)
    
    start = datetime.fromisoformat(slots[0]["start"])
    end = datetime.fromisoformat(slots[-1]["end"])
    chunks = chunk_dates(start, end, days_per_chunk)
    
    total_calls = len(batches) * len(chunks)
    
    print(f"{num_attendees} attendees → {len(batches)} batches × {len(chunks)} chunks = {total_calls} API calls")


async def test_calendar_access(client: GraphServiceClient, attendees: List[str]):
    """Test if we can access the authenticated user's calendar."""
    print("\n[TEST] Checking calendar access...")
    try:
        # Get current user info
        user = await client.me.get()
        print(f"  ✓ Authenticated as: {user.mail or user.user_principal_name}")
        
        # Try to access calendar
        calendar = await client.me.calendar.get()
        print(f"  ✓ Can access calendar")
        return {"ok": 1, "fail": 0}
    except Exception as e:
        print(f"  ✗ Error: {str(e)[:100]}")
        return {"ok": 0, "fail": 1}


async def reproduce_bad_gateway(client: GraphServiceClient):
    """Reproduce 502 Bad Gateway with actual Power Apps parameters."""
    print("\n[1] REPRODUCING 502 BAD GATEWAY\n")

    result = await call_findmeetingtimes(
        client, load_test_attendees(),
        datetime(2026, 1, 29).isoformat(), datetime(2026, 12, 30).isoformat(),
        duration=10, max_candidates=1600, save_debug=True
    )
    
    root_path = Path(__file__).parent
    result_path = root_path / "api_result.json"
    
    with open(result_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2)
    print(f"  ✓ Result saved to: {result_path}")
    print(f"  Result: Status {result['status']}")
    if result['status'] != 200:
        print(f"  ✓ Error reproduced: {str(result['data'])[:100]}...")
    else:
        print(f"  ⚠ Reproduce Failed (may need more extreme params to trigger 502)")
    return result


async def test_mitigation_effectiveness(client: GraphServiceClient):
    """Test mitigation strategy effectiveness."""
    print("\n[2] TESTING MITIGATION STRATEGY\n")
    print("  Call: attendees split into batches × date chunks")
    result = await run_mitigation(
        client, load_test_attendees(),
        datetime(2026, 1, 29), datetime(2026, 2, 28),
        batch_size=5, days_per_chunk=3, max_candidates=50
    )
    print(f"\n  Results:")
    print(f"    Total API calls: {result['total_calls']}")
    print(f"    Successful: {result['successful']}")
    print(f"    Failed: {result['failed']}")
    print(f"    Success rate: {result['success_rate']:.1%}")
    return result


async def run_mitigation(client: GraphServiceClient, attendees: List[str], start: datetime, end: datetime,
                   batch_size: int = 5, days_per_chunk: int = 3, duration: int = 60,
                   max_candidates: int = 50) -> Dict:
    """Execute full mitigation: batch + chunk + call findMeetingTimes."""
    batches = batch_attendees(attendees, batch_size)
    chunks = chunk_dates(start, end, days_per_chunk)
    
    results = []
    successful = 0
    failed = 0
    
    for batch in batches:
        for chunk_start, chunk_end in chunks:
            result = await call_findmeetingtimes(
                client, batch,
                chunk_start.isoformat(), chunk_end.isoformat(),
                duration, max_candidates
            )
            results.append(result)
            if result["status"] == 200:
                successful += 1
            else:
                failed += 1
    
    return {
        "total_calls": len(results),
        "successful": successful,
        "failed": failed,
        "success_rate": successful / len(results) if results else 0,
        "results": results
    }


async def main():
    start_time = time.time()
    print("\n=== FINDMEETINGTIMES 502 BAD GATEWAY MITIGATION POC ===")
    print(f"Start time: {datetime.now().strftime('%H:%M:%S')}")
    
    print("\n[0] STRATEGY CALCULATION")
    show_mitigation_math(25, batch_size=5, days_per_chunk=3)
    
    client = get_graph_client(use_device_code=True)

    print("\n[1] REPRODUCING 502 BAD GATEWAY")
    t1 = time.time()
    bad_result = await reproduce_bad_gateway(client)
    elapsed_1 = time.time() - t1
    print(f"  ⏱️  Elapsed: {elapsed_1:.2f}s")
    
    print("\n[2] TESTING MITIGATION STRATEGY")
    t2 = time.time()
    good_result = await test_mitigation_effectiveness(client)
    elapsed_2 = time.time() - t2
    print(f"  ⏱️  Elapsed: {elapsed_2:.2f}s")
        
    print("\n[3] COMPARISON")
    print(f"  Unmitigated: 1 call, Status={bad_result['status']} ({elapsed_1:.2f}s)")
    print(f"  Mitigated: {good_result['total_calls']} calls, {good_result['success_rate']:.0%} success ({elapsed_2:.2f}s)")
    
    if good_result['success_rate'] >= 0.8:
        print(f"\n Mitigation strategy works!")
        print(f"   Batching + chunking = {good_result['total_calls']} manageable API calls")
    else:
        print(f"\n Low success rate: {good_result['success_rate']:.0%}")
    
    elapsed_total = time.time() - start_time
    print(f"\n📊 Total execution time: {elapsed_total:.2f}s")
    print("=== END OF POC ===\n")


if __name__ == "__main__":
    import sys
    try:
        asyncio.run(main())
    except Exception as e:
        msg = str(e)
        print("\n" + "="*60)
        print("\nError occurred:")
        if "AADSTS7000218" in msg:
            print("   AADSTS7000218: App is confidential. Create/use a PUBLIC client app.")
            print("   - Enable 'Allow public client flows' in Azure AD app")
            print("   - Add redirect URI: http://localhost:8400")
            print("   - Do NOT require client_secret for browser auth")
        elif "AADSTS500113" in msg:
            print("   AADSTS500113: Redirect URI not registered.")
            print("   - Add redirect URI: http://localhost:8400")
        else:
            print(f"   {msg[:200]}")
        sys.exit(1)

