# Graph find meeting times API Performance Testing & Mitigation Strategies

**Problem**: Many attendees & Heavy schedule data → 502 Bad Gateway

**Solution**: Batch attendees + chunk date ranges

## Setup

Microsoft Graph supports two access scenarios, delegated access and app-only access. In delegated access, the app calls Microsoft Graph on behalf of a signed-in user. In application-only access, the app calls Microsoft Graph with its own identity, without a signed-in user.

In the scenario of manipulating other users' schedules, you must use application-only access. In delegated access, you can only manipulate your schedule.

This demo script uses application-only access for creating events and cleaning up events. Delegated access is used for `find_meeting_times`.

### 1. Register Azure AD Application

1. Go to [Azure Portal](https://portal.azure.com) → Entra ID (Azure Active Directory) → App registrations
2. New registration:
   - Name: `FindMeetingTimes POC`
   - Supported account types: Single tenant or Multitenant
3. Copy **Application (client) ID** and **Directory (tenant) ID**

### 2. Add API Permissions

1. API permissions → Add permission → Microsoft Graph → Delegated
2. Add permissions (per [MS documentation](https://learn.microsoft.com/en-us/graph/api/user-findmeetingtimes)):
   - `Calendars.Read.Shared` - **Required** for findMeetingTimes ([docs](https://learn.microsoft.com/en-us/graph/api/user-findmeetingtimes))
   - `Calendars.ReadWrite` - **Required** for test data functions ([create](https://learn.microsoft.com/en-us/graph/api/user-post-events)/[delete](https://learn.microsoft.com/en-us/graph/api/event-delete) events)
   - `User.Read` - Basic user profile
3. **User consent via browser**: With interactive browser auth, users can consent to these permissions themselves during sign-in.
4. API permissions → Add permission → Microsoft Graph → Application
5. Add permissions and consent by administrator.
6. Manage > Authentication > Settings > Allow public client flows > Enabled

### 3. Update Code with Credentials

Rename `.env.template` to `.env`, and replace placeholders:

```ini
AZURE_TENANT_ID=your-tenant-id-here
AZURE_CLIENT_ID=your-client-id-here
AZURE_CLIENT_SECRET=your-client-secret-here
TEST_ATTENDEES="user1@contoso.com, user2@contoso.com, user3@contoso.com"
```
### 4. Install Dependencies

```bash
uv sync
```

### Usage

1. Schedule data creation for testing: `create_events.py`
2. Find meeting times API calls with mitigation strategies: `find_meetings.py`
3. Purge test data: `cleanup_events.py`

### Response sample

```json
{
  "status": 200,
  "data": {
    "meetingTimeSuggestions": [
      {
        "confidence": 87.25,
        "organizerAvailability": "busy",
        "meetingTimeSlot": {
          "start": {
            "dateTime": "2026-01-29T00:30:00.0000000",
            "timeZone": "UTC"
          },
          "end": {
            "dateTime": "2026-01-29T00:40:00.0000000",
            "timeZone": "UTC"
          }
        }
      }
    ],
    "emptySuggestionsReason": ""
  }
}
```

## References

- [user: findMeetingTimes API](https://learn.microsoft.com/en-us/graph/api/user-findmeetingtimes)
- [Microsoft Graph Python SDK](https://learn.microsoft.com/en-us/graph/sdks/sdks-overview)
- [Azure Identity](https://learn.microsoft.com/en-us/python/api/overview/azure/identity-readme)
