# Google Contacts Birthday Sync

A desktop app to sync calendar birthdays from `.ics` files (or CSV) to Google Contacts, with a responsive UI for matching, creating, and updating contacts.

## Features

- **Import .ics files:** Add calendar events to the app's database (CSV).
- **Fuzzy matching:** Suggests best contact matches for each calendar entry using rapidfuzz.
- **Manual search:** Search and select contacts with autocomplete.
- **Create new contacts:** Instantly create a contact for a calendar entry.
- **Update birthdays:** Sync selected birthdays to Google Contacts (month and day only).
- **Remove entries:** Quickly remove unwanted calendar entries.
- **Detect duplicates:** Highlights repeated events (same title, different dates).
- **Progress bar & error tab:** See update progress and any errors.
- **Processed tab:** View entries already synced or removed.

## Usage

1. Install dependencies:
    ```
    pip install -r requirements.txt
    ```
2. Place your Google API `credentials.json` in the project folder.
3. Run the app:
    ```
    python google_contacts_birthday_sync.py
    ```
4. Use "Import .ics file" to add calendar events.
5. Match, create, update, or remove entries as needed.
6. Use the "Update contacts" button to sync selected entries in the background.

## Notes

- The app uses `calendar.csv` as its internal database.
- Only month and day are synced for birthdays (year is omitted).
- Requires Google API credentials for contacts access.
