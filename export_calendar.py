import os
import datetime
import win32com.client
import pytz
import arrow
from icalendar import Calendar, Event, Timezone, TimezoneStandard, vDatetime

def export_calendar(file_path):
    print("Starting export process...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(9)  # 9 corresponds to the Calendar folder
        items = calendar.Items

        local_tz = pytz.timezone("America/Toronto")

        items.Sort("[Start]")
        items.IncludeRecurrences = True

        now = datetime.datetime.now(local_tz)
        future = now + datetime.timedelta(days=90)  # Set limit to 3 months in the future

        filtered_items = []
        for item in items:
            if now <= item.Start <= future:
                filtered_items.append(item)
            if item.Start > future:
                break

        print(f"Filtered items count: {len(filtered_items)}")

        # Create a new iCalendar
        cal = Calendar()
        cal.add("prodid", "-//Outlook Exporter//example.com//")
        cal.add("version", "2.0")

        for item in filtered_items:
            event = Event()
            event.add("summary", item.Subject)
            event.add("dtstart", item.Start.astimezone(pytz.utc))
            event.add("dtend", item.End.astimezone(pytz.utc))
            event.add("location", item.Location)
            event.add("uid", item.EntryID)
            cal.add_component(event)

        print("Writing iCalendar to file...")

        # Save the iCalendar to file
        with open(file_path, "wb") as f:
            f.write(cal.to_ical())

        print(f"Calendar exported successfully to {file_path}")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Close and release Outlook instance
        if 'outlook' in locals():
            outlook.Quit()
            del outlook

if __name__ == "__main__":
    file_path = os.path.join(os.getcwd(), "calendar.ics")
    export_calendar(file_path)
