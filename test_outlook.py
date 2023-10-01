import win32com.client
import datetime
import pytz

def main():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(9)  # 9 corresponds to the Calendar folder
        items = calendar.Items

        local_tz = pytz.timezone("America/Toronto")

        items.Sort("[Start]")
        items.IncludeRecurrences = True

        now = datetime.datetime.now(local_tz)
        future = now + datetime.timedelta(days=90)  # Set a limit to 3 months in the future

        filtered_items = []
        for item in items:
            if now <= item.Start <= future:
                filtered_items.append(item)
            if item.Start > future:
                break

        print(f"Number of items within the specified date range: {len(filtered_items)}")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Close and release Outlook instance
        if 'outlook' in locals():
            outlook.Quit()
            del outlook

if __name__ == "__main__":
    main()
