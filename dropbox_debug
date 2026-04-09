import os
import dropbox

DROPBOX_APP_KEY = os.environ.get("DROPBOX_APP_KEY")
DROPBOX_APP_SECRET = os.environ.get("DROPBOX_APP_SECRET")
DROPBOX_REFRESH_TOKEN = os.environ.get("DROPBOX_REFRESH_TOKEN")
DROPBOX_TOKEN = os.environ.get("DROPBOX_TOKEN")

def init_dropbox():
    if DROPBOX_REFRESH_TOKEN:
        return dropbox.Dropbox(
            oauth2_refresh_token=DROPBOX_REFRESH_TOKEN,
            app_key=DROPBOX_APP_KEY,
            app_secret=DROPBOX_APP_SECRET,
        )
    elif DROPBOX_TOKEN:
        return dropbox.Dropbox(DROPBOX_TOKEN)
    else:
        raise Exception("Missing Dropbox credentials")

dbx = init_dropbox()

def debug_dropbox_paths():
    print("\n===== DROPBOX PATH DEBUG START =====\n")

    candidates = [
        "/YOE",
        "/YOE/חרבות ברזל 2023",
        "/YOE/חרבות ברזל 2023/20260228 - שאגת הארי",
        "/Template",
        "/Template/23-48",
        "/YOE/חרבות ברזל 2023/Template",
        "/YOE/חרבות ברזל 2023/Template/23-48",
        "/20260228 - שאגת הארי",
    ]

    for p in candidates:
        try:
            meta = dbx.files_get_metadata(p)
            print("FOUND PATH:", p)
            print("  TYPE:", type(meta).__name__)
            print("  ID:", getattr(meta, "id", ""))
            print("  PATH_DISPLAY:", getattr(meta, "path_display", ""))
            print("  PATH_LOWER:", getattr(meta, "path_lower", ""))

            sharing_info = getattr(meta, "sharing_info", None)
            if sharing_info:
                print("  PARENT_SHARED_FOLDER_ID:", getattr(sharing_info, "parent_shared_folder_id", ""))

            print("-" * 50)

        except Exception as e:
            print("NOT FOUND:", p, str(e))
            print("-" * 50)

    print("\n===== DROPBOX PATH DEBUG END =====\n")

if __name__ == "__main__":
    debug_dropbox_paths()
