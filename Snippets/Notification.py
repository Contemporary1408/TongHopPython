from winotify import Notification # MOST RECOMMEND!
import shutil
import os
home_directory = os.path.expanduser('~').replace("\\","/")
downloads_path = os.path.join(home_directory, 'Downloads').replace("\\","/")
src_path = r"\\10.118.29.7\BTMV-Data\1-ALL\05. 1611\asset\sua.ico"
ico_path = os.path.join(downloads_path, "sua.ico").replace("\\","/")
shutil.copyfile(src_path, ico_path) # Phải copy icon về ổ 1 folder local thì toast mới nhận, VD:Downloads
toast = Notification(app_id="windows app",
                     title="Winotify Test Toast",
                     msg="New Notification!",
                     icon=ico_path)

toast.show()

# #######################################################
from plyer import notification
# Define a function to show a Windows 10 notification
def show_notification(title, message):
    notification.notify(
        title=title,
        message=message,
        app_name='Python Notification',
        timeout=10
    )
# Call the function to show a notification
show_notification("Hello!", "This is a Windows 10 notification from Python.")

# ###################################################################################

from win10toast import ToastNotifier
# Create an object for ToastNotifier
toaster = ToastNotifier()
# Show a notification with custom icon
toaster.show_toast("Custom Notification",
                   "This notification has a custom icon!",
                   icon_path="custom_icon.ico",
                   duration=10)
