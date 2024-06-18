from winotify import Notification  # MOST RECOMMEND!

toast = Notification(app_id="windows app",
                     title="Winotify Test Toast",
                     msg="New Notification!",
                     icon=r"c:\path\to\icon.png")

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
