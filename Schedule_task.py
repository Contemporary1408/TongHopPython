import schedule
import time

def my_task():
    print("Running my task...")

# Schedule the task to run every 5 seconds
schedule.every(5).seconds.do(my_task)
#schedule.every().day.at("17:00").do(my_task)
#schedule.every().monday.at("18:00").do(my_task)
#schedule.every().wednesday.at("18:00").do(my_task)

# Run the scheduled tasks for 1 minute
end_time = time.time() + 60
while time.time() < end_time:
    schedule.run_pending()
    time.sleep(1)

# Print a completion message
print("Scheduled tasks completed.")
