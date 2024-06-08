from tqdm import tqdm
import time
from JSONtest import jsonscrape

def dynamic_process():
    p_bar = tqdm(total=10)  # Initialize the progress bar with an arbitrary total

    start_time = time.time()
    while True:
        # Simulate some processing
        jsonscrape() # or any script here

        # Update the progress bar based on elapsed time
        elapsed_time = time.time() - start_time
        p_bar.n = elapsed_time * 1  # Update the progress bar value
        p_bar.update()  # Refresh the progress bar display

        # Break the loop after some condition is met (e.g., after 10 seconds)
        if elapsed_time >= 10:
            break

    p_bar.close()  # Close the progress bar when done


# Call the function
dynamic_process()
