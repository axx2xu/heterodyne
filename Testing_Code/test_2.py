import threading
import time
import msvcrt  # Import the msvcrt module for Windows key press detection

flag = True
exit_event = threading.Event()

def exit_loop():
    global flag
    print('Press any key to stop the loop')
    while not exit_event.is_set():
        if msvcrt.kbhit():  # Check for key press
            msvcrt.getch()  # Clear the key press
            print('Stopping the loop...')
            flag = False
            break
        time.sleep(0.1)  # Check for key press every 100ms

# Start the thread to listen for a key press
n = threading.Thread(target=exit_loop)
n.start()

# Example main loop
num_steps = 10  # Example number of steps

try:
    for step in range(num_steps):
        if not flag:
            break
        # Your measurement and plotting logic here
        print(f"Step {step + 1}")
        time.sleep(1)  # Simulate some work with a delay

finally:
    # Signal the exit loop thread to stop
    exit_event.set()
    n.join()  # Wait for the thread to finish
    print("Loop stopped.")
