import sys
import time
import threading
from mfc_logger_v1 import main_logger
from mfc_controller_v1 import main_controller
from airpy import BronkhorstMFC, find_bronkhorst_ports

### SEMI WORKING BUILD ###
### Function order with threading is not working properly ###


def main_logger_event(event, bronkhorsts):
    """
    
    Main logger function that will wait for the event and log data.


    """
    while True:
        if event.is_set():  # If the event is set, proceed with logging
            main_logger(bronkhorsts)  # Call the main_logger function
            event.clear()  # Reset the event after the logger has completed its task
        #time.sleep(1)  # Sleep for 1 second before checking again (to avoid busy waiting)

def main_controller_event(event, bronkhorsts, sleep_time):
    """
    
    Main controller function that sends commands and has priority.
    
    """
    while True:
        main_controller(bronkhorsts, sleep_time)  # Send the command to the device
        event.set()  # Signal that the logger can proceed
        time.sleep(10)  # Wait for some time before sending the next command

def main():
    # Define the sleeptime between steps of the chosen program
    sleep_time = 3600

    # Find the Bronkhorst MFC ports
    bh_ports = list(find_bronkhorst_ports().values())  # This is your method to find the ports
    bronkhorsts = [BronkhorstMFC(bh_ports[0]), BronkhorstMFC(bh_ports[1])]  # Initialize the objects
    
    main_logger(bronkhorsts)
    main_controller(bronkhorsts, sleep_time)


if __name__ == '__main__':
    main()