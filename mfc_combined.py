import sys
import time
import threading
from mfc_logger_v1 import main_logger
from mfc_controller_v1 import main_controller
from airpy import BronkhorstMFC, find_bronkhorst_ports

### SEMI WORKING BUILD ###
### Function order with threading is not working properly ###


def main_logger_event(stop_event, bronkhorsts):
    while not stop_event.is_set():
        main_logger(bronkhorsts)
        time.sleep(1)  # Adjust if needed


def main_controller_event(stop_event, bronkhorsts, sleep_time):
    try:
        main_controller(bronkhorsts, sleep_time)
    finally:
        stop_event.set()  # Signal logger to stop when controller is done


def main():
    # Define the sleeptime between steps of the chosen program
    sleep_time = 10

    # Find the Bronkhorst MFC ports
    bh_ports = list(find_bronkhorst_ports().values())
    bronkhorsts = [BronkhorstMFC(bh_port) for bh_port in bh_ports]

    # Create coordination and shutdown events
    stop_event = threading.Event()

    logger_thread = threading.Thread(target=main_logger_event, args=(stop_event, bronkhorsts))
    controller_thread = threading.Thread(target=main_controller_event, args=(stop_event, bronkhorsts, sleep_time))

    logger_thread.start()
    controller_thread.start()

    try:
        while not stop_event.is_set():
            time.sleep(10)
    except KeyboardInterrupt:
        print("Shutting down gracefully...")
        stop_event.set()

    logger_thread.join()
    controller_thread.join()
    print("All threads have exited. Program terminated.")


if __name__ == '__main__':
    main()