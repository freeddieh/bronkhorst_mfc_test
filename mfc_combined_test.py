import time
import multiprocessing
from mfc_logger_v1 import main_logger
from mfc_controller_v1 import main_controller
from airpy import BronkhorstMFC, find_bronkhorst_ports


def logger_process(bronkhorsts):
    print('logging')
    main_logger(bronkhorsts)


def controller_process(bronkhorsts, sleep_time):
    print('controlling')
    main_controller(bronkhorsts, sleep_time)


def main():
    sleep_time = 10  # 1 hour between instructions

    bh_ports = list(find_bronkhorst_ports().values())
    bronkhorsts = [BronkhorstMFC(bh_port) for bh_port in bh_ports]

    # Start logger in a separate process
    logger = multiprocessing.Process(target=logger_process, args=(bronkhorsts,))
    controller = multiprocessing.Process(target=controller_process, args=(bronkhorsts, sleep_time))

    logger.start()
    controller.start()

    try:
        while controller.is_alive():
            time.sleep(10)
    except KeyboardInterrupt:
        print("Interrupted by user. Terminating processes...")
        controller.terminate()

    # When controller finishes, terminate logger
    if logger.is_alive():
        print("Controller finished. Terminating logger...")
        logger.terminate()

    logger.join()
    controller.join()
    print("All processes have exited. Program terminated.")