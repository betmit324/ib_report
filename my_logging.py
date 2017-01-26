__author__ = 'mneil'

import sys
import time
import logging

debug_it = 0

def printToLog(logString):
    """ Print to log file (stderr)
    Prints the logString to stderr, prepends date and time
    """
    print(time.strftime("%Y%m%d-%H:%M:%S") + ": " + logString, file=sys.stderr)


def print_to_log(log_string):
    """ Print to log file (stderr)
    Prints the logString to stderr, prepends date and time
    """
    print(time.strftime("%H:%M:%S") + ": " + log_string, file=sys.stderr)
    logging.info(time.strftime("%H:%M:%S") + ": " + log_string)
    if debug_it:
        with open("temp.log", mode='a', encoding='utf-8') as templog:
            print(time.strftime("%H:%M:%S") + ": " + log_string, file=templog)

