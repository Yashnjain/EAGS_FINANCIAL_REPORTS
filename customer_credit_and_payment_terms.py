import os
import re
import sys
import time
import glob
import logging
from mail_alert import send_mail
import pandas as pd
import xlwings as xw
import xlwings.constants as win32c
from datetime import date, datetime


loc1 = r""


if __name__ == "__main__":
    print
              