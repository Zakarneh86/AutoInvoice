import json
import os
from openai import OpenAI
import pymupdf
from pathlib import Path
import base64
import pandas as pd
from openpyxl import load_workbook
from copy import copy
from dateutil import parser