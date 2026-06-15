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

##################################
# OpenAI Interface and  Prompts  #
##################################

## Initialize OpenAI client
def client (APIkey):
    try:
        client = OpenAI(APIkey = APIkey)
        error = False
        error_text = "Client initialized successfully."
        return client, error, error_text
    except Exception as e:
        client = None
        error = True
        error_text = f"Failed to initialize OpenAI client: {e}"
        return client, error, error_text

## Loading PO Json Schema
with open("po_schema.json", "r") as f:
    po_schema = json.load(f)

## Loading Timesheet Json Schema
with open("timesheet_schema.json", "r") as f:
    timesheet_schema = json.load(f)

## PO Reader Prompts
PO_ROLE_ALIGNMENT_RULE = """
Extract PO details exactly from the provided PO document.

Rules:

1. Return only valid JSON matching the schema.
2. Do not invent, estimate, infer, or assume missing values.
3. If a value is not explicitly provided and cannot be derived from an explicit PO formula, return null.
4. If information is not found in the PO, return null.

Working Hours:
5. Extract working-hour information exactly as stated.
6. Create one working_hours item per location found in the PO.
7. For onshore_or_offshore return only:
   - onshore
   - offshore
   - not_specified

Rates:
8. Extract daily and hourly rates exactly as stated.
9. Do not calculate derived rates unless the PO explicitly defines the calculation method.
10. If multiple roles exist, create one rate record per role and location.
11. Preserve role names exactly as written in the PO.
12. If no role is specified return "NOT_SPECIFIED".

Derived Rates:
13. If only a base rate is provided and the PO explicitly defines a multiplier or formula, calculate the derived rate.
14. If the multiplier is not explicitly defined, return null.
15. Set rate_source to:
    - explicit
    - calculated
    - null
16. If a derived rate is calculated, populate calculation_note with the formula used.

PO Information:
17. Determine invoicing_type as:
    - daily_rate
    - hourly_rate
    - mixed
    - not_clear

18. Extract currency exactly as written.
19. Return dates exactly as written.
20. Return monetary values as numbers without currency symbols.
"""
PO_SYSTEM_PROMPT = f"""
You are an expert Purchase Order (PO) analysis and data extraction engine.

Your task is to extract structured information from Purchase Orders and return only data that is explicitly stated in the document or can be mathematically derived from stated rules.

{PO_ROLE_ALIGNMENT_RULE}

Final Rules:
1. Return data strictly according to the provided schema.
2. Do not invent, assume, estimate, or infer information that is not present in the PO.
3. When conflicting values are found, prioritize the most specific commercial section, schedule of rates, pricing table, or compensation schedule.
4. Ignore signatures, approvals, logos, headers, footers, and administrative text unless they contain commercial information.
5. If multiple amendments, revisions, change orders, or addendums exist, use the latest revision unless explicitly instructed otherwise.
6. Return only the final structured output. Do not include explanations, commentary, assumptions, confidence scores, markdown, or additional text.
"""

## Timesheet Reader Prompts
TIMESHEET_SYSTEM_PROMPT = f"""
You are an expert timesheet analysis and data extraction engine.
Your task is to extract structured information from timesheets and return only data that is explicitly stated in the document.
Rules:
1. Return only JSON matching the schema.
2. Extract one entry per populated row in the timesheet table.
3. Do not calculate normal, overtime, or hot hours.
4. Do not interpret PO rules.
5. Extract dates exactly from the table.
6. Extract From and To times exactly as written.
7. Extract values from columns:
   A = travel_hours
   B = weekday_friday_hours
   C = saturday_hours
8. If "HOURS ON SITE" does'nt have "FROM" and "TO" return the cell value in "Hours_on_site".
9. If "HOURS ON SITE" has values in "FROM" and "TO" return null in "Hours_on_site".
10. If "HOURS ON SITE" and "TRAV" and "EKD/FRI" and "SAT" are all null return null in all cell even if the date has value.
11. If a cell is blank return null.
12. If a cell is blank return null.
13. Ignore signatures and stamps.
14. Do not infer missing rows.
15. Extract engineer name, PO number, client and plant from the header.
16. The engineer name is typically located near the "FOR EMERSON:" signature block.
17. The engineer name may appear above, below, or beside the Emerson signature.
18. Do not use names found in the "FOR CLIENT:" section."""

##################################
#       PO Reader Functions      #
##################################
## PO Details Extraction using OpenAI
def generate_po_details(po_text: str):
    response = client.responses.create(
        model="gpt-5.5",
        input=[
            {
                "role": "system",
                "content": PO_SYSTEM_PROMPT
            },
            {
                "role": "user",
                "content": f"""Customer PO details:
                {po_text}
                """
            }
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": po_schema["name"],
                "schema": po_schema["schema"],
                "strict": True
            }
        }
    )
    return json.loads(response.output_text)

## PO Text Extraction from PDF
def get_po_data(po_w_pl, po_wo_pl, pl, haspl):
  doc_text = ''
  if haspl:
    po = pymupdf.open(po_w_pl)
    for page in po:
      text = page.get_text()
      doc_text = doc_text + ' '+text
  else:
    po = pymupdf.open(po_wo_pl)
    pl = pymupdf.open(pl)
    for page in po:
      text = page.get_text()
      doc_text = doc_text + ' '+text
    for page in pl:
      text = page.get_text()
      doc_text = doc_text + ' '+text

  po_json = generate_po_details(doc_text)
  return po_json

##################################
#   Timesheet Reader Functions   #
##################################
## Timesheet Conversion to base64
def file_to_base64(file, dpi=300):
  if file.endswith('.pdf'):
    doc = pymupdf.open(file)
    zoom = dpi / 72
    matrix = pymupdf.Matrix(zoom, zoom)
    images = []

    for page in doc:
        pix = page.get_pixmap(
            matrix=matrix,
            alpha=False)
        image_bytes = pix.tobytes("png")
        image_b64 = base64.b64encode(image_bytes).decode("utf-8")
        images.append(image_b64)
    doc.close()
  else:
    with open(file, "rb") as image_file:
      image_bytes = image_file.read()
      image_b64 = base64.b64encode(image_bytes).decode("utf-8")
      images = [image_b64]
  return images

## Timesheet Details Extraction using OpenAI
def generate_ts_details(timesheet):
  user_content = [{"type": "input_text", "text": "Extract the timesheet data."}]
  for image in timesheet:
    user_content.append({"type": "input_image", "image_url": f"data:image/png;base64,{image}"})

  response = client.responses.create(
      model="gpt-5.5",
      input=[
          {
              "role": "system",
              "content": TIMESHEET_SYSTEM_PROMPT
          },
          {
              "role": "user",
              "content": user_content
          }],
      text={
          "format": {
              "type": "json_schema",
              "name": timesheet_schema["name"],
              "schema": timesheet_schema["schema"],
              "strict": True
          }
      }
  )

  return json.loads(response.output_text)

## Timesheets Detail Extraction (This Needs to be Modified)
def get_timesheets_data(path):
  ts_details = {}
  for file in os.listdir(path):
    print(path+file)
    timesheet = file_to_base64(path+file)
    ts_json = generate_ts_details(timesheet)

    pro_info = {"project_name": ts_json['client'], "order_number": ts_json['client_order_number']}
    eng_name = ts_json['engineer_name']

    if 'pro_info' not in ts_details:
      ts_details['pro_info'] = pro_info
    if eng_name not in ts_details:
      ts_details[eng_name]= {'ts1': ts_json['entries']}
    else:
      ts_details[eng_name]['ts_'+str(len(ts_details[eng_name].keys())+1)] = ts_json['entries']
  return ts_details

##################################
#      Time Classification       #
##################################
ENGINEER_BLOCKS = [
    {"name_cell": "B4", "cols": {"normal": "B", "ot": "C", "fri": "D", "sat": "E"}},
    {"name_cell": "F4", "cols": {"normal": "F", "ot": "G", "fri": "H", "sat": "I"}},
    {"name_cell": "J4", "cols": {"normal": "J", "ot": "K", "fri": "L", "sat": "M"}},
    {"name_cell": "N4", "cols": {"normal": "N", "ot": "O", "fri": "P", "sat": "Q"}},
]

## Supporting Functions for Time Classification
def to_num(value):
    if value in [None, "", "null"]:
        return 0
    if isinstance(value, (int, float)):
        return value
    value = str(value).strip()
    if ":" in value:
        return int(value.split(":")[0])
    return float(value)

def duration_hours(from_time, to_time):
    if not from_time or not to_time:
        return 0

    start = to_num(from_time)
    end = to_num(to_time)

    if end < start:
        end += 24

    return end - start

def clean_date(date_text):
    return parser.parse(date_text, dayfirst=True).date()

## Time Classification Dunction
def classify_entry(entry, minimum_normal_hours):
    day_name = entry["day_name"]

    hours_on_site = to_num(entry.get("hours_on_site"))
    travel_hours = to_num(entry.get("travel_hours"))
    from_time = to_num(entry.get("from_time"))
    to_time = to_num(entry.get("to_time"))
    friday_hours = to_num(entry.get("friday_hours"))
    saturday_hours = to_num(entry.get("saturday_hours"))

    has_time_entry = any([
        hours_on_site > 0,
        travel_hours > 0,
        from_time > 0,
        to_time > 0,
        friday_hours > 0,
        saturday_hours > 0
    ])

    if not has_time_entry:
        return None

    normal_hours = 0
    overtime_hours = 0

    if day_name in ["SUN", "MON", "TUE", "WED", "THUR"]:
        if hours_on_site > 0:
            total_hours = hours_on_site + travel_hours
        else:
            total_hours = duration_hours(
                entry.get("from_time"),
                entry.get("to_time")
            ) + travel_hours

        if total_hours <= minimum_normal_hours:
            normal_hours = minimum_normal_hours
            overtime_hours = 0
        else:
            normal_hours = minimum_normal_hours
            overtime_hours = total_hours - minimum_normal_hours

    elif day_name == "FRI":
        friday_hours = to_num(entry.get("friday_hours"))

        if friday_hours == 0:
            friday_hours = hours_on_site or duration_hours(
                entry.get("from_time"),
                entry.get("to_time")
            )

    elif day_name == "SAT":
        saturday_hours = to_num(entry.get("saturday_hours"))

        if saturday_hours == 0:
            saturday_hours = hours_on_site or duration_hours(
                entry.get("from_time"),
                entry.get("to_time")
            )

    return {
        "day_name": day_name,
        "date": clean_date(entry["date"]),
        "normal": normal_hours,
        "ot": overtime_hours,
        "fri": friday_hours,
        "sat": saturday_hours
    }

# Calculation Excel Sheet Generator (This Needs to be Modified)
def fill_calculation_excel(
    output_path,
    ts_details,
    po_hourly_rates,
    po_working_hours,
    role,
    location
):
    template_path = "calculation_template.xlsx"

    po_number = int(ts_details["pro_info"]["order_number"])

    rate_row = po_hourly_rates[(po_hourly_rates["po_number"] == po_number) & (po_hourly_rates["role_name"] == role) & (po_hourly_rates["onshore_or_offshore"] == location)].iloc[0]
    hours_row = po_working_hours[(po_working_hours["po_number"] == po_number) & (po_working_hours["onshore_or_offshore"] == location)].iloc[0]

    normal_rate = rate_row["normal_week_day"]
    ot_rate = rate_row["overtime_week_day"]
    friday_rate = rate_row["friday"]
    saturday_rate = rate_row["saturday"]

    minimum_normal_hours = hours_row["normal_week_day"]

    wb = load_workbook(template_path)
    ws = wb.active

    ws["B2"] = ts_details["pro_info"]["project_name"]
    ws["B3"] = po_number

    all_dates = set()
    engineer_data = {}

    for engineer_name, engineer_timesheets in list(ts_details.items())[1:]:
        engineer_data[engineer_name] = {}

        for ts_name, entries in engineer_timesheets.items():
            for entry in entries:
                classified = classify_entry(entry, minimum_normal_hours)
                if classified is None:
                    continue
                date = classified["date"]

                all_dates.add(date)

                if date not in engineer_data[engineer_name]:
                    engineer_data[engineer_name][date] = {
                        "normal": 0,
                        "ot": 0,
                        "fri": 0,
                        "sat": 0
                    }

                engineer_data[engineer_name][date]["normal"] += classified["normal"]
                engineer_data[engineer_name][date]["ot"] += classified["ot"]
                engineer_data[engineer_name][date]["fri"] += classified["fri"]
                engineer_data[engineer_name][date]["sat"] += classified["sat"]

    sorted_dates = sorted(all_dates)

    start_row = 6

    for i, date in enumerate(sorted_dates):
        row = start_row + i
        ws[f"A{row}"] = date
        ws[f"A{row}"].number_format = "dd-mmm-yyyy"

    for eng_index, (engineer_name, data_by_date) in enumerate(engineer_data.items()):
        if eng_index >= len(ENGINEER_BLOCKS):
            raise ValueError("Template supports maximum 4 engineers only.")

        block = ENGINEER_BLOCKS[eng_index]
        ws[block["name_cell"]] = engineer_name

        for i, date in enumerate(sorted_dates):
            row = start_row + i
            values = data_by_date.get(date, {"normal": 0, "ot": 0, "fri": 0, "sat": 0})

            ws[f'{block["cols"]["normal"]}{row}'] = values["normal"]
            ws[f'{block["cols"]["ot"]}{row}'] = values["ot"]
            ws[f'{block["cols"]["fri"]}{row}'] = values["fri"]
            ws[f'{block["cols"]["sat"]}{row}'] = values["sat"]

        ws[f'{block["cols"]["normal"]}38'] = normal_rate
        ws[f'{block["cols"]["ot"]}38'] = ot_rate
        ws[f'{block["cols"]["fri"]}38'] = friday_rate
        ws[f'{block["cols"]["sat"]}38'] = saturday_rate

    wb.save(output_path)