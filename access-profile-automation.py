import openpyxl
from collections import defaultdict
import time

import json

#dict['name']['UC']['User Groups'] = value

def placeholder(name, key, dict):
    time.sleep(0.1)
    print(key)

def nested_dict():
    """Creates a 3-level nested defaultdict."""
    return defaultdict(lambda: defaultdict(dict))


def print_progress_bar(current, total, bar_length=40, prefix='Progress', suffix='', show=False):
    if not show:
        return
    percent = float(current) / total if total else 0
    filled_length = int(bar_length * percent)
    bar = '\u2588' * filled_length + '-' * (bar_length - filled_length)
    print(f'\r{prefix} |{bar}| {percent*100:6.2f}% ({current}/{total}) {suffix}', end='')
    if current >= total:
        print()  # Newline on complete 

def main():
    f = OpenAccessProfilesXLSX()
    print(f.to_json())
    total = len(f.access_profile_dict.keys())
    i = 1
    for key in f.access_profile_dict.keys():
        print_progress_bar(i, total, prefix='Processing Access Profiles:', suffix=key, show=True)
        placeholder(key, f.access_profile_dict[key]['filter'], f.access_profile_dict[key])  # Simulate work being done
        i += 1






class OpenAccessProfilesXLSX():
    def __init__(self):
        # Load the workbook and select the active sheet
        workbook = openpyxl.load_workbook('AccessProfilesTEAM.xlsx', data_only=True)
        sheet = workbook.active
                
        self.access_profile_dict = nested_dict()
        
        current_row = 0
        for row in sheet.iter_rows(min_row=0, max_row=1, values_only=True):
            current_row += 1
            if any(cell in ['Conductor', 'storm Contact', 'UC', 'DataManagement', 'Dial ', 'Flow'] for cell in row):
                categories = [cell for cell in row if cell is not None]
                #print(categories)
                #print(current_row)
                break

        for row in sheet.iter_rows(min_row=current_row, values_only=True):
            current_row += 1
            if any(cell in ['Actions', 'Announcements', 'Menus'] for cell in row):
                headings = [cell for cell in row if cell is not None]
                #print(headings)
                #print(current_row)
                break

        for row in sheet.iter_rows(min_row=current_row, values_only=True):
            current_row += 1
            row_no_none = [cell for cell in row if cell is not None]
            #print(row_no_none)
    
            #print(current_row)
            break

        for row in sheet.iter_rows(min_row=current_row, values_only=True):
            current_row += 1
            row_no_none = [cell for cell in row if cell is not None]
            name = row_no_none[0]
            filter = row_no_none[1]
            row_no_none = row_no_none[2:]

            cat_index = 0
            for i in range(len(headings)):
                if headings[i] in ['Agent Groups', 'Call Barring Profiles', 'Queries', 'Pacing Profiles', 'Flow Services']:
                    cat_index += 1
                self.access_profile_dict[name]['filter'] = filter
                if row_no_none[2*i].lower() not in ['none', 'not in use']:
                    self.access_profile_dict[name][categories[cat_index]][headings[i]] = [row_no_none[2*i], row_no_none[2*i+1]]

    def to_json(self):
        return json.dumps(self.access_profile_dict, indent=4)
            

    


if __name__ == "__main__":
    main()