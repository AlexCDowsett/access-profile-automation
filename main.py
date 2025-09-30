FILENAME = 'B.X_CONDUCTOR_AZ Italy_AccessProfiles_V1.1_18092025 (1).xlsx'
FILENAME = 'B.X_CONDUCTOR_AZ Bene_AccessProfiles_V1_30092025.xlsx'
FILENAME = 'AccessProfilesTEAM.xlsx'
SHEETNAME = None

from access_profile_automation import OpenAccessProfilesXLSX

def placeholder(name, dict):
    import time
    time.sleep(0.03)

def main():
    f = OpenAccessProfilesXLSX(file=FILENAME, sheetname=SHEETNAME, debug=True)
    f.to_json(True)
    f.to_csv()

    for name, filter_dict in f.to_dict().items():
        f.print_progress_bar(prefix='', suffix=name + "         ")
        placeholder(name=name, dict=filter_dict)  # Simulate work being done

    #print(f.to_json())
    #f.to_csv()

if __name__ == "__main__":
    main()