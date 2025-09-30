FILENAME = 'AccessProfilesTEAM.xlsx'



from access_profile_automation import OpenAccessProfilesXLSX

def placeholder(name, filter, dict):
    import time
    time.sleep(0.03)

def main():
    f = OpenAccessProfilesXLSX(file=FILENAME, debug=True)

    for key in f.access_profile_dict.keys():
        f.print_progress_bar(prefix='', suffix=key + "         ")
        placeholder(name=key, dict=f.to_dict()[key])  # Simulate work being done

    #print(f.to_json())
    #f.to_csv()

if __name__ == "__main__":
    main()