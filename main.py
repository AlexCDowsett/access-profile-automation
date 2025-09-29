from access_profile_automation import OpenAccessProfilesXLSX

def placeholder(name, filter, dict):
    import time
    time.sleep(0.03)

def main():
    f = OpenAccessProfilesXLSX(file='AccessProfilesTEAM.xlsx', debug=True)

    for key in f.access_profile_dict.keys():
        f.print_progress_bar(prefix='', suffix=key + "         ")
        placeholder(name=key, filter=f.filter[key], dict=f.access_profile_dict[key])  # Simulate work being done

    #print(f.to_json())



if __name__ == "__main__":
    main()