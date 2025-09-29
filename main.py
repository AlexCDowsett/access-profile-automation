from access-profile-automation import OpenAccessProfilesXLSX



def placeholder(name, key, dict):
    import time
    time.sleep(0.03)

def main():
    f = OpenAccessProfilesXLSX()
    print(f.to_json())
    total = len(f.access_profile_dict.keys())
    i = 1
    for key in f.access_profile_dict.keys():
        self.print_progress_bar(i, total, prefix='Processing Access Profiles:', suffix=key, show=True)
        placeholder(key, f.access_profile_dict[key]['filter'], f.access_profile_dict[key])  # Simulate work being done
        i += 1



if __name__ == "__main__":
    main()