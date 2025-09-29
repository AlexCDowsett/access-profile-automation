"""
Main script using the V2 implementation of Access Profile Automation.

This script demonstrates the usage of the improved V2 implementation
with robust Excel parsing and better error handling.
"""

import time
from access_profile_automation_v2 import OpenAccessProfilesXLSXV2


def placeholder(name: str, key: str, profile_dict: dict) -> None:
    """
    Placeholder function to simulate work being done.
    
    Args:
        name: Profile name
        key: Filter key
        profile_dict: Profile data dictionary
    """
    time.sleep(0.03)


def main():
    """Main function using V2 implementation."""
    try:
        print("🚀 Starting Access Profile Automation V2")
        print("=" * 50)
        
        # Initialize the V2 parser
        print("📊 Loading Excel file...")
        parser = OpenAccessProfilesXLSXV2()
        
        # Get basic statistics
        total_profiles = len(parser.access_profile_dict.keys())
        total_filters = len(parser.filter.keys())
        
        print(f"✅ Successfully loaded {total_profiles} access profiles")
        print(f"✅ Found {total_filters} filter entries")
        
        # Show sample data
        if total_profiles > 0:
            sample_profile = list(parser.access_profile_dict.keys())[0]
            sample_data = parser.access_profile_dict[sample_profile]
            sample_filter = parser.filter[sample_profile]
            
            print(f"\n📋 Sample Profile: {sample_profile}")
            print(f"   Filter: {sample_filter}")
            print(f"   Categories: {list(sample_data.keys())}")
            
            # Show first category data
            if sample_data:
                first_category = list(sample_data.keys())[0]
                first_heading = list(sample_data[first_category].keys())[0]
                first_values = sample_data[first_category][first_heading]
                print(f"   Sample data: {first_category}[{first_heading}] = {first_values}")
        
        print(f"\n🔄 Processing {total_profiles} profiles...")
        
        # Process each profile with progress bar
        for i, key in enumerate(parser.access_profile_dict.keys(), 1):
            parser.print_progress_bar(
                current=i, 
                total=total_profiles, 
                prefix='', 
                suffix=key + "        "
            )
            placeholder(key, parser.filter[key], parser.access_profile_dict[key])
        
        print()  # Newline after progress bar
        print("✅ Processing complete!")
        
        # Optional: Show JSON output (uncomment if needed)
        # print("\n📄 JSON Output:")
        # print(parser.to_json())
        
        print("\n🎉 V2 Implementation completed successfully!")
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        raise


if __name__ == "__main__":
    main()
