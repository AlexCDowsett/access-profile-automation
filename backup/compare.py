"""
Comparison script to verify that V1 and V2 produce identical JSON outputs.
"""

import json
from access_profile_automation import OpenAccessProfilesXLSX
from access_profile_automation_v2 import OpenAccessProfilesXLSXV2


def compare_json_outputs():
    """Compare the JSON outputs of V1 and V2 implementations."""
    print("🔄 Loading V1 implementation...")
    v1 = OpenAccessProfilesXLSX()
    
    print("🔄 Loading V2 implementation...")
    v2 = OpenAccessProfilesXLSXV2()
    
    print("\n=== JSON COMPARISON ===")
    
    # Get JSON outputs
    v1_json = v1.to_json()
    v2_json = v2.to_json()
    
    # Parse JSON strings back to dictionaries for comparison
    try:
        v1_data = json.loads(v1_json)
        v2_data = json.loads(v2_json)
    except json.JSONDecodeError as e:
        print(f"❌ JSON parsing error: {e}")
        return False
    
    # Compare the parsed data
    if v1_data == v2_data:
        print("✅ JSON outputs are identical!")
        print(f"✅ Both contain {len(v1_data)} profiles")
        return True
    else:
        print("❌ JSON outputs are different!")
        
        # Find differences
        v1_keys = set(v1_data.keys())
        v2_keys = set(v2_data.keys())
        
        if v1_keys != v2_keys:
            print(f"❌ Profile keys differ:")
            print(f"   V1 only: {v1_keys - v2_keys}")
            print(f"   V2 only: {v2_keys - v1_keys}")
        
        # Check a sample profile for detailed differences
        common_keys = v1_keys & v2_keys
        if common_keys:
            sample_key = list(common_keys)[0]
            if v1_data[sample_key] != v2_data[sample_key]:
                print(f"❌ Sample profile '{sample_key}' differs:")
                print(f"   V1: {v1_data[sample_key]}")
                print(f"   V2: {v2_data[sample_key]}")
        
        return False


def compare_basic_stats():
    """Compare basic statistics between V1 and V2."""
    print("\n=== BASIC STATISTICS COMPARISON ===")
    
    v1 = OpenAccessProfilesXLSX()
    v2 = OpenAccessProfilesXLSXV2()
    
    # Profile counts
    v1_count = len(v1.access_profile_dict.keys())
    v2_count = len(v2.access_profile_dict.keys())
    print(f"Profile count - V1: {v1_count}, V2: {v2_count}")
    
    # Filter counts
    v1_filter_count = len(v1.filter.keys())
    v2_filter_count = len(v2.filter.keys())
    print(f"Filter count - V1: {v1_filter_count}, V2: {v2_filter_count}")
    
    # Check if profile names match
    v1_profiles = set(v1.access_profile_dict.keys())
    v2_profiles = set(v2.access_profile_dict.keys())
    profiles_match = v1_profiles == v2_profiles
    print(f"Profile names match: {'✅' if profiles_match else '❌'}")
    
    # Check if filter values match
    filters_match = True
    for profile in v1_profiles & v2_profiles:
        if v1.filter.get(profile) != v2.filter.get(profile):
            print(f"❌ Filter mismatch for {profile}: V1='{v1.filter.get(profile)}', V2='{v2.filter.get(profile)}'")
            filters_match = False
    
    print(f"Filter values match: {'✅' if filters_match else '❌'}")
    
    return profiles_match and filters_match


def main():
    """Run all comparisons."""
    print("🔍 COMPARING V1 AND V2 IMPLEMENTATIONS")
    print("=" * 50)
    
    # Basic statistics comparison
    basic_match = compare_basic_stats()
    
    # JSON output comparison
    json_match = compare_json_outputs()
    
    print("\n" + "=" * 50)
    print("🏁 COMPARISON RESULTS:")
    print(f"   Basic statistics: {'✅ PASS' if basic_match else '❌ FAIL'}")
    print(f"   JSON outputs: {'✅ PASS' if json_match else '❌ FAIL'}")
    
    if basic_match and json_match:
        print("\n🎉 ALL COMPARISONS PASSED!")
        print("   V1 and V2 implementations produce identical results.")
    else:
        print("\n⚠️  SOME COMPARISONS FAILED!")
        print("   V1 and V2 implementations produce different results.")
    
    return basic_match and json_match


if __name__ == "__main__":
    main()
