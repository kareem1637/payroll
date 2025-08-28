#!/usr/bin/env python3
"""
Test script to verify the date range parameter functionality
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import payroll
import re

def test_date_range_validation():
    """Test the date range validation logic"""
    print("Testing date range validation...")

    # Test valid date ranges
    valid_ranges = [
        "07.15.24-07.21.24",
        "01.01.25-01.07.25",
        "12.31.24-01.06.25"
    ]

    for date_range in valid_ranges:
        if re.match(r'^\d{2}\.\d{2}\.\d{2}-\d{2}\.\d{2}\.\d{2}$', date_range):
            print(f"✓ Valid: {date_range}")
        else:
            print(f"✗ Invalid: {date_range}")

    # Test invalid date ranges
    invalid_ranges = [
        "07.15.24-07.21",  # Missing year digits
        "07/15/24-07/21/24",  # Wrong format
        "071524-072124",  # No dots
        "07.15.24",  # Missing end date
        ""  # Empty
    ]

    for date_range in invalid_ranges:
        if not re.match(r'^\d{2}\.\d{2}\.\d{2}-\d{2}\.\d{2}\.\d{2}$', date_range):
            print(f"✓ Correctly rejected: {date_range}")
        else:
            print(f"✗ Incorrectly accepted: {date_range}")

def test_process_files_signature():
    """Test that process_files accepts the new date_range_param"""
    print("\nTesting process_files function signature...")

    # Check function signature
    import inspect
    sig = inspect.signature(payroll.process_files)
    params = list(sig.parameters.keys())

    expected_params = ['payroll_path', 'capture_path', 'payroll_sheet', 'output_filename', 'date_range_param']

    if params == expected_params:
        print(f"✓ Function signature correct: {params}")
    else:
        print(f"✗ Function signature incorrect. Expected: {expected_params}, Got: {params}")

def test_get_weekly_counts():
    """Test the get_weekly_counts function with date range"""
    print("\nTesting get_weekly_counts function...")

    import pandas as pd

    # Create a simple test dataframe
    test_df = pd.DataFrame({
        'Provider': ['Dr. Smith', 'Dr. Jones'],
        'DOS': ['07.16.24', '07.17.24'],
        'CPT Codes': ['99213', '99214']
    })

    date_range = "07.15.24-07.21.24"

    try:
        result = payroll.get_weekly_counts(test_df, date_range)
        if isinstance(result, pd.DataFrame):
            print(f"✓ get_weekly_counts returned DataFrame with shape: {result.shape}")
        else:
            print(f"? get_weekly_counts returned: {type(result)}")
    except Exception as e:
        print(f"✗ get_weekly_counts failed: {e}")

if __name__ == "__main__":
    print("Running date range parameter tests...\n")
    test_date_range_validation()
    test_process_files_signature()
    test_get_weekly_counts()
    print("\nTest completed!")
