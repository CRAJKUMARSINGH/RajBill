import os
import pytest
import pandas as pd
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from streamlit_app import process_bill, number_to_words

def test_number_to_words():
    assert number_to_words(123456) == "One Lakh Twenty Three Thousand Four Hundred And Fifty Six"
    assert number_to_words(1000000) == "Ten Lakh"
    assert number_to_words(10000000) == "One Crore"

def test_process_bill_no_extra_items():
    # Create sample data
    ws_wo = pd.DataFrame({
        0: [1, 2],  # Serial No
        1: ["Item 1", "Item 2"],  # Description
        2: ["Unit1", "Unit2"],  # Unit
        4: [100, 200],  # Rate
        6: ["", ""]  # Remark
    })

    ws_bq = pd.DataFrame({
        3: [2, 3]  # Quantity
    })

    ws_extra = pd.DataFrame()  # Empty extra items

    premium_percent = 10
    premium_type = "fixed"
    amount_paid_last_bill = 0
    is_first_bill = True
    user_inputs = {
        "start_date": "2024-01-01",
        "completion_date": "2024-12-31",
        "work_name": "Test Work",
        "bill_serial": "1",
        "agreement_no": "AG123",
        "work_order_ref": "WO123",
        "work_order_amount": "100000",
        "bill_type": "Regular",
        "bill_number": "B123",
        "last_bill_reference": ""
    }

    result = process_bill(ws_wo, ws_bq, ws_extra, premium_percent, premium_type, 
                         amount_paid_last_bill, is_first_bill, user_inputs)

    assert result["totals"]["grand_total"] == 800  # 2*100 + 3*200
    assert result["totals"]["premium"]["amount"] == 80  # 10% of 800
    assert result["totals"]["payable"] == 880  # 800 + 80

def test_process_bill_with_extra_items():
    # Create sample data with extra items
    ws_wo = pd.DataFrame({
        0: [1, 2],  # Serial No
        1: ["Item 1", "Item 2"],  # Description
        2: ["Unit1", "Unit2"],  # Unit
        4: [100, 200],  # Rate
        6: ["", ""]  # Remark
    })

    ws_bq = pd.DataFrame({
        3: [2, 3]  # Quantity
    })

    ws_extra = pd.DataFrame({
        0: [1],  # Serial No
        1: ["Extra Item"],  # Description
        2: ["Unit"],  # Unit
        3: [5],  # Quantity
        4: [150],  # Rate
        6: ["Extra"]  # Remark
    })

    premium_percent = 10
    premium_type = "fixed"
    amount_paid_last_bill = 0
    is_first_bill = True
    user_inputs = {
        "start_date": "2024-01-01",
        "completion_date": "2024-12-31",
        "work_name": "Test Work",
        "bill_serial": "1",
        "agreement_no": "AG123",
        "work_order_ref": "WO123",
        "work_order_amount": "100000",
        "bill_type": "Regular",
        "bill_number": "B123",
        "last_bill_reference": ""
    }

    result = process_bill(ws_wo, ws_bq, ws_extra, premium_percent, premium_type, 
                         amount_paid_last_bill, is_first_bill, user_inputs)

    assert result["totals"]["grand_total"] == 800  # 2*100 + 3*200
    assert result["totals"]["premium"]["amount"] == 80  # 10% of 800
    assert result["totals"]["payable"] == 880  # 800 + 80
    assert len(result["items"]) > 2  # Should have extra items section

def test_process_bill_with_previous_payment():
    # Test with previous payment
    ws_wo = pd.DataFrame({
        0: [1, 2],  # Serial No
        1: ["Item 1", "Item 2"],  # Description
        2: ["Unit1", "Unit2"],  # Unit
        4: [100, 200],  # Rate
        6: ["", ""]  # Remark
    })

    ws_bq = pd.DataFrame({
        3: [2, 3]  # Quantity
    })

    ws_extra = pd.DataFrame()  # Empty extra items

    premium_percent = 10
    premium_type = "fixed"
    amount_paid_last_bill = 300
    is_first_bill = False
    user_inputs = {
        "start_date": "2024-01-01",
        "completion_date": "2024-12-31",
        "work_name": "Test Work",
        "bill_serial": "2",
        "agreement_no": "AG123",
        "work_order_ref": "WO123",
        "work_order_amount": "100000",
        "bill_type": "Regular",
        "bill_number": "B123",
        "last_bill_reference": "B122"
    }

    result = process_bill(ws_wo, ws_bq, ws_extra, premium_percent, premium_type, 
                         amount_paid_last_bill, is_first_bill, user_inputs)

    assert result["totals"]["grand_total"] == 800  # 2*100 + 3*200
    assert result["totals"]["premium"]["amount"] == 80  # 10% of 800
    assert result["totals"]["payable"] == 580  # 880 - 300

def test_process_bill_invalid_inputs():
    # Test with invalid DataFrame inputs
    with pytest.raises(ValueError, match="Invalid input data"):
        process_bill(None, None, None, 10, "fixed", 0, True, {})

    # Test with invalid premium percent
    with pytest.raises(ValueError, match="Premium percent must be a non-negative number"):
        process_bill(pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), -10, "fixed", 0, True, {})

    # Test with invalid premium type
    with pytest.raises(ValueError, match="Premium type must be either 'fixed' or 'percentage'"):
        process_bill(pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), 10, "invalid", 0, True, {})

    # Test with missing required fields
    with pytest.raises(ValueError, match="Missing required field"):
        process_bill(pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), 10, "fixed", 0, True, {})

def test_process_bill_edge_cases():
    # Test with empty DataFrames
    ws_wo = pd.DataFrame()
    ws_bq = pd.DataFrame()
    ws_extra = pd.DataFrame()
    
    with pytest.raises(ValueError, match="Error calculating total amount"):
        process_bill(ws_wo, ws_bq, ws_extra, 10, "fixed", 0, True, {
            "start_date": "2024-01-01",
            "completion_date": "2024-12-31",
            "work_name": "Test Work",
            "bill_serial": "1",
            "agreement_no": "AG123",
            "work_order_ref": "WO123",
            "work_order_amount": "100000"
        })

    # Test with very large numbers
    ws_wo = pd.DataFrame({
        0: [1],
        1: ["Item 1"],
        2: ["Unit1"],
        4: [1e9],
        6: [""]
    })
    ws_bq = pd.DataFrame({
        3: [1e9]
    })
    
    result = process_bill(ws_wo, ws_bq, pd.DataFrame(), 10, "fixed", 0, True, {
        "start_date": "2024-01-01",
        "completion_date": "2024-12-31",
        "work_name": "Test Work",
        "bill_serial": "1",
        "agreement_no": "AG123",
        "work_order_ref": "WO123",
        "work_order_amount": "1000000000000"
    })
    
    assert result["totals"]["grand_total"] == 1e18

def test_number_to_words_edge_cases():
    # Test with zero
    assert number_to_words(0) == "Zero"
    
    # Test with very large numbers
    assert number_to_words(100000000000) == "One Hundred Crore"
    
    # Test with negative numbers
    with pytest.raises(ValueError):
        number_to_words(-100)

if __name__ == "__main__":
    pytest.main(["-v", "-s", "test_bill_generator.py"])
