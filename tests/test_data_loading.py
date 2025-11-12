"""
Tests for CSV data loading and validation.

Simple tests that verify data loading functions work correctly.
"""

import pandas as pd

from scheduler.data_access.staff_loader import _coerce_numeric, _normalize_columns, _parse_roles


def test_normalize_basic_columns():
    """Test that basic column names are normalized correctly."""
    df = pd.DataFrame(columns=["Name", "ROLES", "target_hours"])
    normalized = _normalize_columns(df)
    
    assert "name" in normalized
    assert "roles" in normalized
    assert "target_hours" in normalized


def test_parse_semicolon_separated_roles():
    """Test parsing roles separated by semicolons."""
    result = _parse_roles("front_desk;marketing;events")
    assert result == ["front_desk", "marketing", "events"]


def test_parse_comma_separated_roles():
    """Test parsing roles separated by commas."""
    result = _parse_roles("front_desk,marketing,events")
    assert result == ["front_desk", "marketing", "events"]


def test_parse_roles_with_whitespace():
    """Test that whitespace is stripped from role names."""
    result = _parse_roles("  front_desk ; marketing  ; events  ")
    assert result == ["front_desk", "marketing", "events"]


def test_parse_empty_roles():
    """Test that empty or NaN roles return empty list."""
    assert _parse_roles(None) == []
    assert _parse_roles("") == []


def test_coerce_integer():
    """Test coercing integer values."""
    result = _coerce_numeric(10, "test_col", "test_record")
    assert result == 10.0


def test_coerce_float():
    """Test coercing float values."""
    result = _coerce_numeric(10.5, "test_col", "test_record")
    assert result == 10.5


def test_coerce_string_number():
    """Test coercing string representations of numbers."""
    result = _coerce_numeric("10.5", "test_col", "test_record")
    assert result == 10.5


def test_coerce_invalid_raises_error():
    """Test that invalid values raise a descriptive error."""
    try:
        _coerce_numeric("not_a_number", "hours", "John")
        assert False, "Should have raised ValueError"
    except ValueError as e:
        assert "Invalid numeric value" in str(e)
