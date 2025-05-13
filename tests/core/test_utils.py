import pytest
from pathlib import Path
import pandas as pd
from datetime import datetime
from scripts.core import utils

def test_safe_divide_normal():
    """정상적인 나눗셈 테스트"""
    assert utils.safe_divide(10, 2) == 5.0

def test_safe_divide_by_zero():
    """0으로 나누는 경우 테스트"""
    assert utils.safe_divide(10, 0) is None

def test_safe_divide_by_zero_with_custom_return():
    """0으로 나누는 경우 사용자 정의 값 반환 테스트"""
    assert utils.safe_divide(10, 0, default_value=0) == 0
    assert utils.safe_divide(10, 0, default_value="Error") == "Error"

def test_format_date():
    """날짜 포맷팅 함수 테스트"""
    dt_obj = datetime(2024, 5, 10, 12, 30, 0)
    assert utils.format_date(dt_obj, "%Y-%m-%d") == "2024-05-10"
    assert utils.format_date(dt_obj, "%Y/%m/%d %H:%M") == "2024/05/10 12:30"

def test_format_date_none():
    """None 입력 시 format_date 함수 테스트"""
    assert utils.format_date(None, "%Y-%m-%d") is None

def test_parse_date():
    """날짜 문자열 파싱 함수 테스트"""
    date_str = "2024-05-10"
    expected_dt_obj = datetime(2024, 5, 10)
    assert utils.parse_date(date_str, "%Y-%m-%d") == expected_dt_obj

def test_parse_date_invalid_format():
    """잘못된 형식의 날짜 문자열 파싱 시도 테스트"""
    assert utils.parse_date("10/05/2024", "%Y-%m-%d") is None

def test_save_and_load_json(tmp_path: Path):
    """JSON 저장 및 로드 함수 테스트"""
    test_data = {"key1": "value1", "numbers": [1, 2, 3], "nested": {"a": True}}
    file_path = tmp_path / "test_data.json"
    utils.save_json(test_data, file_path)
    assert file_path.exists()
    loaded_data = utils.load_json(file_path)
    assert loaded_data == test_data

def test_load_json_nonexistent_file(tmp_path: Path):
    """존재하지 않는 JSON 파일 로드 시도 테스트"""
    non_existent_file = tmp_path / "non_existent.json"
    assert utils.load_json(non_existent_file) is None

def test_save_and_load_excel(tmp_path: Path):
    """Excel 저장 및 로드 함수 테스트"""
    test_df = pd.DataFrame({"colA": [1, 2, 3], "colB": ["x", "y", "z"]})
    file_path = tmp_path / "test_data.xlsx"
    utils.save_excel(test_df, file_path, index=False)
    assert file_path.exists()
    loaded_df = utils.load_excel(file_path)
    pd.testing.assert_frame_equal(loaded_df, test_df)

def test_load_excel_nonexistent_file(tmp_path: Path):
    """존재하지 않는 Excel 파일 로드 시도 테스트"""
    non_existent_file = tmp_path / "non_existent.xlsx"
    assert utils.load_excel(non_existent_file) is None 