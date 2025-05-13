"""
Utilities Module
==============

This module contains common utility functions used throughout the project.
"""

import pandas as pd
from datetime import datetime
from pathlib import Path
import json
from typing import Union, Dict, Any, Optional

def load_excel(file_path: Union[str, Path], **kwargs) -> Optional[pd.DataFrame]:
    """
    Excel 파일을 DataFrame으로 로드합니다.
    
    Args:
        file_path (Union[str, Path]): Excel 파일 경로
        **kwargs: pd.read_excel()에 전달할 추가 인자
        
    Returns:
        Optional[pd.DataFrame]: 로드된 데이터 또는 None (파일이 없거나 형식이 잘못된 경우)
        
    Raises:
        FileNotFoundError: 파일이 존재하지 않을 경우
        ValueError: 파일 형식이 잘못되었을 경우
    """
    try:
        return pd.read_excel(file_path, engine='openpyxl', **kwargs)
    except (FileNotFoundError, ValueError):
        return None

def save_excel(df: pd.DataFrame, file_path: Union[str, Path], **kwargs) -> Path:
    """
    DataFrame을 Excel 파일로 저장합니다.
    
    Args:
        df (pd.DataFrame): 저장할 데이터
        file_path (Union[str, Path]): 저장할 파일 경로
        **kwargs: df.to_excel()에 전달할 추가 인자
        
    Returns:
        Path: 저장된 파일의 경로
        
    Raises:
        ValueError: 저장 중 오류가 발생한 경우
    """
    try:
        file_path = Path(file_path)
        file_path.parent.mkdir(parents=True, exist_ok=True)
        # index 매개변수가 kwargs에 있으면 제거 (이미 기본값으로 False 설정)
        kwargs.pop('index', None)
        df.to_excel(file_path, index=False, engine='openpyxl', **kwargs)
        return file_path
    except Exception as e:
        raise ValueError(f"Excel 파일 저장 중 오류 발생: {str(e)}")

def load_json(file_path: Union[str, Path]) -> Optional[Dict[str, Any]]:
    """
    JSON 파일을 로드합니다.
    
    Args:
        file_path (Union[str, Path]): JSON 파일 경로
        
    Returns:
        Optional[Dict[str, Any]]: 로드된 데이터 또는 None (파일이 없거나 JSON 형식이 잘못된 경우)
        
    Raises:
        FileNotFoundError: 파일이 존재하지 않을 경우
        json.JSONDecodeError: JSON 형식이 잘못된 경우
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None

def save_json(data: Dict[str, Any], file_path: Union[str, Path], **kwargs) -> Path:
    """
    데이터를 JSON 파일로 저장합니다.
    
    Args:
        data (Dict[str, Any]): 저장할 데이터
        file_path (Union[str, Path]): 저장할 파일 경로
        **kwargs: json.dump()에 전달할 추가 인자
        
    Returns:
        Path: 저장된 파일의 경로
        
    Raises:
        ValueError: 저장 중 오류가 발생한 경우
    """
    try:
        file_path = Path(file_path)
        file_path.parent.mkdir(parents=True, exist_ok=True)
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2, **kwargs)
        return file_path
    except Exception as e:
        raise ValueError(f"JSON 파일 저장 중 오류 발생: {str(e)}")

def format_date(date: Union[str, datetime], format_str: str = '%Y-%m-%d') -> Optional[str]:
    """
    날짜를 지정된 형식의 문자열로 변환합니다.
    
    Args:
        date (Union[str, datetime]): 변환할 날짜
        format_str (str): 출력 형식
        
    Returns:
        Optional[str]: 변환된 날짜 문자열 또는 None (입력이 None인 경우)
        
    Raises:
        ValueError: 날짜 형식이 잘못된 경우
    """
    if date is None:
        return None
    try:
        if isinstance(date, str):
            date = pd.to_datetime(date)
        return date.strftime(format_str)
    except Exception as e:
        return None

def parse_date(date_str: str, format_str: Optional[str] = None) -> Optional[datetime]:
    """
    날짜 문자열을 datetime 객체로 변환합니다.
    
    Args:
        date_str (str): 변환할 날짜 문자열
        format_str (Optional[str]): 입력 형식 (None인 경우 자동 감지)
        
    Returns:
        Optional[datetime]: 변환된 datetime 객체 또는 None (변환 실패 시)
        
    Raises:
        ValueError: 날짜 형식이 잘못된 경우
    """
    try:
        if format_str:
            return datetime.strptime(date_str, format_str)
        return pd.to_datetime(date_str)
    except Exception as e:
        return None

def safe_divide(numerator: float, denominator: float, default_value: Any = None) -> Any:
    """
    안전한 나눗셈을 수행합니다.

    Args:
        numerator (float): 분자
        denominator (float): 분모
        default_value (Any, optional): 분모가 0일 때 반환할 값. 기본값은 None.

    Returns:
        Any: 나눗셈 결과 또는 default_value
    """
    try:
        return numerator / denominator
    except ZeroDivisionError:
        return default_value 