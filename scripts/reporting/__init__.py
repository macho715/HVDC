"""
Reporting Module
==============

This module handles report generation and dashboard creation.
"""

from .excel import HVDCExcelReporter
from .dashboard import HVDCDashboardGenerator

__all__ = ['HVDCExcelReporter', 'HVDCDashboardGenerator'] 