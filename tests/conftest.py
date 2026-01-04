"""
Pytest configuration and shared fixtures.
"""

import sys
from pathlib import Path

import pytest

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))


@pytest.fixture
def sample_event():
    """Sample event dictionary for testing."""
    return {
        "project_id": "Test Project: 12345",
        "employee_id": "ABC",
        "start_timestamp": "2025-11-01T09:00:00Z",
        "end_timestamp": "2025-11-01T17:00:00Z",
        "event_date": None,
        "hours": 8.0,
        "task": "PM",
        "phase": "SD",
        "wid": "Test work item",
        "error_message": None,
    }


@pytest.fixture
def sample_events(sample_event):
    """List of sample events for testing."""
    return [
        sample_event,
        {
            **sample_event,
            "project_id": "Another Project: 67890",
            "task": "D-D",
            "phase": "DD",
        },
    ]
