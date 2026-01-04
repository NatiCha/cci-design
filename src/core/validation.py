"""
Event validation and conflict detection.
"""

from collections import defaultdict

from core.config import NON_PROJECT_NAMES, VALID_PHASE_CODES, VALID_TASK_CODES


def get_project_name(project_id: str) -> str:
    """Extract project name from 'Name: ID' format."""
    if ":" in project_id:
        return project_id.split(":")[0].strip().lower()
    return project_id.strip().lower()


def is_non_project(project_id: str) -> bool:
    """Check if this is a non-project (Office, Vacation, etc.)."""
    name = get_project_name(project_id)
    return name in NON_PROJECT_NAMES


def is_office_project(project_id: str) -> bool:
    """Check if this is the Office project."""
    name = get_project_name(project_id)
    return name == "office"


def validate_events(events: list[dict]) -> list[dict]:
    """
    Validate events and populate error_message field.

    Checks:
    1. Task and Phase are present
    2. Task/Phase codes are valid for project type
    3. Project name consistency (same name shouldn't have different IDs)
    """
    # Track project name -> IDs mapping for consistency check
    project_name_ids: dict[str, set[str]] = defaultdict(set)

    for event in events:
        project_id = event["project_id"]
        task = event["task"]
        phase = event["phase"]
        errors = []

        # Track project name -> ID mapping
        project_name = get_project_name(project_id)
        if project_name and project_id:
            project_name_ids[project_name].add(project_id)

        # Check 1: Task and Phase are present
        if not task:
            errors.append("Missing task code")
        if not phase:
            errors.append("Missing phase code")

        # Check 2: Valid codes for project type
        if task or phase:
            if is_non_project(project_id) and not is_office_project(project_id):
                # Non-projects (except Office) must use NA/NA
                if task and task != "NA":
                    errors.append(f"Non-project must use task 'NA', got '{task}'")
                if phase and phase != "NA":
                    errors.append(f"Non-project must use phase 'NA', got '{phase}'")
            elif is_office_project(project_id):
                # Office can use BD or NA for task
                if task and task not in {"BD", "NA"}:
                    errors.append(f"Office must use task 'BD' or 'NA', got '{task}'")
                if phase and phase not in {"NA", "PD", "SD"}:
                    # Allow some flexibility for Office phase
                    if phase not in VALID_PHASE_CODES:
                        errors.append(f"Invalid phase code '{phase}'")
            else:
                # Regular projects - validate codes
                if task and task not in VALID_TASK_CODES:
                    errors.append(f"Invalid task code '{task}'")
                if phase and phase not in VALID_PHASE_CODES:
                    errors.append(f"Invalid phase code '{phase}'")
                # Regular projects shouldn't use NA
                if task == "NA":
                    errors.append("Regular project cannot use task 'NA'")
                if phase == "NA":
                    errors.append("Regular project cannot use phase 'NA'")
                # BD is only for Office
                if task == "BD":
                    errors.append("Task 'BD' is only valid for Office project")

        event["error_message"] = "; ".join(errors) if errors else None

    # Check 3: Project name consistency
    for project_name, ids in project_name_ids.items():
        if len(ids) > 1:
            # Multiple IDs for same project name
            ids_str = ", ".join(sorted(ids))
            for event in events:
                if get_project_name(event["project_id"]) == project_name:
                    consistency_error = f"Project '{project_name}' has multiple IDs: {ids_str}"
                    if event["error_message"]:
                        event["error_message"] += "; " + consistency_error
                    else:
                        event["error_message"] = consistency_error

    return events
