from __future__ import annotations

from datetime import datetime
from typing import Any, Dict, List, Optional

from app.core.config import DEFAULT_PROJECT_ROWS_TO_FILL, PROJECT_SELECTION_MODE
from app.core.constants import (
    MASTER_PROJECTS_KEY,
    PROJECT_SELECTION_BALANCED,
    PROJECT_SELECTION_COMPLETED_FIRST,
    PROJECT_SELECTION_HIGHEST_VALUE,
    PROJECT_SELECTION_NEWEST,
    PROJECT_SELECTION_ONGOING_FIRST,
)


def normalize_project_status(value) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def _safe_float(value: Any) -> float:
    if value is None:
        return 0.0

    text = str(value).strip()
    if not text:
        return 0.0

    cleaned = (
        text.replace(",", "")
        .replace("₹", "")
        .replace("rs.", "")
        .replace("rs", "")
        .replace("inr", "")
        .replace("cr", "")
        .replace("crore", "")
        .replace("lakh", "")
        .strip()
    )

    try:
        return float(cleaned)
    except Exception:
        return 0.0


def _parse_date(value: Any) -> Optional[datetime]:
    if value is None:
        return None

    if isinstance(value, datetime):
        return value

    text = str(value).strip()
    if not text:
        return None

    date_formats = [
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d.%m.%Y",
        "%Y/%m/%d",
        "%d %b %Y",
        "%d %B %Y",
        "%b %Y",
        "%B %Y",
    ]

    for fmt in date_formats:
        try:
            return datetime.strptime(text, fmt)
        except Exception:
            continue

    return None


def _project_sort_date(project: Dict[str, Any]) -> Optional[datetime]:
    for key in ("end_date", "start_date"):
        parsed = _parse_date(project.get(key))
        if parsed:
            return parsed
    return None


def _project_value(project: Dict[str, Any]) -> float:
    return _safe_float(project.get("value"))


def _status_bucket(status: str) -> str:
    status = normalize_project_status(status)

    if status in {"ongoing", "in progress", "running", "active"}:
        return "ongoing"

    if status in {"completed", "complete", "done", "closed", "finished"}:
        return "completed"

    return "other"


def _get_projects_from_master(master_data: dict) -> List[Dict[str, Any]]:
    if not isinstance(master_data, dict):
        return []

    projects = master_data.get(MASTER_PROJECTS_KEY, [])
    if isinstance(projects, list):
        return projects

    # Tolerant fallback for future loaders
    for alt_key in ("projects", "project_master"):
        alt_value = master_data.get(alt_key)
        if isinstance(alt_value, list):
            return alt_value

    return []


def _filter_by_mode(projects: List[Dict[str, Any]], mode: str) -> List[Dict[str, Any]]:
    if mode == "all":
        return projects[:]

    filtered: List[Dict[str, Any]] = []

    for project in projects:
        status = _status_bucket(project.get("status"))

        if mode == "ongoing":
            if status != "ongoing":
                continue

        elif mode == "completed":
            if status != "completed":
                continue

        filtered.append(project)

    return filtered


def _sort_projects(projects: List[Dict[str, Any]], selection_mode: str) -> List[Dict[str, Any]]:
    if selection_mode == PROJECT_SELECTION_HIGHEST_VALUE:
        return sorted(
            projects,
            key=lambda p: (_project_value(p), _project_sort_date(p) or datetime.min),
            reverse=True,
        )

    if selection_mode == PROJECT_SELECTION_NEWEST:
        return sorted(
            projects,
            key=lambda p: (_project_sort_date(p) or datetime.min, _project_value(p)),
            reverse=True,
        )

    if selection_mode == PROJECT_SELECTION_COMPLETED_FIRST:
        return sorted(
            projects,
            key=lambda p: (
                1 if _status_bucket(p.get("status")) == "completed" else 0,
                _project_value(p),
                _project_sort_date(p) or datetime.min,
            ),
            reverse=True,
        )

    if selection_mode == PROJECT_SELECTION_ONGOING_FIRST:
        return sorted(
            projects,
            key=lambda p: (
                1 if _status_bucket(p.get("status")) == "ongoing" else 0,
                _project_value(p),
                _project_sort_date(p) or datetime.min,
            ),
            reverse=True,
        )

    # Default balanced ordering:
    # 1. ongoing/completed relevance
    # 2. higher value
    # 3. newer date
    return sorted(
        projects,
        key=lambda p: (
            1 if _status_bucket(p.get("status")) in {"ongoing", "completed"} else 0,
            _project_value(p),
            _project_sort_date(p) or datetime.min,
        ),
        reverse=True,
    )


def _balanced_pick(projects: List[Dict[str, Any]], limit: int) -> List[Dict[str, Any]]:
    """
    Try to mix ongoing and completed projects when possible.
    """
    ongoing = []
    completed = []
    other = []

    for project in projects:
        bucket = _status_bucket(project.get("status"))
        if bucket == "ongoing":
            ongoing.append(project)
        elif bucket == "completed":
            completed.append(project)
        else:
            other.append(project)

    ongoing = _sort_projects(ongoing, PROJECT_SELECTION_HIGHEST_VALUE)
    completed = _sort_projects(completed, PROJECT_SELECTION_HIGHEST_VALUE)
    other = _sort_projects(other, PROJECT_SELECTION_HIGHEST_VALUE)

    selected: List[Dict[str, Any]] = []

    # Alternate ongoing/completed first
    while len(selected) < limit and (ongoing or completed):
        if ongoing and len(selected) < limit:
            selected.append(ongoing.pop(0))
        if completed and len(selected) < limit:
            selected.append(completed.pop(0))

    # Then fill remaining from leftover pools by value
    leftovers = ongoing + completed + other
    leftovers = _sort_projects(leftovers, PROJECT_SELECTION_HIGHEST_VALUE)

    for project in leftovers:
        if len(selected) >= limit:
            break
        selected.append(project)

    return selected[:limit]


def select_projects(master_data: dict, mode: str = "all", limit: int = 5):
    """
    Deterministic project selector.

    Supported mode values:
    - all
    - ongoing
    - completed

    Ordering is controlled by config PROJECT_SELECTION_MODE.
    """
    projects = _get_projects_from_master(master_data)

    if not isinstance(projects, list) or not projects:
        return []

    if limit is None or limit <= 0:
        limit = DEFAULT_PROJECT_ROWS_TO_FILL

    filtered = _filter_by_mode(projects, mode)

    if not filtered:
        # fallback to all if mode filter produced nothing
        filtered = projects[:]

    if PROJECT_SELECTION_MODE == PROJECT_SELECTION_BALANCED:
        return _balanced_pick(filtered, limit)

    ranked = _sort_projects(filtered, PROJECT_SELECTION_MODE)
    return ranked[:limit]