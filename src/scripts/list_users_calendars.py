#!/usr/bin/env python3
"""
List all users and their calendars from MS365.

Usage:
    uv run python src/scripts/list_users_calendars.py
"""

import asyncio
import sys
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.graph_client import get_graph_client


async def main():
    """List all users and their calendars."""
    graph = get_graph_client()

    # Get all users
    print("Fetching users from MS365...\n")
    users_response = await graph.users.get()
    users = users_response.value if users_response.value else []

    print(f"Found {len(users)} users\n")
    print("=" * 80)

    for user in users:
        print(f"\nUser: {user.display_name}")
        print(f"  Email: {user.user_principal_name}")
        print(f"  ID: {user.id}")

        try:
            # Get calendars for this user
            calendars_response = await graph.users.by_user_id(user.id).calendars.get()
            calendars = calendars_response.value if calendars_response.value else []

            if calendars:
                print(f"  Calendars ({len(calendars)}):")
                for cal in calendars:
                    print(f"    - {cal.name}")
                    print(f"      ID: {cal.id}")
                    if cal.color:
                        print(f"      Color: {cal.color}")
            else:
                print("  Calendars: None")

        except Exception as e:
            print(f"  Error fetching calendars: {e}")

        print("-" * 80)

    print("\nDone!")


if __name__ == "__main__":
    asyncio.run(main())
