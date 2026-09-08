"""Helpers for locating sheets in the inventory workbook."""


def find_sheet(sheetnames, target):
    """Return the sheet name matching target, ignoring case and surrounding spaces.

    Returns None when no sheet matches.
    """
    wanted = target.strip().lower()
    for name in sheetnames:
        if name.strip().lower() == wanted:
            return name
    return None
