import re

START_YEAR = 2012
END_YEAR = 2025

# Verbose regex for filenames with named groups
FILENAME_PATTERN = re.compile(
    r"""
    ^
    (?:
        (?P<country>[A-Z][A-Za-z]+(?:-[A-Z][A-Za-z]+)*)  # Match countries (e.g., Germany, South-Korea)
        |
        (?P<category>[A-Za-z0-9\-]+)                                     # Match categories (e.g., financial-return, non-life-insurance-1)
    )
    (?P<company>\d+)?               # Optional digit(s) for company number
    [_-]                            # Separator (underscore or hyphen)
    (?:(?!\d{4}(?:[-_]|$)).*?[_-])? # Optional intermediate text (e.g., quarterly-financial_), avoiding start year consumption
    (?P<start_year>\d{4})           # 4-digit start year
    (?:[-_](?P<end_year>\d{4}))?    # Optional hyphen/underscore and 4-digit end year
    (?:_(?P<suffix>[A-Za-z]+))?     # Optional suffix, e.g., "_A"
    (?:\.xlsm)?$                    # File extension
    """,
    re.VERBOSE
)
