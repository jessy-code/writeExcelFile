writeExcelFileLine
======

Write a line in an Excel file. Create the file if it doesn't exist.

options:
____
    dest:
        description: Complete file path where the Excel file will be saved
        required: true
        type: str
    sheet:
        description: Sheet name target
        required: true
        type: str
    line_number:
        description: The line number target
        required: true
        type: int
    line:
        description: The line content
        required: true
        type: list of string

author:
____
    Jessy Martin (@jessy-code)