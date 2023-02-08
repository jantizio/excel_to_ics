# Excel to .ics converter

Simple script that creates a calendar file (.ics) from an excel file formatted in a specific way.

The excel file needs to have one row with dates and can be set in the variable `days_index`, and other rows with data associated with the date in corresponding column. The latter is asked from user input

The `line_offset` is equal to two because the first row isn't counted and the numeration starts with 0, while in excel starts with one
