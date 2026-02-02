# Automated Excel Reporting System

This project automates manual Excel-based reporting using Python and SQL.

## Features
- Reads sales data from SQLite database
- Generates formatted Excel reports with charts
- Sends reports via email using Gmail SMTP
- Can be scheduled using Windows Task Scheduler

## Tech Stack
- Python
- SQLite (SQL)
- Pandas
- openpyxl
- Gmail SMTP

## How to Run
1. Install required libraries
2. Configure email settings in config.ini
3. Run `python generate_report.py`
4. Email is disabled by default (send_email = False).

## Use Case
Designed for internship-level automation of repetitive Excel reporting tasks.
