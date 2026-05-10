# Selenium Price Checker

A Python automation project created to practice Selenium, Excel file handling, and price comparison logic.

This project checks product prices from a website, compares the current prices with previous prices stored in an Excel file, and saves the detected price changes into a new Excel file.

## Features

- Read previous product prices from Excel
- Use Selenium for browser automation
- Check current product prices
- Compare old and new prices
- Detect price increases and decreases
- Save price changes into an Excel file
- Generate a simple text report

## Files

- script.py: Main Python automation script
- previous_prices.xlsx: Excel file containing previous product prices
- price_changes.xlsx: Excel file containing detected price changes
- report.txt: Short report about the automation result

## Technologies Used

- Python
- Selenium
- Excel
- openpyxl
- GitHub

## Concepts Practiced

- Web automation
- Selenium browser control
- Excel file reading and writing
- Price comparison logic
- Basic reporting
- Python scripting

## How It Works

The script reads old product prices from the previous_prices.xlsx file. Then it checks the current prices using Selenium. After that, it compares the previous and current prices and writes the changes into price_changes.xlsx.

A simple report is also created in report.txt.

## What I Learned

Through this project, I practiced automating browser actions with Selenium and working with Excel files in Python. I also learned how automation can be used for business data processing tasks such as price monitoring and reporting.
