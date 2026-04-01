# Shopee Seller Dashboard

A desktop application for Shopee sellers to track monthly sales,
manage returns, calculate platform fees, and generate WhatsApp summaries.

Built with Python and CustomTkinter.

## Features

- **Overview** — import Shopee reports and view monthly revenue, fees,
  net income, top products, and a daily sales chart
- **Tax Calculator** — calculate Shopee commissions and fees by price
  tier, based on March 2026 pricing rules
- **Returns** — monitor refund totals, request status, top return reasons,
  and automated performance tips
- **WhatsApp Summary** — generate a formatted message with the month's
  highlights, ready to copy and send

## Requirements

- Python 3.9+
- Dependencies:

\`\`\`
pip install customtkinter pandas openpyxl matplotlib
\`\`\`

## Running locally

\`\`\`bash
python painel_shopee.py
\`\`\`

## Building the .exe

\`\`\`bash
pip install pyinstaller
pyinstaller --onefile --windowed painel_shopee.py
\`\`\`

The executable will be generated in the `dist/` folder.

## How to use

1. Open the app and go to the **Overview** tab
2. Import your monthly Shopee reports:
   - **Orders report** — from Seller Center > My Orders > Export
   - **Returns report** — from Seller Center > Returns/Refunds > Export
   - **Income report** — from Seller Center > My Income > Export
3. Navigate between tabs to view metrics, calculate fees, and generate
   your WhatsApp summary

## Notes

- Fee tiers are based on Shopee Brazil's pricing as of March 2026
- Always verify current rates directly on Shopee Seller Center
