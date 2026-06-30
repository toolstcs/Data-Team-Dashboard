# Lead Database Dashboard

Live web dashboard powered by Streamlit. Reads your CSV/Excel files on every page load, so edits reflect instantly on refresh.

---

## Quick Start

```bash
pip install -r requirements.txt
streamlit run app.py
```

Browser opens at **http://localhost:8501**

---

## File Placement

Put all data files inside the `data/` folder:

```
lead_dashboard/
  app.py                          <-- Run this
  requirements.txt
  data/
    hubspot_contacts.csv          <-- CSV 1: Main HubSpot export
    tcs_email_marketing.csv       <-- CSV 2: TCS email openers/non-openers
    binaryworks_email_marketing.csv  <-- CSV 3: BinaryWorks email openers/non-openers
    contribution.xlsx             <-- CSV 4: Individual contribution (3 sheets)
```

### CSV 1: hubspot_contacts.csv

Your main HubSpot contacts export. Required columns:

| Column | Purpose |
|--------|---------|
| TAG | Brand assignment (TCS, BinaryWorks, ConversionBox, Manufacturing, etc.) |
| Email | Lead count (each email = 1 lead) |
| Website URL | Unique company count |
| e-Commerce Technologies | TCS tech breakdown (Shopify, Magento, etc.) |
| Drupal Partners (CMS) | BinaryWorks breakdown (Drupal 7-11, WordPress, etc.) |
| ConversionBox Competitors | ConversionBox breakdown (show all values) |

**TAG values and where they go:**
- "TCS" --> TCS brand tab
- "BinaryWorks" --> BinaryWorks brand tab
- "ConversionBox" --> ConversionBox brand tab
- "Manufacturing" --> Others tab (card)
- "ConversionBox 200+ Products" --> Others tab (card)
- "Higher Education" --> Others tab (card)
- "Finance" --> Others tab (card)

### CSV 2: tcs_email_marketing.csv

| Column | Example |
|--------|---------|
| Email | john@company.com |
| Status | TCS Opener / TCS Non-Opener |

### CSV 3: binaryworks_email_marketing.csv

| Column | Example |
|--------|---------|
| Email | jane@company.com |
| Status | BinaryWorks Opener / BinaryWorks Non-Opener |

### CSV 4: contribution.xlsx

Excel file with 3 sheets (one per person). Sheet names should contain the person's name (e.g., "Kishore Data count"). The app auto-detects column formats:

- Dates/months --> grouped by month
- Numbers --> treated as count
- URLs --> stored as links
- Text --> category labels

---

## Customizing

### Change file names

Edit these lines at the top of `app.py`:

```python
HUBSPOT_CSV = os.path.join(DATA_DIR, "hubspot_contacts.csv")
TCS_EMAIL_MKT_CSV = os.path.join(DATA_DIR, "tcs_email_marketing.csv")
BW_EMAIL_MKT_CSV = os.path.join(DATA_DIR, "binaryworks_email_marketing.csv")
CONTRIBUTION_XLSX = os.path.join(DATA_DIR, "contribution.xlsx")
```

### Add ConversionBox email marketing later

Uncomment and update this line in `app.py`:

```python
# CB_EMAIL_MKT_CSV = os.path.join(DATA_DIR, "conversionbox_email_marketing.csv")
```

### Change TAG mappings

Edit these sections in `app.py`:

```python
BRAND_TAGS = {
    "TCS": "tcs",
    "BinaryWorks": "drupal",
    "ConversionBox": "conversionbox",
}

OTHERS_TAGS = ["Manufacturing", "ConversionBox 200+ Products", "Higher Education", "Finance"]
```

### Change which technologies show individually

```python
TCS_MAIN_TECHS = ["Shopify", "BigCommerce", "WooCommerce", "Magento", "Shopify Plus"]
DRUPAL_MAIN_CATS = ["Drupal 7", "Drupal 8", "Drupal 9", "Drupal 10", "Drupal 11", "WordPress"]
```

---

## Missing files? No problem.

The dashboard handles missing files gracefully. If a file doesn't exist, that section just shows "No data found" with the expected file path. Add files anytime and refresh.

---

## Free Deployment (Streamlit Cloud)

1. Push this folder to GitHub
2. Go to https://share.streamlit.io
3. Connect your repo, select `app.py`
4. Deploy

Your dashboard will be live at `https://your-app.streamlit.app`

Update data by pushing new files to GitHub.
