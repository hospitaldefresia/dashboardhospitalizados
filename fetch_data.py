name: Actualizar datos censo

on:
  schedule:
    - cron: '0 6 * * *'
  workflow_dispatch:

jobs:
  update:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Install dependencies
        run: pip install requests msal openpyxl

      - name: Fetch data from SharePoint
        env:
          CLIENT_ID: ${{ secrets.CLIENT_ID }}
          TENANT_ID: ${{ secrets.TENANT_ID }}
          CLIENT_SECRET: ${{ secrets.CLIENT_SECRET }}
          SHAREPOINT_FILE_URL: ${{ secrets.SHAREPOINT_FILE_URL }}
        run: python fetch_data.py

      - name: Commit and push if changed
        run: |
          git config user.name "GitHub Actions"
          git config user.email "actions@github.com"
          git add data/censo.json
          git diff --staged --quiet || git commit -m "Actualizar datos censo $(date +'%Y-%m-%d')"
          git push
