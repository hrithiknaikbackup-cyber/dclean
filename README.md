# Account & Contact Quality Check

Single-page Node.js + vanilla JavaScript app for cleaning uploaded account/contact workbooks and generating a management-ready Excel report.

## What it does
- Cleans and standardizes account/contact fields inside one uploaded workbook
- Compares records within that same file
- Builds duplicate clusters (groups, not only pairs)
- Suggests a master record using completeness-first logic
- Calculates data health score and duplicate confidence
- Produces field completeness, geographic breakdown, and orphan insights
- Downloads a polished `.xlsx` report

## Run locally
```bash
npm install
npm start
```

Then open `http://localhost:3000`.
