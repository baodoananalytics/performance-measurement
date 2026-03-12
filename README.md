# Portfolio Performance Measurement Report
Automated Portfolio Performance Measurement Report conducted by the combinations of Python and Excel.

# Overview
This project supports the investment advisory team in generating timely, automated monthly portfolio performance reports. Raw data is cleaned and transformed into key metrics using Python, then combined with VNINDEX as the benchmark. The processed data is loaded into an Excel workbook, where OFFSET, INDEX, and MATCH functions are used to dynamically populate the report page. By simply changing the input period, users can instantly view the report for any historical month. The final output is a fully completed, well-designed A4-sized report created entirely in Excel.

# Technologies Used
![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)
![Pandas](https://img.shields.io/badge/pandas-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white)
![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)

# Data Overview
The pipeline begins with reading raw data in `Portfolio_Raw.xlsx`, which contains daily recommendations in "port" sheet and VNINDEX's daily-closed point. Pandas package is used to read data, calculate each recommendation's return then aggregate to the monthly performance compared to the benchmark VNINDEX in monthly and year-to-date basis. The volatility of performance in report month and Top 5 outstanding stocks are also highlighted. Then some data-manipulation functions such as OFFSET, INDEX, MATCH are used to show key metrics on Report Page. Finally, the Report Page is well-design to readily generate to a final report in pdf format.

The pipeline aggregates high-quality market intelligence from Vietnam’s most prestigious financial and real estate e-magazines, including CafeF.vn, Vietstock.vn, and Batdongsan.com.vn.

1. **Comprehensive Metadata Extraction:**
For every captured article, the workflow extracts a structured dataset to ensure full traceability:
    * Source Attribution: Originating publisher.
    * Metadata: Article Title and a concise Summary.
    * Direct Access: Original Source URL for deep-dive reading.
    * Temporal Tracking: Exact Published Time to maintain chronological relevance.

2. **Intelligent Keyword-based Filtering:**
To maintain the news content exactly, the script implements a strict Boolean filtering logic. An article is only processed if its title contains at least one industry-specific keyword:
    * Real Estate Sector: Targeted via terms such as "Nhà ở" (Housing), "Dự án" (Projects).
    * Construction Materials: Filtered by strategic keywords like "Thép" (Steel) and "Vật liệu xây dựng" (Construction Materials).

# Project Structure
``` text
├── data/
│   ├── Portfolio_Raw.xlsx
│   ├── Report_Template.xlsx
├── code/
│   └── main.py
│   └── hardcode.py
├── demo/
│   ├── Report_Sample.pdf
├── LICENSE
└── README.md
```

# Project Steps
1. The pipeline starts with extracting raw data from Portfolio_Raw.xlsx, including daily stock recommendations in the port sheet and the daily closing level of VNINDEX.
  
2. Python pandas is used to calculate individual recommendation returns and consolidate them into monthly and year-to-date portfolio performance versus the VNINDEX benchmark. The output also captures monthly performance volatility and the top 5 outperforming stocks.
  
3. Key metrics are then displayed dynamically on the Report Page using Excel functions such as OFFSET, INDEX, and MATCH.
  
4. The final Report Page is fully formatted and ready for PDF export.

---
# Scale up
This project featured a highly flexible design. By simply updating the parameters such as keywords. sources and URLs in the hardcode.py module, we will have other industry-specific news with minimal efforts.

# Author
*  [Linkedin](https://www.linkedin.com/in/bao-doan-833a6615a/)
