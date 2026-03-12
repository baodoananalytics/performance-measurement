# Portfolio Performance Measurement Report
Automated Portfolio Performance Measurement Report conducted by the combinations of Python and Excel.

# Overview
This project supports the investment advisory team in generating timely, automated monthly portfolio performance reports. Raw data is cleaned and transformed into key metrics using Python, then combined with VNINDEX as the benchmark. The processed data is loaded into an Excel workbook, where OFFSET, INDEX, and MATCH functions are used to dynamically populate the report page. By simply changing the input period, users can instantly view the report for any historical month. The final output is a fully completed, well-designed A4-sized report created entirely in Excel.

# Technologies Used
![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)
![Pandas](https://img.shields.io/badge/pandas-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white)
![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)

# Data Overview
The whole process only uses the raw data in `Portfolio_raw.xlsx`, which contains 2 sheets: `port` and `vni`.

1. **Portfolio Data:** Containing daily stock recommendations to buy, including:
      * "Recommend_Date": datetime
      * "Ticker": string
      * "Price" (the stock price on Recommend_Date"): decimal
      * "Buy_Range_1", "Buy_Range_2" (the optimal price range for position opening): decimal
      * "Target_1", "Target_2", "Cutloss" (the price where investors should take profit or cut loss): decimal
      * "Open_Date" (when the positions are opened): datetime
      * "Close_Date" (when the positions are closed): datetime
      * "Open_Price" (the actual stock price investor may buy): decimal
      * "Close_Price" (the actual stock price investor may sell): decimal

2. **Benchmark data:** Including VNINDEX daily-closed price
     
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
1. The pipeline starts with extracting raw data from `Portfolio_Raw.xlsx`, including daily stock recommendations in the port sheet and the daily closing level of VNINDEX.
  
2. Python pandas is used to calculate individual recommendation returns and consolidate them into monthly and year-to-date portfolio performance versus the VNINDEX benchmark. The output also captures monthly performance volatility and the top 5 outperforming stocks.
  
3. Key metrics are then displayed dynamically on the Report Page using Excel functions such as OFFSET, INDEX, and MATCH.
  
4. The final Report Page is fully formatted and ready for PDF export.

---
# Scale up
This project features a highly flexible design. By simply changing the reporting month input (e.g., `2025-11` or `2025-09`), users can instantly generate the report for the selected month.

# Author
*  [Linkedin](https://www.linkedin.com/in/bao-doan-833a6615a/)
