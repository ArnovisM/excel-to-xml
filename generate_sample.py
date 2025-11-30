import pandas as pd

data = {
    "Office of Dept/transfer/exit": ["LV02", "LV02"],
    "Voyage Number": ["AA2217", "AA2218"],
    "Date of departure": ["27.11.2025", "28.11.2025"],
    "Waybill reference number": ["1014411101001-2", "1014411101002-3"],
    "Waybill Type": ["T1", "T1"],
    "Previous document": ["1014421101001", "1014421101002"],
    "Place of loading": ["JAMA", "RIGA"],
    "Carrier": ["AAL", "DHL"],
    "Consignee": ["SHAMARA FEDOROV", "JOHN DOE"],
    "Total Containers": ["1 CTNR", "2 CTNR"],
    "Packages": ["11", "20"],
    "Manifested Package": ["11", "20"],
    "Manifested Gross Weight": ["18", "30"],
    "Description of goods": ["GENERAL GOODS", "ELECTRONICS"]
}

df = pd.DataFrame(data)
df.to_excel("sample_manifest.xlsx", index=False)
print("sample_manifest.xlsx created")
