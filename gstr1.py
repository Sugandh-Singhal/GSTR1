import streamlit as st
import pandas as pd
import io
import xlsxwriter

# Template columns
TEMPLATE_COLUMNS = [
    "GSTIN of the Tax Payer*", "Return Period*", "Gross Turnover", "Gross Turnover - April to June, 2017",
    "Document Category", "Document Number*", "Invoice Type*", "Invoice Sub_Type*", "Nature of Supply*",
    "Reverse Charge*", "Reason Code for issuing Debit/Credit Note", "Pre GST Regime Dr./ Cr. Notes*",
    "GSTR2 filing status of counter party", "Counter Party GSTIN/UID*", "GSTIN Ecom Operator*",
    "Counter Party Name", "Invoice Number*", "Invoice Date*", "Line item Number*",
    "Identifier of Goods or Services*", "HSN or SAC of Goods or Services*", "Description of Item",
    "Quantity of goods sold*", "UQC (Unit of Measure) of goods sold*", "Invoice Value*",
    "Taxable value of Goods or Services*", "Tax Rate*", "IGST Amount*", "CGST Amount*",
    "SGST Amount*", "CESS Amount*", "Place of Supply*", "Port Code*",
    "Shipping Bill or Bill of Export Number*", "Shipping Bill or Bill of Export Date*",
    "Original Invoice Number*", "Original Invoice Date*", "Advance Amount Received /Adjusted*",
    "Advance Reference Number", "Ship to Country", "Plant code/BU code", "GL code",
    "Adjustment of any Advance against this Invoice on which you have paid tax", "Diff Percent*"
]

# Column mappings
CUSTOM_COLUMN_MAPPINGS = {
    "My GSTIN": "GSTIN of the Tax Payer*",
    "Customer Billing GSTIN": "Counter Party GSTIN/UID*",
    "Invoice Date": "Invoice Date*",
    "Invoice Number": "Invoice Number*",
    "Customer Billing Name": "Counter Party Name",
    "HSN or SAC code": "HSN or SAC of Goods or Services*",
    "Item desciption": "Description of Item",
    "Item Taxable Value *": "Taxable value of Goods or Services*",
    "Total Transaction Value": "Invoice Value*",
    "CGST Amount": "CGST Amount*",
    "SGST Amount": "SGST Amount*",
    "IGST Amount": "IGST Amount*",
    "Item quantity": "Quantity of goods sold*",
    "Item Unit of Measurement": "UQC (Unit of Measure) of goods sold*"
}

# UI
st.title("GST Excel Processor using XlsxWriter")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
return_period = st.text_input("Enter 6-digit Return Period (YYYYMM)", max_chars=6)

if uploaded_file and return_period.isdigit() and len(return_period) == 6:
    try:
        df_uploaded = pd.read_excel(uploaded_file)
        df_uploaded.columns = (
            df_uploaded.columns
            .str.strip()
            .str.replace(r"\*+$", "", regex=True)
        )

        df_uploaded.dropna(how='all', inplace=True)
        df_uploaded = df_uploaded[
            ~df_uploaded["Invoice Number"].astype(str).str.strip().str.lower().eq("total") &
            df_uploaded["Invoice Number"].notna()
        ]
        df_uploaded.reset_index(drop=True, inplace=True)

        df_template = pd.DataFrame(columns=TEMPLATE_COLUMNS)

        for col in df_uploaded.columns:
            if col in df_template.columns:
                df_template[col] = df_uploaded[col]

        for src_col, tgt_col in CUSTOM_COLUMN_MAPPINGS.items():
            if src_col in df_uploaded.columns:
                df_template[tgt_col] = df_uploaded[src_col]

        df_template = df_template.iloc[:len(df_uploaded)].copy()

        df_template["Return Period*"] = return_period
        df_template["Reverse Charge*"] = "N"
        df_template["Document Category"] = "01-Inv. for Outward Supply"
        df_template["Document Number*"] = 1

        hsn = df_template["HSN or SAC of Goods or Services*"].astype(str)
        df_template.loc[hsn.str.startswith("99", na=False), "Identifier of Goods or Services*"] = "S"

        exp_s = (df_template["Invoice Type*"] == "EXP") & (df_template["Identifier of Goods or Services*"] == "S")
        df_template.loc[exp_s, "Quantity of goods sold*"] = ""
        df_template.loc[exp_s, "UQC (Unit of Measure) of goods sold*"] = ""

        no_gstin = df_template["Counter Party GSTIN/UID*"].isna() | \
                   df_template["Counter Party GSTIN/UID*"].astype(str).str.strip().eq("")
        df_template.loc[no_gstin, "Invoice Type*"] = "EXP"
        df_template.loc[no_gstin, "Invoice Sub_Type*"] = "WOPAY"
        df_template.loc[no_gstin, "Nature of Supply*"] = "Inter"

        has_gstin = df_template["Counter Party GSTIN/UID*"].astype(str).str.strip().ne("")
        has_igst = df_template["IGST Amount*"].notna() & (df_template["IGST Amount*"] != 0)
        b2b_mask = has_gstin & has_igst
        df_template.loc[b2b_mask, "Invoice Type*"] = "B2B"
        df_template.loc[b2b_mask, "Invoice Sub_Type*"] = "R"

        gstin_prefix = df_template["GSTIN of the Tax Payer*"].astype(str).str[:2]
        counter_prefix = df_template["Counter Party GSTIN/UID*"].astype(str).str[:2]
        df_template.loc[has_gstin, "Nature of Supply*"] = [
            "Intra" if a == b else "Inter" for a, b in zip(gstin_prefix, counter_prefix)
        ]

        not_exp = df_template["Invoice Type*"] != "EXP"
        df_template.loc[not_exp, "Place of Supply*"] = counter_prefix[not_exp]

        df_template["Line item Number*"] = (
            df_template.groupby("Invoice Number*", dropna=False).cumcount() + 1
        )

        for tax_col in ["Invoice Value*", "IGST Amount*", "CGST Amount*", "SGST Amount*"]:
            df_template[tax_col] = pd.to_numeric(df_template[tax_col], errors="coerce")

        df_template["_IGST"] = df_template["IGST Amount*"].fillna(0)
        df_template["_CGST"] = df_template["CGST Amount*"].fillna(0)
        df_template["_SGST"] = df_template["SGST Amount*"].fillna(0)
        df_template["_base"] = df_template["Invoice Value*"].fillna(0)

        df_template["_total_invoice_value"] = (
            df_template["_base"] + df_template["_IGST"] + df_template["_CGST"] + df_template["_SGST"]
        )
        invoice_totals = df_template.groupby("Invoice Number*")["_total_invoice_value"].transform("sum")
        df_template["Invoice Value*"] = invoice_totals

        df_template.drop(columns=["_total_invoice_value", "_IGST", "_CGST", "_SGST", "_base"], inplace=True)

        for tax_col in ["IGST Amount*", "CGST Amount*", "SGST Amount*"]:
            df_template[tax_col] = df_template[tax_col].apply(lambda x: "" if pd.isna(x) else x)

        df_template.loc[df_template["Invoice Type*"] != "EXP", "Tax Rate*"] = 18

        st.success("✅ File processed successfully!")
        st.dataframe(df_template)

        # Write to Excel with dropdown using XlsxWriter
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_template.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            dropdowns = {
                "Document Category": [
                    "01-Inv.outward supply", "02-Inv.for inward supply from unregistered person",
                    "03-Revised Income", "04-DN", "05-CN", "06-Receipt Voucher", "07-Payment Voucher",
                    "08-Refund Voucher", "09-job work", "10-supply on approval", "11-liquid gas",
                    "12-other than by way of supply"
                ],
                "Document Number*": [str(i) for i in range(1, 13)],
                "Invoice Type*": ["B2B", "B2C", "EXP", "B2CS"],
                "Invoice Sub_Type*": [
                    "R", "DE", "SEWOP", "SEWP", "NR", "EXMPT", "NGST", "E", "OE", "WPAY", "WOPAY"
                ],
                "Nature of Supply*": ["Inter", "Intra"],
                "Reverse Charge*": ["Y", "N"],
                "Identifier of Goods or Services*": ["G", "S"],
                "Place of Supply*": [str(i) for i in range(1, 38)]
            }

            header = df_template.columns.tolist()
            for col_name, options in dropdowns.items():
                if col_name in header:
                    col_idx = header.index(col_name)
                    for row in range(1, len(df_template) + 1):
                        worksheet.data_validation(row, col_idx, row, col_idx, {
                            'validate': 'list',
                            'source': options
                        })

        st.download_button(
            label="📥 Download Processed Excel (with dropdowns)",
            data=output.getvalue(),
            file_name="processed_gst_file_with_dropdowns.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")

elif uploaded_file:
    st.warning("⚠️ Please enter a valid 6-digit Return Period (e.g., 202407)")
