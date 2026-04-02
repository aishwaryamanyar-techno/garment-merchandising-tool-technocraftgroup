import streamlit as st
import pandas as pd

st.title("Garment Merchandising Order Tracking Tool")

# -------- Manual Fields --------

order_receipt_date = st.date_input("Order Receipt Date")
sales_order = st.text_input("Sales Order No (ERP)")
delivery_date = st.date_input("Committed Delivery Date")

market = st.text_input("Market")
marketing_head = st.text_input("Marketing Head")
merchant_name = st.text_input("Merchant Name")
customer_name = st.text_input("Customer Name")

fabric_desc = st.text_input("Fabric Quality Description")
style = st.text_input("Style")
product_category = st.text_input("Product Category")

sam = st.number_input("SAM")
sam_category = st.text_input("SAM Category")

order_qty = st.number_input("Order Qnty Pcs")

sale_rate = st.number_input("Sale Rate In Rs./USD")
usd_rate = st.number_input("Rs./USD")

export_incentive = st.number_input("Export Incentives Value")
freight = st.number_input("Garment Freight Rs/Pcs")

fabric_source = st.selectbox("Fabric Source", ["Inhouse","Outside"])

fabric_body_tpt = st.number_input("Fabric Body TPT Rs/Kg")

fabric1_cost = st.number_input("Fabric-1 Landed Cost")
fabric1_usage = st.number_input("Fabric-1 Usage")

fabric2_cost = st.number_input("Fabric-2 Landed Cost")
fabric2_usage = st.number_input("Fabric-2 Usage")

trims = st.number_input("Thread & Trims Rs/Piece")
outsource = st.number_input("Outsource Rs/Piece")

# -------- Automated Calculations --------

sale_rate_rs = sale_rate * usd_rate
sale_value_lacs = (order_qty * sale_rate_rs) / 100000
incentive_rs = export_incentive
net_sale_price = sale_rate_rs + incentive_rs - freight
net_sale_value = (net_sale_price * order_qty) / 100000
fabric_cost = (fabric1_cost * fabric1_usage) + (fabric2_cost * fabric2_usage)
tpt_rs_pcs = net_sale_price - fabric_cost - trims - outsource
tpt_percent = (tpt_rs_pcs / net_sale_price * 100) if net_sale_price != 0 else 0
tpt_rs_min = (tpt_rs_pcs / sam) if sam != 0 else 0

# -------- Show Auto Values --------

st.subheader("Automatic Calculations")

st.write("Sale Rate Rs/Pcs:", sale_rate_rs)
st.write("Sale Value in Lacs:", sale_value_lacs)
st.write("Incentive Rs/Pcs:", incentive_rs)
st.write("Net Sale Price:", net_sale_price)
st.write("Net Sale Value:", net_sale_value)
st.write("Fabric Cost Rs/Pcs:", fabric_cost)
st.write("TPT Rs/Pcs:", tpt_rs_pcs)
st.write("TPT %:", tpt_percent)
st.write("TPT Rs/Minute:", tpt_rs_min)

# -------- Save Data --------

if st.button("Submit Order"):

    file_path = r"G:\My Drive\MASTER FILE FOR GARMENT MERCHANDISING ORDER TRACKING.xlsx"

    data = {
        "Order Receipt Date": order_receipt_date,
        "Sales Order No (ERP)": sales_order,
        "Committed Delivery Date": delivery_date,
        "Market": market,
        "Marketing Head": marketing_head,
        "Merchant Name": merchant_name,
        "Customer Name": customer_name,
        "Fabric Quality Description": fabric_desc,
        "Style": style,
        "Product Category": product_category,
        "SAM": sam,
        "SAM Category": sam_category,
        "Order Qnty Pcs": order_qty,
        "Sale Rate In Rs./USD": sale_rate,
        "Rs./USD": usd_rate,
        "Export Incentives Value": export_incentive,
        "Garment Freight Rs/Pcs": freight,
        "Fabric Source": fabric_source,
        "Fabric Body TPT Rs/Kg": fabric_body_tpt,
        "Fabric-1 Landed Cost": fabric1_cost,
        "Fabric-1 Usage": fabric1_usage,
        "Fabric-2 Landed Cost": fabric2_cost,
        "Fabric-2 Usage": fabric2_usage,
        "Thread & Trims Rs/Piece": trims,
        "Outsource Rs/Piece": outsource,

        "Sale Rate Rs/Pcs": sale_rate_rs,
        "Sale Value in Lacs": sale_value_lacs,
        "Incentive Rs/Pcs": incentive_rs,
        "Net Sale Price": net_sale_price,
        "Net Sale Value": net_sale_value,
        "Fabric Cost Rs/Pcs": fabric_cost,
        "TPT Rs/Pcs": tpt_rs_pcs,
        "TPT %": tpt_percent,
        "TPT Rs/Minute": tpt_rs_min
    }

    df = pd.DataFrame([data])

    try:
        old = pd.read_excel(file_path)

        sr_no = len(old) + 1
        df.insert(0, "Sr No", sr_no)

        new = pd.concat([old, df], ignore_index=True)

    except:
        df.insert(0, "Sr No", 1)
        new = df

    new.to_excel(file_path, index=False)

    st.success("Order Saved Successfully")