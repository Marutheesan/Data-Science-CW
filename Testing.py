import streamlit as st
import plotly_express as px
import pandas as pd
import os
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="Global Superstore", page_icon=":chart_with_upwards_trend:",layout="wide")

st.title(" :chart_with_upwards_trend: Global SuperStore Sales Analysis")

fl = st.file_uploader(":file_folder: Upload a file",type=(["xlsx"]))
if fl is not None:
    filename = fl.name
    st.write(filename)
    df = pd.read_excel(filename)
else:
    os.chdir(r"C:\Users\marut\OneDrive\Desktop\DS Str")
    df = pd.read_excel("Global Superstore lite.xlsx")

col1, col2 = st.columns((2))
df["Order Date"] = pd.to_datetime(df["Order Date"])

# Getting the min and max date 
startDate = pd.to_datetime(df["Order Date"]).min()
endDate = pd.to_datetime(df["Order Date"]).max()

with col1:
    date1 = pd.to_datetime(st.date_input("Start Date", startDate))

with col2:
    date2 = pd.to_datetime(st.date_input("End Date", endDate))

df = df[(df["Order Date"] >= date1) & (df["Order Date"] <= date2)].copy()


# Load data from Excel file
try:
    df = pd.read_excel("Global Superstore lite.xlsx")
except PermissionError:
    st.error("Permission denied: Unable to read Excel file. Please check file permissions or file path.")

# Sidebar Filters
st.sidebar.title('Filter Options')


# Create for Region
region = st.sidebar.multiselect("Pick your Region", df["Region"].unique())
if not region:
    df2 = df.copy()
else:
    df2 = df[df["Region"].isin(region)]

# Create for Country
country = st.sidebar.multiselect("Pick the Country", df2["Country"].unique())
if not country:
    df3 = df2.copy()
else:
    df3 = df2[df2["Country"].isin(country)]

# Create for State
state = st.sidebar.multiselect("Pick the State", df3["State"].unique())
if not state:
    df4 = df3.copy()
else:
    df4 = df3[df3["State"].isin(state)]

# Create for City
city = st.sidebar.multiselect("Pick the City", df4["City"].unique())
if not city:
    df5 = df4.copy()
else:
    df5 = df4[df4["City"].isin(city)]

# Filter the data based on Region, Country, State, and City

if not region and not country and not state and not city:
    filtered_df = df
elif not country and not state and not city:
    filtered_df = df[df["Region"].isin(region)]
elif not region and not state and not city:
    filtered_df = df[df["Country"].isin(country)]
elif state and city and not region and not country:
    filtered_df = df4[df4["State"].isin(state) & df4["City"].isin(city)]
elif region and city and not country and not state:
    filtered_df = df4[df4["Region"].isin(region) & df4["City"].isin(city)]
elif region and state and not city and not country:
    filtered_df = df4[df4["Region"].isin(region) & df4["State"].isin(state)]
elif city and not country and not state:
    filtered_df = df4[df4["City"].isin(city)]
elif country and not state and not city:
    filtered_df = df4[df4["Country"].isin(country)]
else:
    filtered_df = df5[df5["Region"].isin(region) & df5["Country"].isin(country) & df5["State"].isin(state) & df5["City"].isin(city)]

SubCategory_df = filtered_df.groupby(by = ["Sub-Category"], as_index= False)["Sales"].sum()

with col1:
    st.subheader("Sales by Sub-Category")
    fig = px.bar(SubCategory_df, x = "Sub-Category", y = "Sales", text = ['${:,.2f}'.format(x) for x in SubCategory_df["Sales"]],
                 template = "seaborn")
    st.plotly_chart(fig,use_container_width=True, height = 200)

with col2:
    st.subheader("Sales by Region")
    fig = px.pie(filtered_df, values = "Sales", names = "Region", hole = 0.5)
    fig.update_traces(text = filtered_df["Region"], textposition = "outside")
    st.plotly_chart(fig,use_container_width=True)

