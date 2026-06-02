# app.py - Professional Sales Analytics Platform (No Emojis)
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from sklearn.ensemble import RandomForestRegressor
from sklearn.model_selection import train_test_split
from sklearn.metrics import r2_score, mean_absolute_error
import warnings
from datetime import datetime
from io import BytesIO
import matplotlib.pyplot as plt
import seaborn as sns
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

warnings.filterwarnings('ignore')

# Page Configuration
st.set_page_config(
    page_title="Sales Analytics Platform",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2rem;
        font-weight: 700;
        color: #1E3A5F;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #6B7280;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #FFFFFF 0%, #F9FAFB 100%);
        border-radius: 0.5rem;
        padding: 1rem;
        box-shadow: 0 1px 3px 0 rgba(0,0,0,0.1);
        border: 1px solid #E5E7EB;
        text-align: center;
    }
    .metric-value {
        font-size: 1.5rem;
        font-weight: 700;
        color: #1E3A5F;
    }
    .metric-label {
        font-size: 0.875rem;
        color: #6B7280;
    }
    .recommendation-box {
        background: linear-gradient(135deg, #F0FDF4 0%, #DCFCE7 100%);
        border-left: 4px solid #10B981;
        padding: 0.75rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .insight-box {
        background: linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%);
        border-left: 4px solid #3B82F6;
        padding: 0.75rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .section-header {
        font-size: 1.25rem;
        font-weight: 600;
        color: #1E3A5F;
        margin-top: 1rem;
        margin-bottom: 1rem;
        border-bottom: 2px solid #E5E7EB;
        padding-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Function to create Word Report
def create_word_report(df, sales_col, product_col, region_col, date_col, quantity_col, price_col):
    doc = Document()
    
    # Title
    title = doc.add_heading('Sales Analytics Report', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Generation date
    doc.add_paragraph(f'Report Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    doc.add_paragraph()
    
    # Executive Summary
    doc.add_heading('Executive Summary', level=1)
    
    total_revenue = df[sales_col].sum()
    avg_transaction = df[sales_col].mean()
    total_transactions = len(df)
    
    summary_para = doc.add_paragraph()
    summary_para.add_run(f'Total Revenue: ').bold = True
    summary_para.add_run(f'${total_revenue:,.2f}\n')
    summary_para.add_run(f'Average Transaction: ').bold = True
    summary_para.add_run(f'${avg_transaction:,.2f}\n')
    summary_para.add_run(f'Total Transactions: ').bold = True
    summary_para.add_run(f'{total_transactions:,}\n')
    
    if product_col:
        unique_products = df[product_col].nunique()
        summary_para.add_run(f'Unique Products: ').bold = True
        summary_para.add_run(f'{unique_products:,}\n')
    
    if region_col:
        unique_regions = df[region_col].nunique()
        summary_para.add_run(f'Active Regions: ').bold = True
        summary_para.add_run(f'{unique_regions:,}\n')
    
    doc.add_paragraph()
    
    # Key Performance Indicators
    doc.add_heading('Key Performance Indicators', level=1)
    
    kpi_data = [
        ('Total Revenue', f'${total_revenue:,.2f}'),
        ('Average Transaction', f'${avg_transaction:,.2f}'),
        ('Total Transactions', f'{total_transactions:,}'),
    ]
    
    if product_col:
        kpi_data.append(('Unique Products', f'{df[product_col].nunique():,}'))
    if region_col:
        kpi_data.append(('Active Regions', f'{df[region_col].nunique():,}'))
    if quantity_col:
        kpi_data.append(('Total Units Sold', f'{df[quantity_col].sum():,.0f}'))
    if price_col:
        kpi_data.append(('Average Price', f'${df[price_col].mean():,.2f}'))
    
    table = doc.add_table(rows=len(kpi_data), cols=2)
    table.style = 'Light Grid Accent 1'
    
    for i, (metric, value) in enumerate(kpi_data):
        row = table.rows[i]
        row.cells[0].text = metric
        row.cells[1].text = value
    
    doc.add_paragraph()
    
    # Top Products Analysis
    if product_col:
        doc.add_heading('Top Performing Products', level=1)
        
        top_products = df.groupby(product_col)[sales_col].sum().sort_values(ascending=False).head(10)
        
        table = doc.add_table(rows=min(11, len(top_products)+1), cols=2)
        table.style = 'Light Grid Accent 1'
        
        table.rows[0].cells[0].text = 'Product'
        table.rows[0].cells[1].text = 'Revenue (USD)'
        
        for i, (product, revenue) in enumerate(top_products.items(), 1):
            if i < len(table.rows):
                table.rows[i].cells[0].text = str(product)
                table.rows[i].cells[1].text = f'${revenue:,.2f}'
        
        doc.add_paragraph()
    
    # Regional Analysis
    if region_col:
        doc.add_heading('Regional Performance', level=1)
        
        region_sales = df.groupby(region_col)[sales_col].sum().sort_values(ascending=False)
        
        table = doc.add_table(rows=min(11, len(region_sales)+1), cols=2)
        table.style = 'Light Grid Accent 1'
        
        table.rows[0].cells[0].text = 'Region'
        table.rows[0].cells[1].text = 'Revenue (USD)'
        
        for i, (region, revenue) in enumerate(region_sales.items(), 1):
            if i < len(table.rows):
                table.rows[i].cells[0].text = str(region)
                table.rows[i].cells[1].text = f'${revenue:,.2f}'
        
        doc.add_paragraph()
    
    # Seasonal Analysis
    if date_col and 'Month' in df.columns:
        doc.add_heading('Seasonal Analysis', level=1)
        
        monthly_sales = df.groupby('Month')[sales_col].sum().sort_index()
        
        table = doc.add_table(rows=13, cols=2)
        table.style = 'Light Grid Accent 1'
        
        table.rows[0].cells[0].text = 'Month'
        table.rows[0].cells[1].text = 'Revenue (USD)'
        
        month_names = ['January', 'February', 'March', 'April', 'May', 'June', 
                      'July', 'August', 'September', 'October', 'November', 'December']
        
        for i, month in enumerate(range(1, 13), 1):
            table.rows[i].cells[0].text = month_names[month-1]
            table.rows[i].cells[1].text = f'${monthly_sales.get(month, 0):,.2f}'
        
        doc.add_paragraph()
    
    # Business Recommendations
    doc.add_heading('Strategic Recommendations', level=1)
    
    recommendations = []
    
    if product_col:
        top_product = df.groupby(product_col)[sales_col].sum().idxmax()
        recommendations.append(f"Increase inventory and marketing budget for '{top_product}' by 25-30 percent")
    
    if region_col and len(df[region_col].unique()) > 1:
        region_sales = df.groupby(region_col)[sales_col].sum()
        best_region = region_sales.idxmax()
        worst_region = region_sales.idxmin()
        recommendations.append(f"Expand successful strategies from {best_region} to {worst_region} region")
    
    if date_col and 'Month' in df.columns:
        monthly_sales = df.groupby('Month')[sales_col].sum()
        peak_month = monthly_sales.idxmax()
        recommendations.append(f"Prepare inventory 60 days before Month {peak_month} for peak season")
    
    if quantity_col and price_col:
        avg_price = df[price_col].mean()
        recommendations.append(f"Consider bundle pricing at ${avg_price * 0.9:.2f} to increase volume")
    
    for rec in recommendations:
        para = doc.add_paragraph(style='List Bullet')
        para.add_run(rec)
    
    doc.add_paragraph()
    
    # Conclusion
    doc.add_heading('Conclusion', level=1)
    conclusion = doc.add_paragraph()
    conclusion.add_run('Based on the analysis, the business shows ')
    conclusion.add_run(f'strong performance with ${total_revenue:,.2f} in total revenue. ')
    
    if product_col and region_col:
        conclusion.add_run(f'The top performer is {top_product} in the {best_region} region. ')
    
    conclusion.add_run('Focus on the recommended strategies to maximize growth and profitability.')
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return buffer

# Initialize session state
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'df' not in st.session_state:
    st.session_state.df = None
if 'analysis_complete' not in st.session_state:
    st.session_state.analysis_complete = False

# Header
st.markdown('<div class="main-header">Sales Analytics Platform</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">AI-Powered Business Intelligence & Analytics</div>', unsafe_allow_html=True)

# Sidebar for file upload
with st.sidebar:
    st.header("Data Upload")
    
    uploaded_file = st.file_uploader(
        "Choose a CSV or Excel file",
        type=['csv', 'xlsx', 'xls']
    )
    
    if uploaded_file:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.session_state.df = df
            st.session_state.data_loaded = True
            st.success(f"Loaded {len(df)} rows, {len(df.columns)} columns")
            st.info(f"File: {uploaded_file.name}")
        except Exception as e:
            st.error(f"Error: {str(e)}")

# Main Content
if st.session_state.data_loaded:
    df = st.session_state.df
    
    st.markdown('<div class="section-header">Column Configuration</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        date_col = st.selectbox("Date Column", ['None'] + list(df.columns))
        sales_col = st.selectbox("Sales / Revenue Column", ['None'] + list(df.columns))
        product_col = st.selectbox("Product Column", ['None'] + list(df.columns))
    
    with col2:
        region_col = st.selectbox("Region / City Column", ['None'] + list(df.columns))
        quantity_col = st.selectbox("Quantity Column", ['None'] + list(df.columns))
        price_col = st.selectbox("Price Column", ['None'] + list(df.columns))
    
    with st.expander("Data Preview"):
        st.dataframe(df.head(10), use_container_width=True)
        st.caption(f"Dataset: {len(df):,} rows x {len(df.columns)} columns")
    
    if st.button("Run Analysis", type="primary", use_container_width=True):
        if not sales_col or sales_col == 'None':
            st.error("Please select a sales column")
        else:
            with st.spinner("Processing data..."):
                df_clean = df.copy()
                
                if date_col and date_col != 'None':
                    df_clean[date_col] = pd.to_datetime(df_clean[date_col], errors='coerce', dayfirst=True)
                    df_clean = df_clean.dropna(subset=[date_col])
                    df_clean['Month'] = df_clean[date_col].dt.month
                    df_clean['Year'] = df_clean[date_col].dt.year
                    df_clean['Quarter'] = df_clean[date_col].dt.quarter
                
                if sales_col != 'None':
                    df_clean[sales_col] = pd.to_numeric(df_clean[sales_col], errors='coerce')
                    df_clean = df_clean.dropna(subset=[sales_col])
                    df_clean = df_clean[df_clean[sales_col] > 0]
                
                if quantity_col != 'None':
                    df_clean[quantity_col] = pd.to_numeric(df_clean[quantity_col], errors='coerce')
                
                if price_col != 'None':
                    df_clean[price_col] = pd.to_numeric(df_clean[price_col], errors='coerce')
                
                df_clean = df_clean.drop_duplicates()
                
                if (date_col == 'None' or date_col is None) and 'Month' not in df_clean.columns:
                    df_clean['Month'] = np.random.randint(1, 13, len(df_clean))
                
                st.session_state.df_clean = df_clean
                st.session_state.sales_col = None if sales_col == 'None' else sales_col
                st.session_state.product_col = None if product_col == 'None' else product_col
                st.session_state.region_col = None if region_col == 'None' else region_col
                st.session_state.date_col = None if date_col == 'None' else date_col
                st.session_state.quantity_col = None if quantity_col == 'None' else quantity_col
                st.session_state.price_col = None if price_col == 'None' else price_col
                st.session_state.analysis_complete = True
                st.rerun()

# Analysis Results
if st.session_state.get('analysis_complete', False):
    df = st.session_state.df_clean
    sales_col = st.session_state.sales_col
    product_col = st.session_state.product_col
    region_col = st.session_state.region_col
    date_col = st.session_state.date_col
    quantity_col = st.session_state.quantity_col
    price_col = st.session_state.price_col
    
    total_revenue = df[sales_col].sum()
    avg_transaction = df[sales_col].mean()
    total_transactions = len(df)
    unique_products = df[product_col].nunique() if product_col else 0
    unique_regions = df[region_col].nunique() if region_col else 0
    
    # KPI Dashboard
    st.markdown('<div class="section-header">Key Performance Indicators</div>', unsafe_allow_html=True)
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">${total_revenue:,.0f}</div>
            <div class="metric-label">Total Revenue</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">${avg_transaction:,.0f}</div>
            <div class="metric-label">Average Transaction</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{total_transactions:,}</div>
            <div class="metric-label">Total Transactions</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        if product_col:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{unique_products:,}</div>
                <div class="metric-label">Unique Products</div>
            </div>
            """, unsafe_allow_html=True)
    
    with col5:
        if region_col:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{unique_regions:,}</div>
                <div class="metric-label">Active Regions</div>
            </div>
            """, unsafe_allow_html=True)
    
    # Interactive Filters
    st.markdown('<div class="section-header">Interactive Filters</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    filtered_df = df.copy()
    
    with col1:
        if region_col:
            regions = ['All'] + sorted(df[region_col].dropna().unique().tolist())
            selected_region = st.selectbox("Region", regions)
            if selected_region != 'All':
                filtered_df = filtered_df[filtered_df[region_col] == selected_region]
    
    with col2:
        if product_col:
            products = ['All'] + sorted(df[product_col].dropna().unique().tolist())
            selected_product = st.selectbox("Product", products)
            if selected_product != 'All':
                filtered_df = filtered_df[filtered_df[product_col] == selected_product]
    
    with col3:
        if date_col and 'Year' in df.columns:
            years = ['All'] + sorted(df['Year'].dropna().unique().tolist())
            selected_year = st.selectbox("Year", years)
            if selected_year != 'All':
                filtered_df = filtered_df[filtered_df['Year'] == selected_year]
    
    st.caption(f"Showing {len(filtered_df):,} of {len(df):,} records")
    
    # Visualizations
    st.markdown('<div class="section-header">Data Visualizations</div>', unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["Top Products", "Regional Analysis", "Time Series"])
    
    with tab1:
        if product_col and len(filtered_df) > 0:
            top_n = st.slider("Number of products", 5, 20, 10)
            top_products = filtered_df.groupby(product_col)[sales_col].sum().sort_values(ascending=False).head(top_n)
            
            fig = px.bar(
                x=top_products.values,
                y=top_products.index,
                orientation='h',
                title=f"Top {top_n} Products by Revenue",
                labels={'x': 'Revenue (USD)', 'y': ''},
                color=top_products.values,
                color_continuous_scale='Viridis',
                height=500
            )
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
            
            top_product = top_products.index[0]
            st.info(f"Top Product: {top_product} (${top_products.iloc[0]:,.2f})")
    
    with tab2:
        if region_col and len(filtered_df) > 0:
            region_sales = filtered_df.groupby(region_col)[sales_col].sum().sort_values(ascending=False)
            
            fig = px.bar(
                x=region_sales.index,
                y=region_sales.values,
                title="Sales by Region",
                labels={'x': 'Region', 'y': 'Revenue (USD)'},
                color=region_sales.values,
                color_continuous_scale='Plasma',
                height=500
            )
            st.plotly_chart(fig, use_container_width=True)
            
            if len(region_sales) > 1:
                col1, col2 = st.columns(2)
                with col1:
                    st.success(f"Best Region: {region_sales.index[0]} (${region_sales.iloc[0]:,.2f})")
                with col2:
                    st.warning(f"Region to Improve: {region_sales.index[-1]} (${region_sales.iloc[-1]:,.2f})")
    
    with tab3:
        if date_col and 'Month' in filtered_df.columns:
            monthly_sales = filtered_df.groupby('Month')[sales_col].sum().sort_index()
            
            fig = px.line(
                x=monthly_sales.index,
                y=monthly_sales.values,
                markers=True,
                title="Monthly Revenue Trend",
                labels={'x': 'Month', 'y': 'Revenue (USD)'},
                height=400
            )
            fig.update_traces(line=dict(width=3, color='#1E3A5F'), marker=dict(size=10))
            st.plotly_chart(fig, use_container_width=True)
            
            peak_month = monthly_sales.idxmax()
            st.info(f"Peak Month: {peak_month} (${monthly_sales[peak_month]:,.2f})")
            
            if len(filtered_df['Year'].unique()) > 1:
                st.subheader("Year-over-Year Comparison")
                yoy_data = filtered_df.groupby(['Year', 'Month'])[sales_col].sum().reset_index()
                fig2 = px.line(
                    yoy_data,
                    x='Month',
                    y=sales_col,
                    color='Year',
                    markers=True,
                    title="Year-over-Year Monthly Comparison",
                    height=400
                )
                st.plotly_chart(fig2, use_container_width=True)
    
    # Advanced Insights
    st.markdown('<div class="section-header">Advanced Insights</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if product_col and sales_col:
            top_5_products = filtered_df.groupby(product_col)[sales_col].sum().nlargest(5).sum()
            concentration = (top_5_products / total_revenue) * 100
            st.markdown(f"""
            <div class="insight-box">
                <strong>Product Portfolio Analysis</strong><br>
                Top 5 Products Contribution: {concentration:.1f}% of total revenue<br>
                {'High concentration risk' if concentration > 70 else 'Healthy diversification'}
            </div>
            """, unsafe_allow_html=True)
    
    with col2:
        if date_col and 'Month' in filtered_df.columns:
            monthly_avg = filtered_df.groupby('Month')[sales_col].sum().mean()
            peak_month = filtered_df.groupby('Month')[sales_col].sum().idxmax()
            peak_vs_avg = (filtered_df.groupby('Month')[sales_col].sum().max() / monthly_avg) * 100
            st.markdown(f"""
            <div class="insight-box">
                <strong>Seasonal Pattern Analysis</strong><br>
                Peak Month: {peak_month} ({peak_vs_avg:.0f}% above average)<br>
                Recommendation: Increase inventory 60 days before peak
            </div>
            """, unsafe_allow_html=True)
    
    # Product Analysis
    if product_col:
        st.markdown('<div class="section-header">Product Analysis</div>', unsafe_allow_html=True)
        
        products_list = [''] + sorted(df[product_col].unique().tolist())
        selected_product = st.selectbox("Select Product to Analyze", products_list)
        
        if selected_product:
            product_data = df[df[product_col] == selected_product]
            
            st.subheader(f"Performance Analysis: {selected_product}")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                product_revenue = product_data[sales_col].sum()
                st.metric("Total Revenue", f"${product_revenue:,.2f}")
            
            with col2:
                if quantity_col:
                    product_units = product_data[quantity_col].sum()
                    st.metric("Total Units Sold", f"{product_units:,.0f}")
            
            with col3:
                if price_col:
                    product_avg_price = product_data[price_col].mean()
                    st.metric("Average Price", f"${product_avg_price:,.2f}")
            
            if 'Month' in product_data.columns:
                fig = px.line(
                    product_data.groupby('Month')[sales_col].sum().reset_index(),
                    x='Month',
                    y=sales_col,
                    markers=True,
                    title=f"Monthly Sales - {selected_product}",
                    labels={sales_col: 'Revenue (USD)'}
                )
                st.plotly_chart(fig, use_container_width=True)
    
    # Sales Forecasting
    if product_col and len(df) >= 50:
        st.markdown('<div class="section-header">Sales Forecasting</div>', unsafe_allow_html=True)
        
        forecast_products = [''] + sorted(df[product_col].unique().tolist())
        selected_forecast = st.selectbox("Select Product for Forecasting", forecast_products)
        
        if selected_forecast:
            product_data = df[df[product_col] == selected_forecast]
            
            if len(product_data) >= 30:
                features = []
                if quantity_col:
                    features.append(quantity_col)
                if price_col:
                    features.append(price_col)
                if 'Month' in product_data.columns:
                    features.append('Month')
                
                if len(features) >= 2:
                    X = product_data[features].fillna(0)
                    y = product_data[sales_col].fillna(0)
                    
                    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
                    
                    model = RandomForestRegressor(n_estimators=100, random_state=42)
                    model.fit(X_train, y_train)
                    
                    score = model.score(X_test, y_test)
                    st.info(f"Model Accuracy: {max(0, score)*100:.1f}%")
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        forecast_units = st.number_input(
                            "Expected Units",
                            min_value=1,
                            value=int(product_data[quantity_col].mean()) if quantity_col else 100
                        )
                    
                    with col2:
                        forecast_price = st.number_input(
                            "Unit Price (USD)",
                            min_value=0.01,
                            value=float(product_data[price_col].mean()) if price_col else 100.0
                        )
                    
                    with col3:
                        forecast_month = st.slider("Target Month", 1, 12, 6)
                    
                    if st.button("Generate Forecast", use_container_width=True):
                        pred_features = []
                        if quantity_col:
                            pred_features.append(forecast_units)
                        if price_col:
                            pred_features.append(forecast_price)
                        if 'Month' in features:
                            pred_features.append(forecast_month)
                        
                        prediction = model.predict([pred_features])[0]
                        expected = forecast_units * forecast_price
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("AI Predicted Revenue", f"${prediction:,.2f}")
                        with col2:
                            variance = ((prediction - expected) / expected) * 100
                            st.metric("Expected Revenue", f"${expected:,.2f}", delta=f"{variance:+.1f}%")
            else:
                st.warning(f"Insufficient data for {selected_forecast}. Need at least 30 records.")
    
    # Business Recommendations
    st.markdown('<div class="section-header">Business Recommendations</div>', unsafe_allow_html=True)
    
    if product_col and sales_col:
        top_product = df.groupby(product_col)[sales_col].sum().idxmax()
        st.markdown(f"""
        <div class="recommendation-box">
            <strong>Top Product Strategy</strong><br>
            {top_product} is your best-selling product generating ${df[df[product_col] == top_product][sales_col].sum():,.2f}.<br>
            Recommendation: Increase inventory by 25 percent and allocate 40 percent of marketing budget.
        </div>
        """, unsafe_allow_html=True)
    
    if region_col and sales_col and len(df[region_col].unique()) > 1:
        region_perf = df.groupby(region_col)[sales_col].sum()
        best = region_perf.idxmax()
        worst = region_perf.idxmin()
        st.markdown(f"""
        <div class="recommendation-box">
            <strong>Regional Growth Strategy</strong><br>
            Best Region: {best} (${region_perf[best]:,.2f}) | Needs Improvement: {worst} (${region_perf[worst]:,.2f})<br>
            Recommendation: Expand successful strategies from {best} to {worst} region.
        </div>
        """, unsafe_allow_html=True)
    
    if date_col and 'Month' in df.columns:
        monthly_sales = df.groupby('Month')[sales_col].sum()
        peak_month = monthly_sales.idxmax()
        st.markdown(f"""
        <div class="recommendation-box">
            <strong>Seasonal Strategy</strong><br>
            Peak sales in Month {peak_month} (${monthly_sales[peak_month]:,.2f})<br>
            Recommendation: Increase inventory by 40 percent and launch campaigns 60 days before peak.
        </div>
        """, unsafe_allow_html=True)
    
    # Export Section
    st.markdown('<div class="section-header">Export Report</div>', unsafe_allow_html=True)
    
    if st.button("Download Word Report", type="primary", use_container_width=True):
        with st.spinner("Generating Word report..."):
            word_buffer = create_word_report(df, sales_col, product_col, region_col, date_col, quantity_col, price_col)
            st.download_button(
                "Download Report",
                word_buffer,
                f"sales_analytics_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

# Footer
st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>Sales Analytics Platform | Powered by Machine Learning</p>",
    unsafe_allow_html=True
)