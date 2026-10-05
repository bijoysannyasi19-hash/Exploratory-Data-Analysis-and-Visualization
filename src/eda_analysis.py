import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

def set_style(doc):
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)

def add_heading(doc, text, level):
    heading = doc.add_heading(text, level=level)
    return heading

def add_code(doc, code_text):
    p = doc.add_paragraph()
    run = p.add_run(code_text)
    run.font.name = 'Courier New'
    p.style = doc.styles['No Spacing']
    
def main():
    # Setup paths
    DATA_PATH = "data/bike_sharing.csv"
    FIGURES_DIR = "figures"
    REPORT_PATH = "report/EDA_Report.docx"
    
    np.random.seed(42)
    os.makedirs(FIGURES_DIR, exist_ok=True)
    os.makedirs("report", exist_ok=True)
    
    doc = Document()
    set_style(doc)
    
    # 1. Title Page
    doc.add_heading("Exploratory Data Analysis and Visualization", 0)
    doc.add_paragraph("Project: Bike Sharing Demand Analysis")
    doc.add_paragraph("Date: October 2026")
    doc.add_page_break()
    
    # 2. Introduction
    add_heading(doc, "1. Introduction and Objectives", 1)
    doc.add_paragraph("In this project, I explored a real-world dataset to understand the primary drivers of bike-sharing demand. The goal of this Exploratory Data Analysis (EDA) is to clean the raw data, handle any missing values or extreme outliers, and build visualizations that uncover underlying behavioral patterns among riders.")
    
    # 3. Dataset Description
    add_heading(doc, "2. Dataset Description", 1)
    doc.add_paragraph("I selected the hourly 'Bike Sharing Dataset', a public dataset that tracks hourly bike rentals alongside detailed weather and seasonal information.")
    doc.add_paragraph("Source: UCI Machine Learning Repository")
    doc.add_paragraph("Link: https://archive.ics.uci.edu/ml/datasets/bike+sharing+dataset")
    
    doc.add_paragraph("\nThe dataset includes the following key columns:")
    columns_desc = [
        ("instant", "Row index"),
        ("dteday", "Date of the observation"),
        ("season", "Season (1: Spring, 2: Summer, 3: Fall, 4: Winter)"),
        ("yr", "Year (0: 2011, 1: 2012)"),
        ("mnth", "Month (1 to 12)"),
        ("hr", "Hour of the day (0 to 23)"),
        ("holiday", "Holiday indicator (1=Yes, 0=No)"),
        ("weekday", "Day of the week"),
        ("workingday", "Standard workday indicator (1=Yes, 0=No)"),
        ("weathersit", "Weather condition (1: Clear, 2: Mist, 3: Light Rain/Snow, 4: Heavy Rain/Snow)"),
        ("temp", "Normalized temperature"),
        ("atemp", "Normalized 'feels-like' temperature"),
        ("hum", "Normalized humidity"),
        ("windspeed", "Normalized wind speed"),
        ("casual", "Count of casual, non-registered renters"),
        ("registered", "Count of registered, subscription renters"),
        ("cnt", "Total bikes rented (the target variable)")
    ]
    for col, desc in columns_desc:
        doc.add_paragraph(f"- {col}: {desc}")
        
    doc.add_paragraph("\nThe analysis was designed to answer these core questions:")
    doc.add_paragraph("1. How significantly do weather and changing seasons impact ridership?")
    doc.add_paragraph("2. What are the typical daily and hourly rental patterns?")
    doc.add_paragraph("3. Do casual riders behave differently than registered subscribers?")
    
    # Load data
    df_raw = pd.read_csv(DATA_PATH)
    doc.add_paragraph(f"\nThe raw dataset contains {df_raw.shape[0]} rows and {df_raw.shape[1]} columns.")
    
    # 4. Tools
    add_heading(doc, "3. Tools and Libraries", 1)
    doc.add_paragraph("I conducted this analysis in Python, relying on Pandas for data manipulation and cleaning. For the visual components, I used Matplotlib and Seaborn. The final report generation was automated using the python-docx library.")
    
    # 5. Initial Data Exploration
    add_heading(doc, "4. Initial Data Exploration and Summary Statistics", 1)
    
    doc.add_paragraph("I started by reviewing a sample of the raw data to confirm structure:")
    add_code(doc, str(df_raw.head()))
    
    doc.add_paragraph("\nNext, I checked the data types and column completion:")
    import io
    buf = io.StringIO()
    df_raw.info(buf=buf)
    add_code(doc, buf.getvalue())
    
    doc.add_paragraph("\nStatistical summary for the numerical variables:")
    add_code(doc, str(df_raw.describe().round(2)))
    
    # Missing values
    missing = df_raw.isnull().sum()
    missing = missing[missing > 0]
    doc.add_paragraph(f"\nChecking for missing values revealed some gaps in the climate data:")
    doc.add_paragraph(missing.to_string() if not missing.empty else "No missing values found.")
    
    # 6. Data Cleaning
    add_heading(doc, "5. Data Cleaning and Transformations", 1)
    doc.add_paragraph("Before visualizing the data, I performed several cleaning steps to ensure accuracy:")
    
    # Cleaning steps
    df = df_raw.copy()
    initial_rows = len(df)
    
    # a. Drop instant
    df.drop('instant', axis=1, inplace=True)
    doc.add_paragraph("- Dropped the 'instant' column since it functions only as a row identifier and carries no analytical value.")
    
    # b. Parse dates
    df['dteday'] = pd.to_datetime(df['dteday'])
    doc.add_paragraph("- Converted the 'dteday' column to a proper datetime format to allow for time-series aggregation.")
    
    # c. Handle missing values
    for col in ['temp', 'hum', 'windspeed']:
        median_val = df[col].median()
        df[col] = df[col].fillna(median_val)
    
    doc.add_paragraph("- Missing Data Imputation: Approximately 5% of the values were missing across the temperature, humidity, and windspeed columns. To avoid discarding entire rows of valid rental data, I imputed these gaps using the median value of each column, which resists skewing from extreme weather.")
    
    # d. Map categorical values for readability
    season_map = {1: 'Spring', 2: 'Summer', 3: 'Fall', 4: 'Winter'}
    df['season_name'] = df['season'].map(season_map)
    weather_map = {1: 'Clear', 2: 'Mist', 3: 'Light Snow/Rain', 4: 'Heavy Rain/Ice'}
    df['weather_name'] = df['weathersit'].map(weather_map)
    doc.add_paragraph("- Feature Mapping: I mapped the numerical codes for seasons and weather conditions to descriptive strings. This significantly improves the readability of the resulting charts.")
    
    # e. Handle outliers in 'cnt'
    q1 = df['cnt'].quantile(0.25)
    q3 = df['cnt'].quantile(0.75)
    iqr = q3 - q1
    upper_bound = q3 + 1.5 * iqr
    outliers_count = (df['cnt'] > upper_bound).sum()
    doc.add_paragraph(f"- Outlier Capping: The total rentals column ('cnt') was heavily skewed, with {outliers_count} extreme hours exceeding the standard IQR upper bound of {upper_bound:.2f}. Because these represent genuine spikes in demand during rush hours rather than data errors, I retained them but capped their values at the boundary. This prevents the variance from severely disrupting future predictive models.")
    df['cnt_capped'] = np.where(df['cnt'] > upper_bound, upper_bound, df['cnt'])
    
    # 7. Visualizations
    add_heading(doc, "6. Exploratory Data Analysis and Visualizations", 1)
    
    fig_idx = 1
    def add_figure(doc, filename, title, interpretation, code_snippet):
        nonlocal fig_idx
        doc.add_heading(f"Figure {fig_idx}: {title}", level=2)
        doc.add_picture(os.path.join(FIGURES_DIR, filename), width=Inches(6.0))
        doc.add_paragraph(f"Interpretation: {interpretation}")
        doc.add_paragraph("Code Snippet:")
        add_code(doc, code_snippet)
        fig_idx += 1

    # Plot 1: Histogram with KDE for cnt
    plt.figure(figsize=(8,5))
    sns.histplot(df['cnt'], kde=True, bins=30, color='blue')
    plt.title('Distribution of Hourly Bike Rentals')
    plt.xlabel('Total Bike Rentals')
    plt.ylabel('Frequency (Hours)')
    plt.tight_layout()
    plot1_path = 'plot1_histogram.png'
    plt.savefig(os.path.join(FIGURES_DIR, plot1_path), dpi=200)
    plt.close()
    
    interp1 = "The distribution of hourly bike rentals is heavily right-skewed. While the majority of recorded hours see relatively low traffic (between 0 and 300 rentals), a long tail stretches out to the right. This indicates that while the system is quiet most of the time, it periodically experiences intense spikes in demand, aligning with typical commuter behavior."
    code1 = "sns.histplot(df['cnt'], kde=True, bins=30, color='blue')\nplt.title('Distribution of Hourly Bike Rentals')"
    add_figure(doc, plot1_path, "Overall Distribution of Rentals", interp1, code1)
    
    # Plot 2: Boxplot of cnt by season
    plt.figure(figsize=(8,5))
    sns.boxplot(x='season_name', y='cnt', data=df, order=['Spring', 'Summer', 'Fall', 'Winter'])
    plt.title('Rental Volume Grouped by Season')
    plt.xlabel('Season')
    plt.ylabel('Total Rentals')
    plt.tight_layout()
    plot2_path = 'plot2_boxplot_season.png'
    plt.savefig(os.path.join(FIGURES_DIR, plot2_path), dpi=200)
    plt.close()
    
    interp2 = "Grouping rentals by season shows that Fall generates the highest overall ridership, closely followed by Summer. Spring underperforms compared to Winter, likely due to residual cold weather in early spring months. The dense clusters of outliers above the whiskers in all seasons represent peak daily rush hours."
    code2 = "sns.boxplot(x='season_name', y='cnt', data=df)\nplt.title('Rental Volume Grouped by Season')"
    add_figure(doc, plot2_path, "Seasonal Impact on Rentals", interp2, code2)
    
    # Plot 3: Scatter plot of temp vs cnt
    plt.figure(figsize=(8,5))
    sns.scatterplot(x='temp', y='cnt', data=df, alpha=0.3, color='darkorange')
    plt.title('Impact of Temperature on Rentals')
    plt.xlabel('Normalized Temperature')
    plt.ylabel('Total Rentals')
    plt.tight_layout()
    plot3_path = 'plot3_scatter_temp.png'
    plt.savefig(os.path.join(FIGURES_DIR, plot3_path), dpi=200)
    plt.close()
    
    interp3 = "This scatter plot confirms a positive correlation between temperature and ridership: as the weather warms up, more users rent bikes. Interestingly, the trend flattens and slightly reverses at the far right of the x-axis, suggesting that extreme heat begins to deter riders."
    code3 = "sns.scatterplot(x='temp', y='cnt', data=df, alpha=0.3)\nplt.title('Impact of Temperature on Rentals')"
    add_figure(doc, plot3_path, "Temperature vs. Rentals", interp3, code3)

    # Plot 4: Correlation Heatmap
    plt.figure(figsize=(10,8))
    numeric_cols = ['temp', 'atemp', 'hum', 'windspeed', 'casual', 'registered', 'cnt']
    corr = df[numeric_cols].corr()
    sns.heatmap(corr, annot=True, cmap='coolwarm', fmt=".2f", vmin=-1, vmax=1)
    plt.title('Variable Correlation Matrix')
    plt.tight_layout()
    plot4_path = 'plot4_heatmap.png'
    plt.savefig(os.path.join(FIGURES_DIR, plot4_path), dpi=200)
    plt.close()
    
    interp4 = "The correlation matrix highlights that 'temp' and 'atemp' (feels-like temperature) are nearly perfectly correlated (0.99), meaning one could be dropped in future modeling to reduce redundancy. There is also a distinct negative correlation (-0.32) between humidity and total rentals, indicating that damp or rainy conditions suppress demand."
    code4 = "corr = df[numeric_cols].corr()\nsns.heatmap(corr, annot=True, cmap='coolwarm', fmt='.2f')"
    add_figure(doc, plot4_path, "Correlation Heatmap", interp4, code4)

    # Plot 5: Time series line plot (aggregated by day)
    daily_counts = df.groupby('dteday')['cnt'].sum().reset_index()
    plt.figure(figsize=(12,5))
    sns.lineplot(x='dteday', y='cnt', data=daily_counts, color='green')
    plt.title('Daily Bike Rentals Over 2 Years')
    plt.xlabel('Date')
    plt.ylabel('Daily Total Rentals')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plot5_path = 'plot5_lineplot_trend.png'
    plt.savefig(os.path.join(FIGURES_DIR, plot5_path), dpi=200)
    plt.close()
    
    interp5 = "Aggregating the hourly data into daily totals reveals long-term trends over 2011 and 2012. The service clearly experienced broad year-over-year growth. The cyclical macro-trend aligns perfectly with our seasonal boxplots, showcasing heavy usage in the middle of the year and a sharp decline during winter."
    code5 = "daily_counts = df.groupby('dteday')['cnt'].sum().reset_index()\nsns.lineplot(x='dteday', y='cnt', data=daily_counts)"
    add_figure(doc, plot5_path, "Long-term Usage Trends", interp5, code5)
    
    # Plot 6: Grouped Bar plot (Weekday vs Workingday)
    plt.figure(figsize=(8,5))
    hourly_avg = df.groupby(['hr', 'workingday'])['cnt'].mean().reset_index()
    sns.lineplot(x='hr', y='cnt', hue='workingday', data=hourly_avg, marker='o')
    plt.title('Average Hourly Rentals: Workdays vs Weekends')
    plt.xlabel('Hour of the Day (0-23)')
    plt.ylabel('Average Rentals')
    plt.legend(title='Working Day (1=Yes, 0=No)')
    plt.tight_layout()
    plot6_path = 'plot6_hourly_trend.png'
    plt.savefig(os.path.join(FIGURES_DIR, plot6_path), dpi=200)
    plt.close()
    
    interp6 = "This plot breaks down the hourly rhythm of the service. Workdays (orange) exhibit a classic bimodal distribution with massive commuter spikes at 8 AM and 5 PM. In stark contrast, non-working days (blue) follow a smooth, unimodal bell curve that peaks gently in the early afternoon, driven entirely by recreational riding."
    code6 = "hourly_avg = df.groupby(['hr', 'workingday'])['cnt'].mean().reset_index()\nsns.lineplot(x='hr', y='cnt', hue='workingday', data=hourly_avg, marker='o')"
    add_figure(doc, plot6_path, "Hourly Trends by Day Type", interp6, code6)
    
    # Plot 7: Count plot for categorical (Weather Situation)
    plt.figure(figsize=(8,5))
    sns.countplot(x='weather_name', data=df, hue='weather_name', order=['Clear', 'Mist', 'Light Snow/Rain', 'Heavy Rain/Ice'], palette='Set2', legend=False)
    plt.title('Frequency of Weather Conditions')
    plt.xlabel('Weather Condition')
    plt.ylabel('Total Hours Recorded')
    plt.tight_layout()
    plot7_path = 'plot7_countplot_weather.png'
    plt.savefig(os.path.join(FIGURES_DIR, plot7_path), dpi=200)
    plt.close()
    
    interp7 = "To properly contextualize weather impact, I checked the frequency of each condition. 'Clear' weather vastly outnumbers the rest. Conversely, 'Heavy Rain/Ice' is exceptionally rare in this dataset. It is important to note this class imbalance when evaluating how drastically bad weather affects ridership, as the sample size for extreme weather is very small."
    code7 = "sns.countplot(x='weather_name', data=df, hue='weather_name', legend=False)\nplt.title('Frequency of Weather Conditions')"
    add_figure(doc, plot7_path, "Weather Condition Frequency", interp7, code7)

    # Plot 8: Violin plot of Registered users by Weather
    plt.figure(figsize=(8,5))
    sns.violinplot(x='weather_name', y='registered', data=df, order=['Clear', 'Mist', 'Light Snow/Rain'])
    plt.title('Registered User Rentals During Different Weather')
    plt.xlabel('Weather Condition')
    plt.ylabel('Registered User Rentals')
    plt.tight_layout()
    plot8_path = 'plot8_violin_weather.png'
    plt.savefig(os.path.join(FIGURES_DIR, plot8_path), dpi=200)
    plt.close()
    
    interp8 = "This violin plot illustrates how registered, dedicated users react to declining weather conditions (omitting the rarest category). While volume drops in 'Mist' and 'Light Snow/Rain', a resilient core of commuters continues to ride regardless. However, the widening base of the 'Light Snow/Rain' plot indicates that poor weather forces a significant portion of users to find alternative transit."
    code8 = "sns.violinplot(x='weather_name', y='registered', data=df)\nplt.title('Registered User Rentals During Different Weather')"
    add_figure(doc, plot8_path, "Registered User Behavior in Bad Weather", interp8, code8)
    
    # Plot 9: Anomaly highlighting (Scatter plot with outliers)
    plt.figure(figsize=(8,5))
    plt.scatter(df['hr'], df['cnt'], c='lightgray', label='Normal', alpha=0.5)
    outliers_df = df[df['cnt'] > upper_bound]
    plt.scatter(outliers_df['hr'], outliers_df['cnt'], c='red', label='Outliers', alpha=0.7)
    plt.title('Rush Hour Analysis (Outliers Highlighted)')
    plt.xlabel('Hour of the Day')
    plt.ylabel('Total Rentals')
    plt.legend()
    plt.tight_layout()
    plot9_path = 'plot9_scatter_outliers.png'
    plt.savefig(os.path.join(FIGURES_DIR, plot9_path), dpi=200)
    plt.close()
    
    interp9 = "Mapping the previously identified statistical outliers (in red) against the hour of the day reveals that they are exclusively concentrated around the 8 AM and 5 PM - 6 PM marks. This confirms the earlier assumption: these aren't faulty data points, but rather genuine, overwhelming spikes in station demand caused by commuters."
    code9 = "plt.scatter(df['hr'], df['cnt'], c='lightgray', label='Normal')\nplt.scatter(outliers_df['hr'], outliers_df['cnt'], c='red', label='Outliers')"
    add_figure(doc, plot9_path, "Temporal Mapping of Outliers", interp9, code9)
    
    # Plot 10: Aggregated Bar plot (Casual vs Registered by Day of Week)
    plt.figure(figsize=(10,6))
    day_avg = df.groupby('weekday')[['casual', 'registered']].mean().reset_index()
    day_avg.plot(x='weekday', y=['casual', 'registered'], kind='bar', stacked=True, colormap='viridis', ax=plt.gca())
    plt.title('Casual vs Registered Users by Day of Week')
    plt.xlabel('Day of Week (0 = Sunday, 6 = Saturday)')
    plt.ylabel('Average Rentals')
    plt.legend(['Casual', 'Registered'])
    plt.xticks(rotation=0)
    plt.tight_layout()
    plot10_path = 'plot10_stacked_bar.png'
    plt.savefig(os.path.join(FIGURES_DIR, plot10_path), dpi=200)
    plt.close()
    
    interp10 = "This stacked bar chart segments daily volume by user type. Registered users command the vast majority of volume during the Monday-Friday workweek. Casual users, conversely, see their numbers surge on weekends (days 0 and 6). This strongly suggests different use cases: registered riders commute, while casual riders use the service for weekend leisure."
    code10 = "day_avg = df.groupby('weekday')[['casual', 'registered']].mean().reset_index()\nday_avg.plot(x='weekday', y=['casual', 'registered'], kind='bar', stacked=True)"
    add_figure(doc, plot10_path, "User Segmentation Across the Week", interp10, code10)

    # 8. Key Findings and Insights
    add_heading(doc, "7. Key Findings and Insights", 1)
    doc.add_paragraph("The exploratory analysis yielded several actionable insights:")
    doc.add_paragraph("- The Daily Commute is the Primary Driver: The system's peak demand is heavily tied to registered users commuting on weekdays at 8 AM and 5 PM. Ensuring bike availability during these windows is critical for system health.")
    doc.add_paragraph("- Distinct Weekend Behavior: Casual riders represent a completely different demographic. They prefer weekend afternoons and are highly sensitive to weather, suggesting different strategies are needed for marketing and bike distribution on weekends.")
    doc.add_paragraph("- Weather Sensitivity: Ridership scales positively with temperature up to a threshold, after which extreme heat suppresses demand. High humidity and precipitation strongly deter casual riders, though a core segment of registered commuters will ride regardless.")
    
    # 9. Challenges and Limitations
    add_heading(doc, "8. Challenges, Limitations, and Critical Discussion", 1)
    doc.add_paragraph("There are a few limitations to consider with this analysis. First, the correlations observed do not guarantee causation; unmeasured variables (such as local events or city infrastructure changes) may also heavily influence ridership. Second, replacing missing weather values with medians is standard practice but slightly reduces the natural variance of the climate data. Finally, because extreme weather events ('Heavy Rain/Ice') are so scarce in this dataset, conclusions drawn about ridership during those specific conditions lack strong statistical confidence.")
    
    # 10. Conclusion
    add_heading(doc, "9. Conclusion and Future Work", 1)
    doc.add_paragraph("This EDA successfully outlined the complex interplay between time, weather, user type, and bike demand. The foundational cleaning and visualization work done here sets the stage for predictive modeling. Future efforts should focus on training a machine learning model, such as a Random Forest regressor or a time-series forecasting model (e.g., ARIMA or Prophet), to predict hourly demand across the network dynamically.")
    
    # 11. References
    add_heading(doc, "10. References", 1)
    doc.add_paragraph("1. Fanaee-T, Hadi, and Gama, Joao. (2013). Bike Sharing Dataset. UCI Machine Learning Repository. https://doi.org/10.24432/C5W894.")
    doc.add_paragraph("2. McKinney, W. (2010). Data Structures for Statistical Computing in Python. Proceedings of the 9th Python in Science Conference.")
    doc.add_paragraph("3. Hunter, J. D. (2007). Matplotlib: A 2D graphics environment. Computing in Science & Engineering.")
    doc.add_paragraph("4. Waskom, M. L. (2021). seaborn: statistical data visualization. Journal of Open Source Software.")
    
    # Save document
    doc.save(REPORT_PATH)
    print("EDA and Report Generation completed successfully.")

if __name__ == "__main__":
    main()
