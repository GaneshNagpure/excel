# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio
# import matplotlib.pyplot as plt
# import io
# import base64

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (1).xlsx"
    
#     # Read the Excel file (data starts from row 4; header is at index 3)
#     df = pd.read_excel(file_path, header=3)
#     df.columns = df.columns.str.strip()  # Clean column names

#     # Replace "NIL" with 0 and fill missing values for numeric columns
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({
#         'Quantity': 0,
#         'Market Value (Rs. In Lakhs)': 0,
#         '% age to NAV': 0,
#         'ISIN': '',
#     }, inplace=True)
    
#     # Convert to numeric types where needed
#     df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
    
#     print("Columns in DataFrame:", df.columns.tolist())

#     # Get or create the Portfolio object
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     # Initialize individual category totals
#     equity_total = 0
#     debt_total = 0
#     money_market_total = 0
#     others_total = 0

#     # Variable to track the current section/category from the Excel file
#     current_category = None

#     # Create a dictionary to store the total investment per industry/industry rating
#     industry_investments = {}

#     # Create a list to store the instruments and their % age to NAV for top 5
#     instrument_nav_percentages = []

#     # Process each row in the DataFrame
#     for index, row in df.iterrows():
#         name = str(row.get('Name of the Instruments', '')).strip()
#         lower_name = name.lower()

#         # Detect section header rows and update current_category
#         if lower_name.startswith('equity'):
#             current_category = 'Equity'
#             continue  # Skip header row
#         elif lower_name.startswith('debt'):
#             current_category = 'Debt'
#             continue
#         elif lower_name.startswith('money market'):
#             current_category = 'Money Market'
#             continue
#         elif lower_name.startswith('other'):
#             current_category = 'Others'
#             continue

#         # Detect subtotal/total rows. For Money Market and Others, update the total if desired.
#         if lower_name.startswith('subtotal') or lower_name.startswith('total'):
#             if current_category == 'Money Market':
#                 money_market_total = row['Market Value (Rs. In Lakhs)']
#             elif current_category == 'Others':
#                 others_total = row['Market Value (Rs. In Lakhs)']
#             continue

#         # Skip rows that are not instrument data (empty name or missing ISIN)
#         if not name:
#             continue
#         isin = str(row.get('ISIN', '')).strip()
#         if not isin:
#             continue

#         # Determine instrument type; default to 'Others' if current_category is not set.
#         instrument_type = current_category if current_category is not None else 'Others'
#         market_value = row['Market Value (Rs. In Lakhs)']

#         # Use get_or_create() to prevent duplicate objects based on ISIN
#         instrument, created = Instrument.objects.get_or_create(
#             isin=isin,
#             defaults={
#                 'instrument_name': name,
#                 'industry_rating': row.get('Industry/Rating', ''),
#                 'quantity': row['Quantity'],
#                 'market_value': market_value,
#                 'percentage_to_nav': row['% age to NAV'],
#                 'yield_percentage': row.get('Yield %', None),
#                 'ytc': row.get('^YTC (AT1/Tier 2 bonds)', None),
#                 'instrument_type': instrument_type,
#             }
#         )

#         # If the instrument already exists, update its fields as needed
#         if not created:
#             instrument.instrument_name = name
#             instrument.industry_rating = row.get('Industry/Rating', '')
#             instrument.quantity = row['Quantity']
#             instrument.market_value = market_value
#             instrument.percentage_to_nav = row['% age to NAV']
#             instrument.yield_percentage = row.get('Yield %', None)
#             instrument.ytc = row.get('^YTC (AT1/Tier 2 bonds)', None)
#             instrument.instrument_type = instrument_type
#             instrument.save()
        
#         # Update totals based on the instrument type
#         if instrument_type == 'Equity':
#             equity_total += market_value
#         elif instrument_type == 'Debt':
#             debt_total += market_value
#         elif instrument_type == 'Money Market':
#             money_market_total += market_value
#         elif instrument_type == 'Others':
#             others_total += market_value
        
#         # Aggregate investment for industry/ratings
#         industry = row.get('Industry/Rating', '').strip()
#         if industry:
#             if industry not in industry_investments:
#                 industry_investments[industry] = 0
#             industry_investments[industry] += market_value
        
#         # Store the instrument and its % age to NAV for top 5 instruments
#         nav_percentage = row['% age to NAV']
#         instrument_nav_percentages.append((name, nav_percentage, market_value))

#     # Combine Money Market and Others into one category
#     combined_others_total = money_market_total + others_total

#     # Calculate the overall total market value
#     total_market_value = equity_total + debt_total + combined_others_total

#     # Save these totals to the Portfolio (for Equity, Debt, and Others combined)
#     portfolio.equity_total = equity_total
#     portfolio.debt_total = debt_total
#     portfolio.other_total = combined_others_total
#     portfolio.total_market_value = total_market_value
#     portfolio.save()

#     # Calculate percentage shares (avoiding division by zero)
#     total = total_market_value if total_market_value > 0 else 1
#     equity_percentage = (equity_total / total) * 100
#     debt_percentage = (debt_total / total) * 100
#     others_percentage = (combined_others_total / total) * 100

#     # Get top 5 industries/ratings by total market value
#     sorted_industries = sorted(industry_investments.items(), key=lambda x: x[1], reverse=True)[:5]
#     top_industries = [{"industry": industry, "investment": round(investment, 2)} for industry, investment in sorted_industries]

#     # Get top 5 instruments by % age to NAV
#     sorted_instruments = sorted(instrument_nav_percentages, key=lambda x: x[1], reverse=True)[:5]
#     top_instruments = [{"instrument_name": instrument, "nav_percentage": round(nav_percentage, 2)} for instrument, nav_percentage, _ in sorted_instruments]

#     # Plot pie chart for top 5 instruments
#     labels_instruments = [x['instrument_name'] for x in top_instruments]
#     values_instruments = [x['nav_percentage'] for x in top_instruments]

#     fig_instruments, ax_instruments = plt.subplots()
#     ax_instruments.pie(values_instruments, labels=labels_instruments, autopct='%1.1f%%', startangle=90)
#     ax_instruments.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

#     # Plot pie chart for top 5 industries
#     labels_industries = [x['industry'] for x in top_industries]
#     values_industries = [x['investment'] for x in top_industries]

#     fig_industries, ax_industries = plt.subplots()
#     ax_industries.pie(values_industries, labels=labels_industries, autopct='%1.1f%%', startangle=90)
#     ax_industries.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

#     # -------------------------------
#     # Updated horizontal bar chart using percentages
#     # -------------------------------
#     categories = ['Equity', 'Debt', 'Others']
#     percentage_values = [equity_percentage, debt_percentage, others_percentage]

#     fig_bar, ax_bar = plt.subplots()
#     ax_bar.barh(categories, percentage_values, color=['green', 'blue', 'orange'])
#     ax_bar.set_xlabel('Percentage (%)')
#     ax_bar.set_title('Portfolio Market Value Distribution by Category (in %)')

#     # Add percentage labels on the bars
#     for bar in ax_bar.patches:
#         width = bar.get_width()
#         ax_bar.text(width + 1, bar.get_y() + bar.get_height() / 2.0,
#                     f'{width:.2f}%', va='center', fontsize=10)

#     # Convert charts to base64 to display in a web page
#     buffer_instruments = io.BytesIO()
#     fig_instruments.savefig(buffer_instruments, format='png', bbox_inches="tight")
#     buffer_instruments.seek(0)
#     img_instruments = base64.b64encode(buffer_instruments.getvalue()).decode('utf-8')

#     buffer_industries = io.BytesIO()
#     fig_industries.savefig(buffer_industries, format='png', bbox_inches="tight")
#     buffer_industries.seek(0)
#     img_industries = base64.b64encode(buffer_industries.getvalue()).decode('utf-8')

#     buffer_bar = io.BytesIO()
#     fig_bar.savefig(buffer_bar, format='png', bbox_inches="tight")
#     buffer_bar.seek(0)
#     img_bar = base64.b64encode(buffer_bar.getvalue()).decode('utf-8')

#     # Return the results in a Django template
#     return render(request, 'upload_result.html', {
#         'equity_total': equity_total,
#         'debt_total': debt_total,
#         'others_total': combined_others_total,
#         'total_market_value': total_market_value,
#         'equity_percentage': round(equity_percentage, 2),
#         'debt_percentage': round(debt_percentage, 2),
#         'others_percentage': round(others_percentage, 2),
#         'top_industries': top_industries,
#         'top_instruments': top_instruments,
#         'img_instruments': img_instruments,
#         'img_industries': img_industries,
#         'img_bar': img_bar,  # Updated bar chart with percentages
#     })


# import matplotlib.pyplot as plt
# from io import BytesIO
# import base64
# from django.shortcuts import render
# from .models import Portfolio

# def bar_graph(request):
#     # Retrieve the latest portfolio record (you can adjust the filtering as needed)
#     portfolio = Portfolio.objects.latest('portfolio_date')
    
#     # Prepare data for the graph
#     total_value = portfolio.equity_total + portfolio.debt_total + portfolio.other_total  # Total market value
    
#     # Calculate percentages for each category
#     equity_percentage = (portfolio.equity_total / total_value) * 100
#     debt_percentage = (portfolio.debt_total / total_value) * 100
#     other_percentage = (portfolio.other_total / total_value) * 100
    
#     # Prepare the labels and values for the bar graph
#     labels = ['Equity', 'Debt', 'Combined Others']
#     values = [equity_percentage, debt_percentage, other_percentage]
    
#     # Create a horizontal bar graph using Matplotlib
#     plt.figure(figsize=(8, 6))
#     bars = plt.barh(labels, values, color=['blue', 'green', 'orange'])
#     plt.xlabel("Percentage (%)")
#     plt.ylabel("Categories")
#     plt.title("Portfolio Market Value Distribution by Category")
    
#     # Optionally, add the percentage value on the bars
#     for bar in bars:
#         width = bar.get_width()
#         plt.text(
#             width, 
#             bar.get_y() + bar.get_height() / 2.0, 
#             f'{width:.2f}%', 
#             ha='left', 
#             va='center'
#         )
    
#     # Save the figure to a BytesIO buffer
#     buffer = BytesIO()
#     plt.savefig(buffer, format='png')
#     buffer.seek(0)
#     image_png = buffer.getvalue()
#     buffer.close()
    
#     # Encode the image to base64 string
#     graph = base64.b64encode(image_png).decode('utf-8')
#     plt.close()
    
#     # Pass the graph to the template
#     return render(request, 'bar_graph.html', {'graph': graph})


# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio
# import matplotlib.pyplot as plt
# import io
# import base64

# def safe_strip(value):
#     """Convert to string and strip if it's not NaN."""
#     if isinstance(value, str):
#         return value.strip()
#     return str(value) if pd.notna(value) else ''  # Convert NaN to empty string

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (5).xlsx"
    
#     # Read the Excel file (data starts from row 4; header is at index 3)
#     df = pd.read_excel(file_path, header=3)
#     print(df.columns)
#     df.columns = df.columns.str.strip()  # Clean column names

#     # Replace "NIL" with 0 and fill missing values for numeric columns
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({
#         'Quantity': 0,
#         'Market Value (Rs. In Lakhs)': 0,
#         '% age to NAV': 0,
#         'ISIN': '',
#     }, inplace=True)
    
#     # Convert to numeric types where needed
#     df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
    
#     print("Columns in DataFrame:", df.columns.tolist())

#     # Get or create the Portfolio object
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     # Initialize individual category totals
#     equity_total = 0
#     debt_total = 0
#     money_market_total = 0
#     others_total = 0

#     # Variable to track the current section/category from the Excel file
#     current_category = None

#     # Create a dictionary to store the total investment per industry/industry rating
#     industry_investments = {}

#     # Create a list to store the instruments and their % age to NAV for top 5
#     instrument_nav_percentages = []

#     # Process each row in the DataFrame
#     for index, row in df.iterrows():
#         name = safe_strip(row.get('Name of the Instruments', ''))
#         lower_name = name.lower()

#         # Detect section header rows and update current_category
#         if lower_name.startswith('equity'):
#             current_category = 'Equity'
#             continue  # Skip header row
#         elif lower_name.startswith('debt'):
#             current_category = 'Debt'
#             continue
#         elif lower_name.startswith('money market'):
#             current_category = 'Money Market'
#             continue
#         elif lower_name.startswith('other'):
#             current_category = 'Others'
#             continue

#         # Detect subtotal/total rows. For Money Market and Others, update the total if desired.
#         if lower_name.startswith('subtotal') or lower_name.startswith('total'):
#             if current_category == 'Money Market':
#                 money_market_total = row['Market Value (Rs. In Lakhs)']
#             elif current_category == 'Others':
#                 others_total = row['Market Value (Rs. In Lakhs)']
#             continue

#         # Skip rows that are not instrument data (empty name or missing ISIN)
#         if not name:
#             continue
#         isin = safe_strip(row.get('ISIN', ''))
#         if not isin:
#             continue

#         # Determine instrument type; default to 'Others' if current_category is not set.
#         instrument_type = current_category if current_category is not None else 'Others'
#         market_value = row['Market Value (Rs. In Lakhs)']

#         # Use get_or_create() to prevent duplicate objects based on ISIN
#         instrument, created = Instrument.objects.get_or_create(
#             isin=isin,
#             defaults={
#                 'instrument_name': name,
#                 'industry_rating': safe_strip(row.get('Industry/Rating', '')),
#                 'quantity': row['Quantity'],
#                 'market_value': market_value,
#                 'percentage_to_nav': row['% age to NAV'],
#                 'yield_percentage': row.get('Yield %', None),
#                 'ytc': row.get('^YTC (AT1/Tier 2 bonds)', None),
#                 'instrument_type': instrument_type,
#             }
#         )

#         # If the instrument already exists, update its fields as needed
#         if not created:
#             instrument.instrument_name = name
#             instrument.industry_rating = safe_strip(row.get('Industry/Rating', ''))
#             instrument.quantity = row['Quantity']
#             instrument.market_value = market_value
#             instrument.percentage_to_nav = row['% age to NAV']
#             instrument.yield_percentage = row.get('Yield %', None)
#             instrument.ytc = row.get('^YTC (AT1/Tier 2 bonds)', None)
#             instrument.instrument_type = instrument_type
#             instrument.save()
        
#         # Update totals based on the instrument type
#         if instrument_type == 'Equity':
#             equity_total += market_value
#         elif instrument_type == 'Debt':
#             debt_total += market_value
#         elif instrument_type == 'Money Market':
#             money_market_total += market_value
#         elif instrument_type == 'Others':
#             others_total += market_value
        
#         # Aggregate investment for industry/ratings
#         industry = safe_strip(row.get('Industry/Rating', ''))
#         if industry:
#             if industry not in industry_investments:
#                 industry_investments[industry] = 0
#             industry_investments[industry] += market_value
        
#         # Store the instrument and its % age to NAV for top 5 instruments
#         nav_percentage = row['% age to NAV']
#         instrument_nav_percentages.append((name, nav_percentage, market_value))

#     # Combine Money Market and Others into one category
#     combined_others_total = money_market_total + others_total

#     # Calculate the overall total market value
#     total_market_value = equity_total + debt_total + combined_others_total

#     # Save these totals to the Portfolio (for Equity, Debt, and Others combined)
#     portfolio.equity_total = equity_total
#     portfolio.debt_total = debt_total
#     portfolio.other_total = combined_others_total
#     portfolio.total_market_value = total_market_value
#     portfolio.save()

#     # Calculate percentage shares (avoiding division by zero)
#     total = total_market_value if total_market_value > 0 else 1
#     equity_percentage = (equity_total / total) * 100
#     debt_percentage = (debt_total / total) * 100
#     others_percentage = (combined_others_total / total) * 100

#     # Get top 5 industries/ratings by total market value
#     sorted_industries = sorted(industry_investments.items(), key=lambda x: x[1], reverse=True)[:5]
#     top_industries = [{"industry": industry, "investment": round(investment, 2)} for industry, investment in sorted_industries]

#     # Get top 5 instruments by % age to NAV
#     sorted_instruments = sorted(instrument_nav_percentages, key=lambda x: x[1], reverse=True)[:5]
#     top_instruments = [{"instrument_name": instrument, "nav_percentage": round(nav_percentage, 2)} for instrument, nav_percentage, _ in sorted_instruments]

#     # Plot pie chart for top 5 instruments
#     labels_instruments = [x['instrument_name'] for x in top_instruments]
#     values_instruments = [x['nav_percentage'] for x in top_instruments]

#     fig_instruments, ax_instruments = plt.subplots()
#     ax_instruments.pie(values_instruments, labels=labels_instruments, autopct='%1.1f%%', startangle=90)
#     ax_instruments.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

#     # Plot pie chart for top 5 industries
#     labels_industries = [x['industry'] for x in top_industries]
#     values_industries = [x['investment'] for x in top_industries]

#     fig_industries, ax_industries = plt.subplots()
#     ax_industries.pie(values_industries, labels=labels_industries, autopct='%1.1f%%', startangle=90)
#     ax_industries.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

#     # -------------------------------
#     # Updated horizontal bar chart using percentages
#     # -------------------------------
#     categories = ['Equity', 'Debt', 'Others']
#     percentage_values = [equity_percentage, debt_percentage, others_percentage]

#     fig_bar, ax_bar = plt.subplots()
#     ax_bar.barh(categories, percentage_values, color=['green', 'blue', 'orange'])
#     ax_bar.set_xlabel('Percentage (%)')
#     ax_bar.set_title('Portfolio Market Value Distribution by Category (in %)')

#     # Add percentage labels on the bars
#     for bar in ax_bar.patches:
#         width = bar.get_width()
#         ax_bar.text(width + 1, bar.get_y() + bar.get_height() / 2.0,
#                     f'{width:.2f}%', va='center', fontsize=10)

#     # Convert charts to base64 to display in a web page
#     buffer_instruments = io.BytesIO()
#     fig_instruments.savefig(buffer_instruments, format='png', bbox_inches="tight")
#     buffer_instruments.seek(0)
#     img_instruments = base64.b64encode(buffer_instruments.getvalue()).decode('utf-8')

#     buffer_industries = io.BytesIO()
#     fig_industries.savefig(buffer_industries, format='png', bbox_inches="tight")
#     buffer_industries.seek(0)
#     img_industries = base64.b64encode(buffer_industries.getvalue()).decode('utf-8')

#     buffer_bar = io.BytesIO()
#     fig_bar.savefig(buffer_bar, format='png', bbox_inches="tight")
#     buffer_bar.seek(0)
#     img_bar = base64.b64encode(buffer_bar.getvalue()).decode('utf-8')

#     # Return the results in a Django template
#     return render(request, 'upload_result.html', {
#         'equity_total': equity_total,
#         'debt_total': debt_total,
#         'others_total': combined_others_total,
#         'total_market_value': total_market_value,
#         'equity_percentage': round(equity_percentage, 2),
#         'debt_percentage': round(debt_percentage, 2),
#         'others_percentage': round(others_percentage, 2),
#         'top_industries': top_industries,
#         'top_instruments': top_instruments,
#         'img_instruments': img_instruments,
#         'img_industries': img_industries,
#         'img_bar': img_bar,  # Updated bar chart with percentages
#     })

# new new with handling error of quantity to quantity/face value because some files have quantity and some have quantity/face value code 
# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio
# import matplotlib.pyplot as plt
# import io
# import base64

# def safe_strip(value):
#     """Convert to string and strip if it's not NaN."""
#     if isinstance(value, str):
#         return value.strip()
#     return str(value) if pd.notna(value) else ''  # Convert NaN to empty string

# def upload_excel(request):
#     file_path ="c:/Users/91787/Downloads/Monthly Portfolio of Schemes (10).xlsx"
    
#     # Read the Excel file (data starts from row 4; header is at index 3)
#     df = pd.read_excel(file_path, header=3)
#     print(df.columns)
#     df.columns = df.columns.str.strip()  # Clean column names

#     # Replace "NIL" with 0 and fill missing values for numeric columns
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({
#         'Quantity': 0,
#         'Market Value (Rs. In Lakhs)': 0,
#         '% age to NAV': 0,
#         'ISIN': '',
#     }, inplace=True)
    
#     # Convert to numeric types where needed
#     # df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     if 'Quantity' in df.columns:
#         df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     elif 'Quantity/Face Value' in df.columns:
#         df['Quantity/Face Value'] = pd.to_numeric(df['Quantity/Face Value'], errors='coerce').fillna(0)
#         df.rename(columns={'Quantity/Face Value': 'Quantity'}, inplace=True)
    
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
    

#     # convert yield % column to folat after stripping %
#     if 'Yield %' in df.columns:
#         df['Yield %'] = df['Yield %'].astype(str).str.rstrip('%').replace('nan', None)
#         df['Yield %'] = pd.to_numeric(df['Yield %'], errors='coerce').fillna(0)

#     print("Columns in DataFrame:", df.columns.tolist())

#     # Get or create the Portfolio object
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     # Initialize individual category totals
#     equity_total = 0
#     debt_total = 0
#     money_market_total = 0
#     others_total = 0

#     # Variable to track the current section/category from the Excel file
#     current_category = None

#     # Create a dictionary to store the total investment per industry/industry rating
#     industry_investments = {}

#     # Create a list to store the instruments and their % age to NAV for top 5
#     instrument_nav_percentages = []

#     # Process each row in the DataFrame
#     for index, row in df.iterrows():
#         name = safe_strip(row.get('Name of the Instruments', ''))
#         lower_name = name.lower()

#         # Detect section header rows and update current_category
#         if lower_name.startswith('equity'):
#             current_category = 'Equity'
#             continue  # Skip header row
#         elif lower_name.startswith('debt'):
#             current_category = 'Debt'
#             continue
#         elif lower_name.startswith('money market'):
#             current_category = 'Money Market' 
#             continue
#         elif lower_name.startswith('others'):
#             current_category = 'Others'
#             continue
        
#         # Detect subtotal/total rows. For Money Market and Others, update the total if desired.
#         if lower_name.startswith('subtotal') or lower_name.startswith('total'):
#             if current_category == 'Money Market':
#                 money_market_total = row['Market Value (Rs. In Lakhs)']
#             elif current_category == 'Others':
#                 others_total = row['Market Value (Rs. In Lakhs)']
#             continue

#         # Skip rows that are not instrument data (empty name or missing ISIN)
#         if not name:
#             continue
#         isin = safe_strip(row.get('ISIN', ''))
#         if not isin:
#             continue

#         # Determine instrument type; default to 'Others' if current_category is not set.
#         instrument_type = current_category if current_category is not None else 'Others'
#         market_value = row['Market Value (Rs. In Lakhs)']

#         # Use get_or_create() to prevent duplicate objects based on ISIN
#         instrument, created = Instrument.objects.get_or_create(
#             isin=isin,
#             defaults={
#                 'instrument_name': name,
#                 'industry_rating': safe_strip(row.get('Industry/Rating', '')),
#                 'quantity': row['Quantity'],
#                 'market_value': market_value,
#                 'percentage_to_nav': row['% age to NAV'],
#                 'yield_percentage': row['yield %'] if 'yield %' in df.columns else None,
#                 'ytc': row.get('^YTC (AT1/Tier 2 bonds)', None),
#                 'instrument_type': instrument_type,
#             }
#         )

#         # If the instrument already exists, update its fields as needed
#         if not created:
#             instrument.instrument_name = name
#             instrument.industry_rating = safe_strip(row.get('Industry/Rating', ''))
#             instrument.quantity = row['Quantity']
#             instrument.market_value = market_value
#             instrument.percentage_to_nav = row['% age to NAV']
#             instrument.yield_percentage = row.get('Yield %', None)
#             instrument.ytc = row.get('^YTC (AT1/Tier 2 bonds)', None)
#             instrument.instrument_type = instrument_type
#             instrument.save()
        
#         # Update totals based on the instrument type
#         if instrument_type == 'Equity':
#             equity_total += market_value
#         elif instrument_type == 'Debt':
#             debt_total += market_value
#         elif instrument_type == 'Money Market':
#             money_market_total += market_value
#         elif instrument_type == 'Others':
#             others_total += market_value
        
#         # Aggregate investment for industry/ratings
#         industry = safe_strip(row.get('Industry/Rating', ''))
#         if industry:
#             if industry not in industry_investments:
#                 industry_investments[industry] = 0
#             industry_investments[industry] += market_value
        
#         # Store the instrument and its % age to NAV for top 5 instruments
#         nav_percentage = row['% age to NAV']
#         instrument_nav_percentages.append((name, nav_percentage, market_value))

#     # Combine Money Market and Others into one category
#     combined_others_total = money_market_total + others_total

#     print(money_market_total, others_total)
#     print(combined_others_total)
#     print("hello world")

#     # Calculate the overall total market value
#     total_market_value = equity_total + debt_total + combined_others_total

#     # Save these totals to the Portfolio (for Equity, Debt, and Others combined)
#     portfolio.equity_total = equity_total
#     portfolio.debt_total = debt_total
#     portfolio.other_total = combined_others_total
#     portfolio.total_market_value = total_market_value
#     portfolio.save()

#     # Calculate percentage shares (avoiding division by zero)
#     total = total_market_value if total_market_value > 0 else 1
#     equity_percentage = (equity_total / total) * 100
#     debt_percentage = (debt_total / total) * 100
#     others_percentage = (combined_others_total / total) * 100

#     # Get top 5 industries/ratings by total market value
#     sorted_industries = sorted(industry_investments.items(), key=lambda x: x[1], reverse=True)[:5]
#     top_industries = [{"industry": industry, "investment": round(investment, 2)} for industry, investment in sorted_industries]
#     print(top_industries)

#     # Get top 5 instruments by % age to NAV
#     sorted_instruments = sorted(instrument_nav_percentages, key=lambda x: x[1], reverse=True)[:5]
#     top_instruments = [{"instrument_name": instrument, "nav_percentage": round(nav_percentage, 2)} for instrument, nav_percentage, _ in sorted_instruments]
#     print(top_instruments)
#     # Plot pie chart for top 5 instruments
#     labels_instruments = [x['instrument_name'] for x in top_instruments]
#     values_instruments = [x['nav_percentage'] for x in top_instruments]

#     fig_instruments, ax_instruments = plt.subplots()
#     ax_instruments.pie(values_instruments, labels=labels_instruments, autopct='%1.1f%%', startangle=90)
#     ax_instruments.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

#     # Plot pie chart for top 5 industries
#     labels_industries = [x['industry'] for x in top_industries]
#     values_industries = [x['investment'] for x in top_industries]

#     fig_industries, ax_industries = plt.subplots()
#     ax_industries.pie(values_industries, labels=labels_industries, autopct='%1.1f%%', startangle=90)
#     ax_industries.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

#     # -------------------------------
#     # Updated horizontal bar chart using percentages
#     # -------------------------------
#     categories = ['Equity', 'Debt', 'Others']
#     percentage_values = [equity_percentage, debt_percentage, others_percentage]

#     fig_bar, ax_bar = plt.subplots()
#     ax_bar.barh(categories, percentage_values, color=['green', 'blue', 'orange'])
#     ax_bar.set_xlabel('Percentage (%)')
#     ax_bar.set_title('Portfolio Market Value Distribution by Category (in %)')

#     # Add percentage labels on the bars
#     for bar in ax_bar.patches:
#         width = bar.get_width()
#         ax_bar.text(width + 1, bar.get_y() + bar.get_height() / 2.0,
#                     f'{width:.2f}%', va='center', fontsize=10)

#     # Convert charts to base64 to display in a web page
#     buffer_instruments = io.BytesIO()
#     fig_instruments.savefig(buffer_instruments, format='png', bbox_inches="tight")
#     buffer_instruments.seek(0)
#     img_instruments = base64.b64encode(buffer_instruments.getvalue()).decode('utf-8')

#     buffer_industries = io.BytesIO()
#     fig_industries.savefig(buffer_industries, format='png', bbox_inches="tight")
#     buffer_industries.seek(0)
#     img_industries = base64.b64encode(buffer_industries.getvalue()).decode('utf-8')

#     buffer_bar = io.BytesIO()
#     fig_bar.savefig(buffer_bar, format='png', bbox_inches="tight")
#     buffer_bar.seek(0)
#     img_bar = base64.b64encode(buffer_bar.getvalue()).decode('utf-8')

#     # Return the results in a Django template
#     return render(request, 'upload_result.html', {
#         'equity_total': equity_total,
#         'debt_total': debt_total,
#         'others_total': combined_others_total,
#         'total_market_value': total_market_value,
#         'equity_percentage': round(equity_percentage, 2),
#         'debt_percentage': round(debt_percentage, 2),
#         'others_percentage': round(others_percentage, 2),
#         'top_industries': top_industries,
#         'top_instruments': top_instruments,
#         'img_instruments': img_instruments,
#         'img_industries': img_industries,
#         'img_bar': img_bar,  # Updated bar chart with percentages
#     })



# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio
# import matplotlib.pyplot as plt
# import io
# import base64

# def safe_strip(value):
#     """Convert to string and strip if it's not NaN."""
#     if isinstance(value, str):
#         return value.strip()
#     return str(value) if pd.notna(value) else ''  # Convert NaN to empty string

# def upload_excel(request):
#     file_path ="c:/Users/91787/Downloads/Monthly Portfolio of Schemes (2).xlsx"
    
#     # Read the Excel file (data starts from row 4; header is at index 3)
#     df = pd.read_excel(file_path, header=3)
#     print(df.columns)
#     df.columns = df.columns.str.strip()  # Clean column names

#     # Replace "NIL" with 0 and fill missing values for numeric columns
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({
#         'Quantity': 0,
#         'Market Value (Rs. In Lakhs)': 0,
#         '% age to NAV': 0,
#         'ISIN': '',
#     }, inplace=True)
    
#     # Convert to numeric types where needed
#     # df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     if 'Quantity' in df.columns:
#         df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     elif 'Quantity/Face Value' in df.columns:
#         df['Quantity/Face Value'] = pd.to_numeric(df['Quantity/Face Value'], errors='coerce').fillna(0)
#         df.rename(columns={'Quantity/Face Value': 'Quantity'}, inplace=True)
    
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
    

#     # convert yield % column to folat after stripping %
#     if 'Yield %' in df.columns:
#         df['Yield %'] = df['Yield %'].astype(str).str.rstrip('%').replace('nan', None)
#         df['Yield %'] = pd.to_numeric(df['Yield %'], errors='coerce').fillna(0)

#     print("Columns in DataFrame:", df.columns.tolist())

#     # Get or create the Portfolio object
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     # Initialize individual category totals
#     equity_total = 0
#     debt_total = 0
#     money_market_total = 0
#     alternative_investment_fund_units_total = 0
#     treps_reverse_repo_investments_total = 0
#     cash_and_cash_equivalents_total = 0
#     others_total = 0

#     # Variable to track the current section/category from the Excel file
#     current_category = None

#     # Create a dictionary to store the total investment per industry/industry rating
#     industry_investments = {}

#     # Create a list to store the instruments and their % age to NAV for top 5
#     instrument_nav_percentages = []

#     # Process each row in the DataFrame
#     for index, row in df.iterrows():
#         name = safe_strip(row.get('Name of the Instruments', ''))
#         lower_name = name.lower()

#         # Detect section header rows and update current_category
#         if lower_name.startswith('equity'):
#             current_category = 'Equity'
#             continue  # Skip header row
#         elif lower_name.startswith('debt'):
#             current_category = 'Debt'
#             continue
#         elif lower_name.startswith('money market'):
#             current_category = 'Money Market' 
#             continue
#         elif lower_name.startswith('others'):
#             current_category = 'Others'
#             continue
#         elif lower_name.startswith('alternative investment fund units'):
#             current_category = 'Alternative Investment Fund Units'
#             continue
#         elif lower_name.startswith('treps'):
#             current_category = 'TREPS / Reverse Repo Investments / Corporate Debt Repo'
#             print("hello world")
#             continue
#         elif lower_name.startswith('cash equivalents'):
#             current_category = 'Cash & Cash Equivalents'
#             continue


#         #Detect subtotal/total rows. For Money Market and Others, update the total if desired.
#         if lower_name.startswith('subtotal') or lower_name.startswith('total'):
#             if current_category == 'Money Market':
#                 money_market_total = row['Market Value (Rs. In Lakhs)']
#             elif current_category == 'Others':
#                 others_total = row['Market Value (Rs. In Lakhs)']
#             continue

#         # Skip rows that are not instrument data (empty name or missing ISIN)
#         if not name:
#             continue
#         isin = safe_strip(row.get('ISIN', ''))
#         if not isin:
#             continue

#         # Determine instrument type; default to 'Others' if current_category is not set.
#         instrument_type = current_category if current_category is not None else 'Others'
#         market_value = row['Market Value (Rs. In Lakhs)']

#         # Use get_or_create() to prevent duplicate objects based on ISIN
#         instrument, created = Instrument.objects.get_or_create(
#             isin=isin,
#             defaults={
#                 'instrument_name': name,
#                 'industry_rating': safe_strip(row.get('Industry/Rating', '')),
#                 'quantity': row['Quantity'],
#                 'market_value': market_value,
#                 'percentage_to_nav': row['% age to NAV'],
#                 'yield_percentage': row['yield %'] if 'yield %' in df.columns else None,
#                 'ytc': row.get('^YTC (AT1/Tier 2 bonds)', None),
#                 'instrument_type': instrument_type,
#             }
#         )

#         # If the instrument already exists, update its fields as needed
#         if not created:
#             instrument.instrument_name = name
#             instrument.industry_rating = safe_strip(row.get('Industry/Rating', ''))
#             instrument.quantity = row['Quantity']
#             instrument.market_value = market_value
#             instrument.percentage_to_nav = row['% age to NAV']
#             instrument.yield_percentage = row.get('Yield %', None)
#             instrument.ytc = row.get('^YTC (AT1/Tier 2 bonds)', None)
#             instrument.instrument_type = instrument_type
#             instrument.save()
        
#         # Update totals based on the instrument type
#         if instrument_type == 'Equity':
#             equity_total += market_value
#         elif instrument_type == 'Debt':
#             debt_total += market_value
#         elif instrument_type == 'Money Market':
#             money_market_total += market_value
#         elif instrument_type == 'Others':
#             others_total += market_value
#         elif instrument_type == 'Alternative Investment Fund Units':
#             alternative_investment_fund_units_total += market_value
#         elif instrument_type == 'TREPS / Reverse Repo Investments / Corporate Debt Repo':
#             treps_reverse_repo_investments_total += market_value
#         elif instrument_type == 'Cash & Cash Equivalents':
#             cash_and_cash_equivalents_total += market_value

        
#         # Aggregate investment for industry/ratings
#         industry = safe_strip(row.get('Industry/Rating', ''))
#         if industry:
#             if industry not in industry_investments:
#                 industry_investments[industry] = 0
#             industry_investments[industry] += market_value
        
#         # Store the instrument and its % age to NAV for top 5 instruments
#         nav_percentage = row['% age to NAV']
#         instrument_nav_percentages.append((name, nav_percentage, market_value))

#     # Combine Money Market and Others into one category
#     combined_others_total = money_market_total + others_total

#     print(money_market_total, others_total)
#     print(combined_others_total)
#     print(cash_and_cash_equivalents_total)
#     print(alternative_investment_fund_units_total)
#     print(treps_reverse_repo_investments_total)
#     print("hello world")

#     # Calculate the overall total market value
#     total_market_value = equity_total + debt_total + combined_others_total

#     # Save these totals to the Portfolio (for Equity, Debt, and Others combined)
#     portfolio.equity_total = equity_total
#     portfolio.debt_total = debt_total
#     portfolio.other_total = combined_others_total
#     portfolio.total_market_value = total_market_value
#     portfolio.save()

#     # Calculate percentage shares (avoiding division by zero)
#     total = total_market_value if total_market_value > 0 else 1
#     equity_percentage = (equity_total / total) * 100
#     debt_percentage = (debt_total / total) * 100
#     others_percentage = (combined_others_total / total) * 100

#     # Get top 5 industries/ratings by total market value
#     sorted_industries = sorted(industry_investments.items(), key=lambda x: x[1], reverse=True)[:5]
#     top_industries = [{"industry": industry, "investment": round(investment, 2)} for industry, investment in sorted_industries]
#     print(top_industries)

#     # Get top 5 instruments by % age to NAV
#     sorted_instruments = sorted(instrument_nav_percentages, key=lambda x: x[1], reverse=True)[:5]
#     top_instruments = [{"instrument_name": instrument, "nav_percentage": round(nav_percentage, 2)} for instrument, nav_percentage, _ in sorted_instruments]
#     print(top_instruments)
#     # Plot pie chart for top 5 instruments
#     labels_instruments = [x['instrument_name'] for x in top_instruments]
#     values_instruments = [x['nav_percentage'] for x in top_instruments]

#     fig_instruments, ax_instruments = plt.subplots()
#     ax_instruments.pie(values_instruments, labels=labels_instruments, autopct='%1.1f%%', startangle=90)
#     ax_instruments.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

#     # Plot pie chart for top 5 industries
#     labels_industries = [x['industry'] for x in top_industries]
#     values_industries = [x['investment'] for x in top_industries]

#     fig_industries, ax_industries = plt.subplots()
#     ax_industries.pie(values_industries, labels=labels_industries, autopct='%1.1f%%', startangle=90)
#     ax_industries.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

#     # -------------------------------
#     # Updated horizontal bar chart using percentages
#     # -------------------------------
#     categories = ['Equity', 'Debt', 'Others']
#     percentage_values = [equity_percentage, debt_percentage, others_percentage]

#     fig_bar, ax_bar = plt.subplots()
#     ax_bar.barh(categories, percentage_values, color=['green', 'blue', 'orange'])
#     ax_bar.set_xlabel('Percentage (%)')
#     ax_bar.set_title('Portfolio Market Value Distribution by Category (in %)')

#     # Add percentage labels on the bars
#     for bar in ax_bar.patches:
#         width = bar.get_width()
#         ax_bar.text(width + 1, bar.get_y() + bar.get_height() / 2.0,
#                     f'{width:.2f}%', va='center', fontsize=10)

#     # Convert charts to base64 to display in a web page
#     buffer_instruments = io.BytesIO()
#     fig_instruments.savefig(buffer_instruments, format='png', bbox_inches="tight")
#     buffer_instruments.seek(0)
#     img_instruments = base64.b64encode(buffer_instruments.getvalue()).decode('utf-8')

#     buffer_industries = io.BytesIO()
#     fig_industries.savefig(buffer_industries, format='png', bbox_inches="tight")
#     buffer_industries.seek(0)
#     img_industries = base64.b64encode(buffer_industries.getvalue()).decode('utf-8')

#     buffer_bar = io.BytesIO()
#     fig_bar.savefig(buffer_bar, format='png', bbox_inches="tight")
#     buffer_bar.seek(0)
#     img_bar = base64.b64encode(buffer_bar.getvalue()).decode('utf-8')

#     # Return the results in a Django template
#     return render(request, 'upload_result.html', {
#         'equity_total': equity_total,
#         'debt_total': debt_total,
#         'others_total': combined_others_total,
#         'total_market_value': total_market_value,
#         'equity_percentage': round(equity_percentage, 2),
#         'debt_percentage': round(debt_percentage, 2),
#         'others_percentage': round(others_percentage, 2),
#         'top_industries': top_industries,
#         'top_instruments': top_instruments,
#         'img_instruments': img_instruments,
#         'img_industries': img_industries,
#         'img_bar': img_bar,  # Updated bar chart with percentages
#     })

# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio
# import matplotlib.pyplot as plt
# import io
# import base64

# def safe_strip(value):
#     """Convert to string and strip if it's not NaN."""
#     if isinstance(value, str):
#         return value.strip()
#     return str(value) if pd.notna(value) else ''  # Convert NaN to empty string

# def upload_excel(request):
#     file_path ="c:/Users/91787/Downloads/Monthly Portfolio of Schemes (10).xlsx"
    
#     df = pd.read_excel(file_path, header=3)
#     print(df.columns)
#     df.columns = df.columns.str.strip()

#     df.replace("NIL", 0, inplace=True)
#     df.fillna({
#         'Quantity': 0,
#         'Market Value (Rs. In Lakhs)': 0,
#         '% age to NAV': 0,
#         'ISIN': '',
#     }, inplace=True)
    
#     if 'Quantity' in df.columns:
#         df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     elif 'Quantity/Face Value' in df.columns:
#         df['Quantity/Face Value'] = pd.to_numeric(df['Quantity/Face Value'], errors='coerce').fillna(0)
#         df.rename(columns={'Quantity/Face Value': 'Quantity'}, inplace=True)
    
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
    
#     if 'Yield %' in df.columns:
#         df['Yield %'] = df['Yield %'].astype(str).str.rstrip('%').replace('nan', None)
#         df['Yield %'] = pd.to_numeric(df['Yield %'], errors='coerce').fillna(0)

#     print("Columns in DataFrame:", df.columns.tolist())

#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     equity_total = 0
#     debt_total = 0
#     money_market_total = 0
#     others_total = 0

#     current_category = None
#     industry_investments = {}
#     instrument_nav_percentages = []

#     for index, row in df.iterrows():
#         name = safe_strip(row.get('Name of the Instruments', ''))
#         lower_name = name.lower()

#         if lower_name.startswith('equity'):
#             current_category = 'Equity'
#             continue
#         elif lower_name.startswith('debt'):
#             current_category = 'Debt'
#             continue
#         elif lower_name.startswith('money market'):
#             current_category = 'Money Market' 
#             continue
#         elif lower_name.startswith('others'):
#             current_category = 'Others'
#             continue
        
#         if lower_name.startswith('subtotal') or lower_name.startswith('total'):
#             if current_category == 'Money Market':
#                 money_market_total = row['Market Value (Rs. In Lakhs)']
#             elif current_category == 'Others':
#                 others_total = row['Market Value (Rs. In Lakhs)']
#             continue

#         if not name:
#             continue
#         isin = safe_strip(row.get('ISIN', ''))
#         if not isin:
#             continue

#         instrument_type = current_category if current_category is not None else 'Others'
#         market_value = row['Market Value (Rs. In Lakhs)']

#         instrument, created = Instrument.objects.get_or_create(
#             isin=isin,
#             defaults={
#                 'instrument_name': name,
#                 'industry_rating': safe_strip(row.get('Industry/Rating', '')),
#                 'quantity': row['Quantity'],
#                 'market_value': market_value,
#                 'percentage_to_nav': row['% age to NAV'],
#                 'yield_percentage': row['yield %'] if 'yield %' in df.columns else None,
#                 'ytc': row.get('^YTC (AT1/Tier 2 bonds)', None),
#                 'instrument_type': instrument_type,
#             }
#         )

#         if not created:
#             instrument.instrument_name = name
#             instrument.industry_rating = safe_strip(row.get('Industry/Rating', ''))
#             instrument.quantity = row['Quantity']
#             instrument.market_value = market_value
#             instrument.percentage_to_nav = row['% age to NAV']
#             instrument.yield_percentage = row.get('Yield %', None)
#             instrument.ytc = row.get('^YTC (AT1/Tier 2 bonds)', None)
#             instrument.instrument_type = instrument_type
#             instrument.save()
        
#         if instrument_type == 'Equity':
#             equity_total += market_value
#         elif instrument_type == 'Debt':
#             debt_total += market_value
#         elif instrument_type == 'Money Market':
#             money_market_total += market_value
#         elif instrument_type == 'Others':
#             others_total += market_value
        
#         industry = safe_strip(row.get('Industry/Rating', ''))
#         if industry:
#             industry_investments[industry] = industry_investments.get(industry, 0) + market_value
        
#         instrument_nav_percentages.append((name, row['% age to NAV'], market_value))
    
#     total_market_value = equity_total + debt_total + money_market_total + others_total
    
#     portfolio.equity_total = equity_total
#     portfolio.debt_total = debt_total
#     portfolio.other_total = money_market_total + others_total
#     portfolio.total_market_value = total_market_value
#     portfolio.save()
    
#     equity_percentage = (equity_total / total_market_value) * 100
#     debt_percentage = (debt_total / total_market_value) * 100
#     others_percentage = ((money_market_total + others_total) / total_market_value) * 100
    
#     print(f"Money Market Total: {money_market_total}")
#     print(f"Others Total: {others_total}")
#     print(f"Total Market Value: {total_market_value}")
    
#     return JsonResponse({
#         'equity_total': equity_total,
#         'debt_total': debt_total,
#         'money_market_total': money_market_total,
#         'others_total': others_total,
#         'total_market_value': total_market_value,
#         'equity_percentage': round(equity_percentage, 2),
#         'debt_percentage': round(debt_percentage, 2),
#         'others_percentage': round(others_percentage, 2),
#     })

# import pandas as pd
# import json
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Portfolio

# def safe_strip(value):
#     """Convert to string and strip if it's not NaN."""
#     if isinstance(value, str):
#         return value.strip()
#     return str(value) if pd.notna(value) else ''  # Convert NaN to empty string

# def upload_excel(request):
#     file_path = "c:/Users/91787/Downloads/Monthly Portfolio of Schemes (10).xlsx"
#     df = pd.read_excel(file_path, header=3)
#     df.columns = df.columns.str.strip()
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({
#         'Quantity': 0,
#         'Market Value (Rs. In Lakhs)': 0,
#         '% age to NAV': 0,
#         'ISIN': '',
#     }, inplace=True)
    
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
    
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     category_totals = {
#         'Equity': 0,
#         'Debt Instruments': 0,
#         'Money Market Instruments': 0,
#         'Alternative Investment Fund Units': 0,
#         'TREPS/Reverse Repo Investments/Corporate Debt Repo': 0,
#         'Cash and Cash Equivalents': 0,
#         'Others': 0,
#     }
    
#     current_category = None
    
#     for index, row in df.iterrows():
#         name = safe_strip(row.get('Name of the Instruments', ''))
#         lower_name = name.lower()
        
#         if any(keyword in lower_name for keyword in ['equity', 'shares', 'stocks']):
#             current_category = 'Equity'
#         elif any(keyword in lower_name for keyword in ['debt', 'bonds', 'ncd']):
#             current_category = 'Debt Instruments'
#         elif 'money market' in lower_name:
#             current_category = 'Money Market Instruments'
#         elif 'alternative investment fund' in lower_name:
#             current_category = 'Alternative Investment Fund Units'
#         elif any(keyword in lower_name for keyword in ['treps', 'reverse repo', 'corporate debt repo']):
#             current_category = 'TREPS/Reverse Repo Investments/Corporate Debt Repo'
#         elif any(keyword in lower_name for keyword in ['cash', 'cash equivalents']):
#             current_category = 'Cash and Cash Equivalents'
#         else:
#             current_category = 'Others'
        
#         if lower_name.startswith('subtotal') or lower_name.startswith('total'):
#             category_totals[current_category] += row['Market Value (Rs. In Lakhs)']
#             continue
    
#     total_market_value = sum(category_totals.values())
    
#     portfolio.equity_total = category_totals['Equity']
#     portfolio.debt_total = category_totals['Debt Instruments']
#     portfolio.other_total = category_totals['Others']
#     portfolio.total_market_value = total_market_value
#     portfolio.save()
    
#     category_totals['Total Market Value'] = total_market_value
    
#     return JsonResponse(category_totals)

# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio
# import matplotlib.pyplot as plt
# import io
# import base64

# def safe_strip(value):
#     """Convert to string and strip if it's not NaN."""
#     if isinstance(value, str):
#         return value.strip()
#     return str(value) if pd.notna(value) else ''  # Convert NaN to empty string

# def upload_excel(request):
#     file_path ="c:/Users/91787/Downloads/Monthly Portfolio of Schemes (10).xlsx"
    
#     # Read the Excel file (data starts from row 4; header is at index 3)
#     df = pd.read_excel(file_path, header=3)
#     print(df.columns)
#     df.columns = df.columns.str.strip()  # Clean column names

#     # Replace "NIL" with 0 and fill missing values for numeric columns
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({
#         'Quantity': 0,
#         'Market Value (Rs. In Lakhs)': 0,
#         '% age to NAV': 0,
#         'ISIN': '',
#     }, inplace=True)
    
#     # Convert to numeric types where needed
#     # df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     if 'Quantity' in df.columns:
#         df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     elif 'Quantity/Face Value' in df.columns:
#         df['Quantity/Face Value'] = pd.to_numeric(df['Quantity/Face Value'], errors='coerce').fillna(0)
#         df.rename(columns={'Quantity/Face Value': 'Quantity'}, inplace=True)
    
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
    

#     # convert yield % column to folat after stripping %
#     if 'Yield %' in df.columns:
#         df['Yield %'] = df['Yield %'].astype(str).str.rstrip('%').replace('nan', None)
#         df['Yield %'] = pd.to_numeric(df['Yield %'], errors='coerce').fillna(0)

#     print("Columns in DataFrame:", df.columns.tolist())

#     # Get or create the Portfolio object
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     # Initialize individual category totals
#     equity_total = 0
#     debt_total = 0
#     money_market_total = 0
#     others_total = 0

#     # Variable to track the current section/category from the Excel file
#     current_category = None

#     # Create a dictionary to store the total investment per industry/industry rating
#     industry_investments = {}

#     # Create a list to store the instruments and their % age to NAV for top 5
#     instrument_nav_percentages = []

#     # Process each row in the DataFrame
#     for index, row in df.iterrows():
#         name = safe_strip(row.get('Name of the Instruments', ''))
#         lower_name = name.lower()

#         # Detect section header rows and update current_category
#         if lower_name.startswith('equity'):
#             current_category = 'Equity'
#             continue  # Skip header row
#         elif lower_name.startswith('debt'):
#             current_category = 'Debt'
#             continue
#         elif lower_name.startswith('money market'):
#             current_category = 'Money Market' 
#             continue
#         elif lower_name.startswith('others'):
#             current_category = 'Others'
#             continue
        
#         # Detect subtotal/total rows. For Money Market and Others, update the total if desired.
#         if lower_name.startswith('subtotal') or lower_name.startswith('total'):
#             if current_category == 'Money Market':
#                 money_market_total = row['Market Value (Rs. In Lakhs)']
#             elif current_category == 'Others':
#                 others_total = row['Market Value (Rs. In Lakhs)']
#             continue

#         # Skip rows that are not instrument data (empty name or missing ISIN)
#         if not name:
#             continue
#         isin = safe_strip(row.get('ISIN', ''))
#         if not isin:
#             continue

#         # Determine instrument type; default to 'Others' if current_category is not set.
#         instrument_type = current_category if current_category is not None else 'Others'
#         market_value = row['Market Value (Rs. In Lakhs)']

#         # Use get_or_create() to prevent duplicate objects based on ISIN
#         instrument, created = Instrument.objects.get_or_create(
#             isin=isin,
#             defaults={
#                 'instrument_name': name,
#                 'industry_rating': safe_strip(row.get('Industry/Rating', '')),
#                 'quantity': row['Quantity'],
#                 'market_value': market_value,
#                 'percentage_to_nav': row['% age to NAV'],
#                 'yield_percentage': row['yield %'] if 'yield %' in df.columns else None,
#                 'ytc': row.get('^YTC (AT1/Tier 2 bonds)', None),
#                 'instrument_type': instrument_type,
#             }
#         )

#         # If the instrument already exists, update its fields as needed
#         if not created:
#             instrument.instrument_name = name
#             instrument.industry_rating = safe_strip(row.get('Industry/Rating', ''))
#             instrument.quantity = row['Quantity']
#             instrument.market_value = market_value
#             instrument.percentage_to_nav = row['% age to NAV']
#             instrument.yield_percentage = row.get('Yield %', None)
#             instrument.ytc = row.get('^YTC (AT1/Tier 2 bonds)', None)
#             instrument.instrument_type = instrument_type
#             instrument.save()
        
#         # Update totals based on the instrument type
#         if instrument_type == 'Equity':
#             equity_total += market_value
#         elif instrument_type == 'Debt':
#             debt_total += market_value
#         elif instrument_type == 'Money Market':
#             money_market_total += market_value
#         elif instrument_type == 'Others':
#             others_total += market_value
        
#         # Aggregate investment for industry/ratings
#         industry = safe_strip(row.get('Industry/Rating', ''))
#         if industry:
#             if industry not in industry_investments:
#                 industry_investments[industry] = 0
#             industry_investments[industry] += market_value
        
#         # Store the instrument and its % age to NAV for top 5 instruments
#         nav_percentage = row['% age to NAV']
#         instrument_nav_percentages.append((name, nav_percentage, market_value))

#     # Combine Money Market and Others into one category
#     combined_others_total = money_market_total + others_total

#     print(money_market_total, others_total)
#     print(combined_others_total)
#     print("hello world")

#     # Calculate the overall total market value
#     total_market_value = equity_total + debt_total + combined_others_total

#     # Save these totals to the Portfolio (for Equity, Debt, and Others combined)
#     portfolio.equity_total = equity_total
#     portfolio.debt_total = debt_total
#     portfolio.other_total = combined_others_total
#     portfolio.total_market_value = total_market_value
#     portfolio.save()

#     # Calculate percentage shares (avoiding division by zero)
#     total = total_market_value if total_market_value > 0 else 1
#     equity_percentage = (equity_total / total) * 100
#     debt_percentage = (debt_total / total) * 100
#     others_percentage = (combined_others_total / total) * 100

#     # Get top 5 industries/ratings by total market value
#     sorted_industries = sorted(industry_investments.items(), key=lambda x: x[1], reverse=True)[:5]
#     top_industries = [{"industry": industry, "investment": round(investment, 2)} for industry, investment in sorted_industries]
#     print(top_industries)

#     # Get top 5 instruments by % age to NAV
#     sorted_instruments = sorted(instrument_nav_percentages, key=lambda x: x[1], reverse=True)[:5]
#     top_instruments = [{"instrument_name": instrument, "nav_percentage": round(nav_percentage, 2)} for instrument, nav_percentage, _ in sorted_instruments]
#     print(top_instruments)
#     # Plot pie chart for top 5 instruments
#     labels_instruments = [x['instrument_name'] for x in top_instruments]
#     values_instruments = [x['nav_percentage'] for x in top_instruments]

#     fig_instruments, ax_instruments = plt.subplots()
#     ax_instruments.pie(values_instruments, labels=labels_instruments, autopct='%1.1f%%', startangle=90)
#     ax_instruments.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

#     # Plot pie chart for top 5 industries
#     labels_industries = [x['industry'] for x in top_industries]
#     values_industries = [x['investment'] for x in top_industries]

#     fig_industries, ax_industries = plt.subplots()
#     ax_industries.pie(values_industries, labels=labels_industries, autopct='%1.1f%%', startangle=90)
#     ax_industries.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

#     # -------------------------------
#     # Updated horizontal bar chart using percentages
#     # -------------------------------
#     categories = ['Equity', 'Debt', 'Others']
#     percentage_values = [equity_percentage, debt_percentage, others_percentage]

#     fig_bar, ax_bar = plt.subplots()
#     ax_bar.barh(categories, percentage_values, color=['green', 'blue', 'orange'])
#     ax_bar.set_xlabel('Percentage (%)')
#     ax_bar.set_title('Portfolio Market Value Distribution by Category (in %)')

#     # Add percentage labels on the bars
#     for bar in ax_bar.patches:
#         width = bar.get_width()
#         ax_bar.text(width + 1, bar.get_y() + bar.get_height() / 2.0,
#                     f'{width:.2f}%', va='center', fontsize=10)

#     # Convert charts to base64 to display in a web page
#     buffer_instruments = io.BytesIO()
#     fig_instruments.savefig(buffer_instruments, format='png', bbox_inches="tight")
#     buffer_instruments.seek(0)
#     img_instruments = base64.b64encode(buffer_instruments.getvalue()).decode('utf-8')

#     buffer_industries = io.BytesIO()
#     fig_industries.savefig(buffer_industries, format='png', bbox_inches="tight")
#     buffer_industries.seek(0)
#     img_industries = base64.b64encode(buffer_industries.getvalue()).decode('utf-8')

#     buffer_bar = io.BytesIO()
#     fig_bar.savefig(buffer_bar, format='png', bbox_inches="tight")
#     buffer_bar.seek(0)
#     img_bar = base64.b64encode(buffer_bar.getvalue()).decode('utf-8')

#     # Return the results in a Django template
#     return render(request, 'upload_result.html', {
#         'equity_total': equity_total,
#         'debt_total': debt_total,
#         'others_total': combined_others_total,
#         'total_market_value': total_market_value,
#         'equity_percentage': round(equity_percentage, 2),
#         'debt_percentage': round(debt_percentage, 2),
#         'others_percentage': round(others_percentage, 2),
#         'top_industries': top_industries,
#         'top_instruments': top_instruments,
#         'img_instruments': img_instruments,
#         'img_industries': img_industries,
#         'img_bar': img_bar,  # Updated bar chart with percentages
#     })

# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio
# import matplotlib.pyplot as plt
# import io
# import base64

# def safe_strip(value):
#     """Convert to string and strip if it's not NaN."""
#     if isinstance(value, str):
#         return value.strip()
#     return str(value) if pd.notna(value) else ''  # Convert NaN to empty string

# def upload_excel(request):
#     file_path ="c:/Users/91787/Downloads/Monthly Portfolio of Schemes (10).xlsx"
    
#     # Read the Excel file (data starts from row 4; header is at index 3)
#     df = pd.read_excel(file_path, header=3)
#     print(df.columns)
#     df.columns = df.columns.str.strip()  # Clean column names

#     # Replace "NIL" with 0 and fill missing values for numeric columns
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({
#         'Quantity': 0,
#         'Market Value (Rs. In Lakhs)': 0,
#         '% age to NAV': 0,
#         'ISIN': '',
#     }, inplace=True)
    
#     # Convert to numeric types where needed
#     # df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     if 'Quantity' in df.columns:
#         df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     elif 'Quantity/Face Value' in df.columns:
#         df['Quantity/Face Value'] = pd.to_numeric(df['Quantity/Face Value'], errors='coerce').fillna(0)
#         df.rename(columns={'Quantity/Face Value': 'Quantity'}, inplace=True)
    
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
    

#     # convert yield % column to folat after stripping %
#     if 'Yield %' in df.columns:
#         df['Yield %'] = df['Yield %'].astype(str).str.rstrip('%').replace('nan', None)
#         df['Yield %'] = pd.to_numeric(df['Yield %'], errors='coerce').fillna(0)

#     print("Columns in DataFrame:", df.columns.tolist())

#     # Get or create the Portfolio object
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     # Initialize individual category totals
#     equity_total = 0
#     debt_total = 0
#     money_market_total = 0
#     others_total = 0

#     # Variable to track the current section/category from the Excel file
#     current_category = None

#     # Create a dictionary to store the total investment per industry/industry rating
#     industry_investments = {}

#     # Create a list to store the instruments and their % age to NAV for top 5
#     instrument_nav_percentages = []

#     # Process each row in the DataFrame
#     for index, row in df.iterrows():
#         name = safe_strip(row.get('Name of the Instruments', ''))
#         lower_name = name.lower()

#         # Detect section header rows and update current_category
#         if lower_name.startswith('equity'):
#             current_category = 'Equity'
#             continue  # Skip header row
#         elif lower_name.startswith('debt'):
#             current_category = 'Debt'
#             continue
#         elif lower_name.startswith('money market'):
#             current_category = 'Money Market' 
#             continue
#         elif lower_name.startswith('others'):
#             current_category = 'Others'
#             continue
        
#         # Detect subtotal/total rows. For Money Market and Others, update the total if desired.
#         if lower_name.startswith('subtotal') or lower_name.startswith('total'):
#             if current_category == 'Money Market':
#                 money_market_total = row['Market Value (Rs. In Lakhs)']
#             elif current_category == 'Others':
#                 others_total = row['Market Value (Rs. In Lakhs)']
#             continue

#         # Skip rows that are not instrument data (empty name or missing ISIN)
#         if not name:
#             continue
#         isin = safe_strip(row.get('ISIN', ''))
#         if not isin:
#             continue

#         # Determine instrument type; default to 'Others' if current_category is not set.
#         instrument_type = current_category if current_category is not None else 'Others'
#         market_value = row['Market Value (Rs. In Lakhs)']

#         # Use get_or_create() to prevent duplicate objects based on ISIN
#         instrument, created = Instrument.objects.get_or_create(
#             isin=isin,
#             defaults={
#                 'instrument_name': name,
#                 'industry_rating': safe_strip(row.get('Industry/Rating', '')),
#                 'quantity': row['Quantity'],
#                 'market_value': market_value,
#                 'percentage_to_nav': row['% age to NAV'],
#                 'yield_percentage': row['yield %'] if 'yield %' in df.columns else None,
#                 'ytc': row.get('^YTC (AT1/Tier 2 bonds)', None),
#                 'instrument_type': instrument_type,
#             }
#         )

#         # If the instrument already exists, update its fields as needed
#         if not created:
#             instrument.instrument_name = name
#             instrument.industry_rating = safe_strip(row.get('Industry/Rating', ''))
#             instrument.quantity = row['Quantity']
#             instrument.market_value = market_value
#             instrument.percentage_to_nav = row['% age to NAV']
#             instrument.yield_percentage = row.get('Yield %', None)
#             instrument.ytc = row.get('^YTC (AT1/Tier 2 bonds)', None)
#             instrument.instrument_type = instrument_type
#             instrument.save()
        
#         # Update totals based on the instrument type
#         if instrument_type == 'Equity':
#             equity_total += market_value
#         elif instrument_type == 'Debt':
#             debt_total += market_value
#         elif instrument_type == 'Money Market':
#             money_market_total += market_value
#         elif instrument_type == 'Others':
#             others_total += market_value
        
#         # Aggregate investment for industry/ratings
#         industry = safe_strip(row.get('Industry/Rating', ''))
#         if industry:
#             if industry not in industry_investments:
#                 industry_investments[industry] = 0
#             industry_investments[industry] += market_value
        
#         # Store the instrument and its % age to NAV for top 5 instruments
#         nav_percentage = row['% age to NAV']
#         instrument_nav_percentages.append((name, nav_percentage, market_value))

#     # Combine Money Market and Others into one category
#     combined_others_total = money_market_total + others_total

#     print(money_market_total, others_total)
#     print(combined_others_total)
#     print("hello world")

#     # Calculate the overall total market value
#     total_market_value = equity_total + debt_total + combined_others_total

#     # Save these totals to the Portfolio (for Equity, Debt, and Others combined)
#     portfolio.equity_total = equity_total
#     portfolio.debt_total = debt_total
#     portfolio.other_total = combined_others_total
#     portfolio.total_market_value = total_market_value
#     portfolio.save()

#     # Calculate percentage shares (avoiding division by zero)
#     total = total_market_value if total_market_value > 0 else 1
#     equity_percentage = (equity_total / total) * 100
#     debt_percentage = (debt_total / total) * 100
#     others_percentage = (combined_others_total / total) * 100


#     # Get top 5 industries/ratings by total market value
#     sorted_industries = sorted(industry_investments.items(), key=lambda x: x[1], reverse=True)[:5]
#     top_industries = [{"industry": industry, "investment": round(investment, 2)} for industry, investment in sorted_industries]
#     print(top_industries)

#     # Get top 5 instruments by % age to NAV
#     sorted_instruments = sorted(instrument_nav_percentages, key=lambda x: x[1], reverse=True)[:5]
#     top_instruments = [{"instrument_name": instrument, "nav_percentage": round(nav_percentage, 2)} for instrument, nav_percentage, _ in sorted_instruments]
#     print(top_instruments)
#     # Plot pie chart for top 5 instruments
#     labels_instruments = [x['instrument_name'] for x in top_instruments]
#     values_instruments = [x['nav_percentage'] for x in top_instruments]

#     fig_instruments, ax_instruments = plt.subplots()
#     ax_instruments.pie(values_instruments, labels=labels_instruments, autopct='%1.1f%%', startangle=90)
#     ax_instruments.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

#     # Plot pie chart for top 5 industries
#     labels_industries = [x['industry'] for x in top_industries]
#     values_industries = [x['investment'] for x in top_industries]

#     fig_industries, ax_industries = plt.subplots()
#     ax_industries.pie(values_industries, labels=labels_industries, autopct='%1.1f%%', startangle=90)
#     ax_industries.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.

#     # -------------------------------
#     # Updated horizontal bar chart using percentages
#     # -------------------------------
#     categories = ['Equity', 'Debt', 'Others']
#     percentage_values = [equity_percentage, debt_percentage, others_percentage]

#     fig_bar, ax_bar = plt.subplots()
#     ax_bar.barh(categories, percentage_values, color=['green', 'blue', 'orange'])
#     ax_bar.set_xlabel('Percentage (%)')
#     ax_bar.set_title('Portfolio Market Value Distribution by Category (in %)')

#     # Add percentage labels on the bars
#     for bar in ax_bar.patches:
#         width = bar.get_width()
#         ax_bar.text(width + 1, bar.get_y() + bar.get_height() / 2.0,
#                     f'{width:.2f}%', va='center', fontsize=10)

#     # Convert charts to base64 to display in a web page
#     buffer_instruments = io.BytesIO()
#     fig_instruments.savefig(buffer_instruments, format='png', bbox_inches="tight")
#     buffer_instruments.seek(0)
#     img_instruments = base64.b64encode(buffer_instruments.getvalue()).decode('utf-8')

#     buffer_industries = io.BytesIO()
#     fig_industries.savefig(buffer_industries, format='png', bbox_inches="tight")
#     buffer_industries.seek(0)
#     img_industries = base64.b64encode(buffer_industries.getvalue()).decode('utf-8')

#     buffer_bar = io.BytesIO()
#     fig_bar.savefig(buffer_bar, format='png', bbox_inches="tight")
#     buffer_bar.seek(0)
#     img_bar = base64.b64encode(buffer_bar.getvalue()).decode('utf-8')

#     # Return the results in a Django template
#     return render(request, 'upload_result.html', {
#         'equity_total': equity_total,
#         'debt_total': debt_total,
#         'others_total': combined_others_total,
#         'total_market_value': total_market_value,
#         'equity_percentage': round(equity_percentage, 2),
#         'debt_percentage': round(debt_percentage, 2),
#         'others_percentage': round(others_percentage, 2),
#         'top_industries': top_industries,
#         'top_instruments': top_instruments,
#         'img_instruments': img_instruments,
#         'img_industries': img_industries,
#         'img_bar': img_bar,  # Updated bar chart with percentages
#     })

# import pandas as pd
# from django.http import JsonResponse
# from .models import Instrument, Portfolio

# def safe_strip(value):
#     if isinstance(value, str):
#         return value.strip()
#     return str(value) if pd.notna(value) else ''

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes.xlsx"
#     df = pd.read_excel(file_path, header=3)
#     df.columns = df.columns.str.strip()
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({'Quantity': 0, 'Market Value (Rs. In Lakhs)': 0, '% age to NAV': 0, 'ISIN': ''}, inplace=True)
    
#     if 'Quantity' in df.columns:
#         df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     elif 'Quantity/Face Value' in df.columns:
#         df['Quantity/Face Value'] = pd.to_numeric(df['Quantity/Face Value'], errors='coerce').fillna(0)
#         df.rename(columns={'Quantity/Face Value': 'Quantity'}, inplace=True)
    
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
    
#     if 'Yield %' in df.columns:
#         df['Yield %'] = df['Yield %'].astype(str).str.rstrip('%').replace('nan', None)
#         df['Yield %'] = pd.to_numeric(df['Yield %'], errors='coerce').fillna(0)
    
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     category_totals = {}
#     current_category = None
    
#     for index, row in df.iterrows():
#         name = safe_strip(row.get('Name of the Instruments', ''))
#         lower_name = name.lower()

#         if lower_name.startswith(('equity', 'debt', 'money market', 'others')):
#             current_category = name.strip()
#             category_totals[current_category] = 0
#             continue
        
#         if lower_name.startswith(('subtotal', 'total')):
#             if current_category:
#                 category_totals[current_category] = row['Market Value (Rs. In Lakhs)']
#             continue
        
#         if not name:
#             continue
#         isin = safe_strip(row.get('ISIN', ''))
#         if not isin:
#             continue
        
#         market_value = row['Market Value (Rs. In Lakhs)']
#         if current_category:
#             category_totals[current_category] += market_value
    
#     portfolio.total_market_value = sum(category_totals.values())
#     portfolio.save()
    
#     return JsonResponse({'category_totals': category_totals, 'total_market_value': portfolio.total_market_value})

# import pandas as pd
# from django.http import JsonResponse
# from .models import Instrument, Portfolio

# def safe_strip(value):
#     if isinstance(value, str):
#         return value.strip()
#     return str(value) if pd.notna(value) else ''

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes.xlsx"
#     print(f"Reading file from: {file_path}")
#     df = pd.read_excel(file_path, header=3)
#     print("Excel file loaded successfully.")
    
#     df.columns = df.columns.str.strip()
#     print("Column names stripped of whitespace.")
    
#     df.replace("NIL", 0, inplace=True)
#     print("Replaced 'NIL' values with 0.")
    
#     df.fillna({'Quantity': 0, 'Market Value (Rs. In Lakhs)': 0, '% age to NAV': 0, 'ISIN': ''}, inplace=True)
#     print("Filled missing values with defaults.")
    
#     if 'Quantity' in df.columns:
#         df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#         print("Converted 'Quantity' column to numeric.")
#     elif 'Quantity/Face Value' in df.columns:
#         df['Quantity/Face Value'] = pd.to_numeric(df['Quantity/Face Value'], errors='coerce').fillna(0)
#         df.rename(columns={'Quantity/Face Value': 'Quantity'}, inplace=True)
#         print("Converted 'Quantity/Face Value' column to numeric and renamed to 'Quantity'.")
    
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
#     print("Converted 'Market Value' column to numeric.")
    
#     if 'Yield %' in df.columns:
#         df['Yield %'] = df['Yield %'].astype(str).str.rstrip('%').replace('nan', None)
#         df['Yield %'] = pd.to_numeric(df['Yield %'], errors='coerce').fillna(0)
#         print("Processed 'Yield %' column.")
    
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
#     print(f"Portfolio {'created' if created else 'found'}: {portfolio.name} ({portfolio.portfolio_date})")
    
#     category_totals = {}
#     current_category = None
    
#     for index, row in df.iterrows():
#         name = safe_strip(row.get('Name of the Instruments', ''))
#         lower_name = name.lower()

#         if lower_name.startswith(('equity', 'debt', 'money market', 'others')):
#             current_category = name.strip()
#             category_totals[current_category] = 0
#             print(f"Detected category: {current_category}")
#             continue
        
#         if lower_name.startswith(('subtotal', 'total')):
#             if current_category:
#                 category_totals[current_category] = row['Market Value (Rs. In Lakhs)']
#                 print(f"Subtotal for {current_category}: {category_totals[current_category]}")
#             continue
        
#         if not name:
#             continue
        
#         isin = safe_strip(row.get('ISIN', ''))
#         if not isin:
#             continue
        
#         market_value = row['Market Value (Rs. In Lakhs)']
#         if current_category:
#             category_totals[current_category] += market_value
    
#     portfolio.total_market_value = sum(category_totals.values())
#     portfolio.save()
#     print(f"Total Market Value updated: {portfolio.total_market_value}")
    
#     return JsonResponse({'category_totals': category_totals, 'total_market_value': portfolio.total_market_value})


## final updated code 
# import pandas as pd
# from django.http import JsonResponse
# from .models import Instrument, Portfolio

# def safe_strip(value):
#     if isinstance(value, str):
#         return value.strip()
#     return str(value) if pd.notna(value) else ''

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (14).xlsx"
#     df = pd.read_excel(file_path, header=3)
#     df.columns = df.columns.str.strip()
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({'Quantity': 0, 'Market Value (Rs. In Lakhs)': 0, '% age to NAV': 0, 'ISIN': ''}, inplace=True)
    
#     if 'Quantity' in df.columns:
#         df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     elif 'Quantity/Face Value' in df.columns:
#         df['Quantity/Face Value'] = pd.to_numeric(df['Quantity/Face Value'], errors='coerce').fillna(0)
#         df.rename(columns={'Quantity/Face Value': 'Quantity'}, inplace=True)
    
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
    
#     if 'Yield %' in df.columns:
#         df['Yield %'] = df['Yield %'].astype(str).str.rstrip('%').replace('nan', None)
#         df['Yield %'] = pd.to_numeric(df['Yield %'], errors='coerce').fillna(0)
    
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     category_totals = {}
#     current_category = None
    
#     for index, row in df.iterrows():
#         name = safe_strip(row.get('Name of the Instruments', ''))
#         lower_name = name.lower()

#         if lower_name.startswith(('equity', 'debt', 'money market', 'alternative investment', 'treps', 'cash & cash equivalents', 'others')):
#             current_category = name.strip()
#             category_totals[current_category] = 0
#             continue
        
#         if lower_name.startswith(('subtotal', 'total')):
#             if current_category:
#                 category_totals[current_category] = row['Market Value (Rs. In Lakhs)']
#             continue
        
#         if not name:
#             continue
#         isin = safe_strip(row.get('ISIN', ''))
#         if not isin:
#             continue
        
#         market_value = row['Market Value (Rs. In Lakhs)']
#         if current_category:
#             category_totals[current_category] += market_value
    
#     portfolio.total_market_value = sum(category_totals.values())
#     portfolio.save()
    
#     return JsonResponse({'category_totals': category_totals, 'total_market_value': portfolio.total_market_value})

# import pandas as pd
# from django.http import JsonResponse
# from .models import Instrument, Portfolio

# def safe_strip(value):
#     if isinstance(value, str):
#         return value.strip()
#     return str(value) if pd.notna(value) else ''

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (16).xlsx"
#     df = pd.read_excel(file_path, header=3)
#     df.columns = df.columns.str.strip()
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({'Quantity': 0, 'Market Value (Rs. In Lakhs)': 0, '% age to NAV': 0, 'ISIN': ''}, inplace=True)
    
#     if 'Quantity' in df.columns:
#         df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     elif 'Quantity/Face Value' in df.columns:
#         df['Quantity/Face Value'] = pd.to_numeric(df['Quantity/Face Value'], errors='coerce').fillna(0)
#         df.rename(columns={'Quantity/Face Value': 'Quantity'}, inplace=True)
    
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
    
#     if 'Yield %' in df.columns:
#         df['Yield %'] = df['Yield %'].astype(str).str.rstrip('%').replace('nan', None)
#         df['Yield %'] = pd.to_numeric(df['Yield %'], errors='coerce').fillna(0)
    
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     category_totals = {}
#     current_category = None
    
#     for index, row in df.iterrows():
#         name = safe_strip(row.get('Name of the Instruments', ''))
#         lower_name = name.lower()

#         if lower_name.startswith(('equity', 'debt', 'money market', 'alternative investment', 'treps', 'cash & cash equivalents', 'others')):
#             current_category = name.strip()
#             category_totals[current_category] = 0
#             continue
        
#         if lower_name.startswith(('subtotal', 'total')):
#             if current_category:
#                 category_totals[current_category] = row['Market Value (Rs. In Lakhs)']
#             continue
        
#         if not name:
#             continue
#         isin = safe_strip(row.get('ISIN', ''))
#         if not isin:
#             continue
        
#         market_value = row['Market Value (Rs. In Lakhs)']
#         if current_category:
#             category_totals[current_category] += market_value
    
#     portfolio.total_market_value = sum(category_totals.values())
#     portfolio.save()
    
#     return JsonResponse({'category_totals': category_totals, 'total_market_value': portfolio.total_market_value})

import pandas as pd
from django.http import JsonResponse
from collections import Counter
from .models import Instrument, Portfolio

def safe_strip(value):
    if isinstance(value, str):
        return value.strip()
    return str(value) if pd.notna(value) else ''

def upload_excel(request):
    file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (10).xlsx"
    df = pd.read_excel(file_path, header=3)
    df.columns = df.columns.str.strip()
    df.replace("NIL", 0, inplace=True)
    df.fillna({'Quantity': 0, 'Market Value (Rs. In Lakhs)': 0, '% age to NAV': 0, 'ISIN': ''}, inplace=True)
    
    if 'Quantity' in df.columns:
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
    elif 'Quantity/Face Value' in df.columns:
        df['Quantity/Face Value'] = pd.to_numeric(df['Quantity/Face Value'], errors='coerce').fillna(0)
        df.rename(columns={'Quantity/Face Value': 'Quantity'}, inplace=True)
    
    df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
    
    if 'Yield %' in df.columns:
        df['Yield %'] = df['Yield %'].astype(str).str.rstrip('%').replace('nan', None)
        df['Yield %'] = pd.to_numeric(df['Yield %'], errors='coerce').fillna(0)
    
    portfolio, created = Portfolio.objects.get_or_create(
        name="JM Financial Mutual Fund", 
        portfolio_date="2025-01-31"
    )
    
    category_totals = {}
    industry_totals = Counter()
    instrument_totals = Counter()
    current_category = None
    
    for index, row in df.iterrows():
        name = safe_strip(row.get('Name of the Instruments', ''))
        industry = safe_strip(row.get('Industry/Rating', ''))
        market_value = row['Market Value (Rs. In Lakhs)']
        lower_name = name.lower()

        if lower_name.startswith(('equity', 'debt', 'money market', 'others', 'alternative investment fund units', 'treps', 'cash & cash equivalents')):
            current_category = name.strip()
            category_totals[current_category] = 0
            continue
        
        if lower_name.startswith(('subtotal', 'total')):
            if current_category:
                category_totals[current_category] = market_value
            continue
        
        if not name:
            continue
        
        isin = safe_strip(row.get('ISIN', ''))
        if not isin:
            continue
        
        if current_category:
            category_totals[current_category] += market_value
        
        if name:
            instrument_totals[name] += market_value
        
        if industry:
            industry_totals[industry] += market_value
    
    portfolio.total_market_value = sum(category_totals.values())
    portfolio.save()
    
    top_five_instruments = dict(instrument_totals.most_common(5))
    top_five_industries = dict(industry_totals.most_common(5))
    
    return JsonResponse({
        'category_totals': category_totals,
        'total_market_value': portfolio.total_market_value,
        'top_five_instruments': top_five_instruments,
        'top_five_industries': top_five_industries
    })



