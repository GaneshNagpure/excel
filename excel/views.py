# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio, CashEquivalents

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (1).xlsx"
    
#     # Read the Excel file using pandas
#     df = pd.read_excel(file_path)

#     # Process the data and store it in the database
#     portfolio, created = Portfolio.objects.get_or_create(name="JM Financial Mutual Fund", portfolio_date="2025-01-31")
    
#     # Assuming the Excel data has the structure like the one you provided
#     for index, row in df.iterrows():
#         instrument = Instrument.objects.create(
#             instrument_name=row['Name of the Instruments'],
#             industry_rating=row['Industry/Rating'],
#             quantity=row['Quantity'],
#             market_value=row['Market Value (Rs. In Lakhs)'],
#             percentage_to_nav=row['% age to NAV'],
#             isin=row['ISIN'],
#             yield_percentage=row['Yield %'] if 'Yield %' in row else None,
#             ytc=row['^YTC (AT1/Tier 2 bonds)'] if '^YTC (AT1/Tier 2 bonds)' in row else None,
#             instrument_type=row['Instrument Type']  # Make sure to categorize instrument type (e.g., Equity, Debt, etc.)
#         )
        
#         # Update the totals in the portfolio object based on the instrument type
#         if 'Equity' in row['Instrument Type']:
#             portfolio.equity_total += row['Market Value (Rs. In Lakhs)']
#         elif 'Debt' in row['Instrument Type']:
#             portfolio.debt_total += row['Market Value (Rs. In Lakhs)']
#         else:
#             portfolio.other_total += row['Market Value (Rs. In Lakhs)']
    
#     # Save the updated portfolio
#     portfolio.total_market_value = portfolio.equity_total + portfolio.debt_total + portfolio.other_total
#     portfolio.save()

#     # Return the totals as a JSON response for later visualization
#     data = {
#         'equity_total': portfolio.equity_total,
#         'debt_total': portfolio.debt_total,
#         'other_total': portfolio.other_total,
#     }

#     return JsonResponse(data)


# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio, CashEquivalents

# def upload_excel(request):
#     file_path =  "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (1).xlsx"
    
#     # Read the Excel file using pandas
#     df = pd.read_excel(file_path, header=3)
#     print(df.columns)
#     df.columns = df.columns.str.strip()
#     print("Columns in DataFrame:", df.columns.tolist())



#     # Process the data and store it in the database
#     portfolio, created = Portfolio.objects.get_or_create(name="JM Financial Mutual Fund", portfolio_date="2025-01-31")
    
#     # Initialize all totals to zero
#     portfolio.equity_total = 0
#     portfolio.debt_total = 0
#     portfolio.other_total = 0

#     # Assuming the Excel data has the structure like the one you provided
#     for index, row in df.iterrows():
#         instrument = Instrument.objects.create(
#             instrument_name=row['Name of the Instruments'],
#             industry_rating=row['Industry/Rating'],
#             quantity=row['Quantity'],
#             market_value=row['Market Value (Rs. In Lakhs)'],
#             percentage_to_nav=row['% age to NAV'],
#             isin=row['ISIN'],
#             yield_percentage=row['Yield %'] if 'Yield %' in row else None,
#             ytc=row['^YTC (AT1/Tier 2 bonds)'] if '^YTC (AT1/Tier 2 bonds)' in row else None,
#             instrument_type=row['Instrument Type']
#         )
        
#         # Update the totals in the portfolio object based on the instrument type
#         if 'Equity' in row['Instrument Type']:
#             portfolio.equity_total += row['Market Value (Rs. In Lakhs)']
#         elif 'Debt' in row['Instrument Type']:
#             portfolio.debt_total += row['Market Value (Rs. In Lakhs)']
#         else:
#             portfolio.other_total += row['Market Value (Rs. In Lakhs)']
    
#     # Ensure all totals are correctly calculated and set before saving
#     portfolio.total_market_value = portfolio.equity_total + portfolio.debt_total + portfolio.other_total

#     # Save the updated portfolio
#     portfolio.save()

#     # Return the totals as a JSON response for later visualization
#     data = {
#         'equity_total': portfolio.equity_total,
#         'debt_total': portfolio.debt_total,
#         'other_total': portfolio.other_total,
#     }

#     return JsonResponse(data)

# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio

# def categorize_instrument(instrument_name):
#     """Categorizes instruments based on their names."""
#     if any(keyword in instrument_name.lower() for keyword in ['equity', 'stock', 'share']):
#         return 'Equity'
#     elif any(keyword in instrument_name.lower() for keyword in ['bond', 'debt', 'fund']):
#         return 'Debt'
#     else:
#         return 'Other'

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (1).xlsx"
    
#     # Read the Excel file using pandas (header starts from row 4)
#     df = pd.read_excel(file_path, header=3)
#     df.columns = df.columns.str.strip()  # Remove spaces from column names

#     print("Columns in DataFrame:", df.columns.tolist())  # Debugging output

#     # Ensure Portfolio object exists
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     # Initialize totals
#     portfolio.equity_total = 0
#     portfolio.debt_total = 0
#     portfolio.other_total = 0

#     # Process the data and store it in the database
#     for index, row in df.iterrows():
#         instrument_name = row.get('Name of the Instruments', 'Unknown')
#         instrument_type = categorize_instrument(instrument_name)  # Auto-categorize

#         instrument = Instrument.objects.create(
#             instrument_name=instrument_name,
#             industry_rating=row.get('Industry/Rating', ''),
#             quantity=row.get('Quantity', 0),
#             market_value=row.get('Market Value (Rs. In Lakhs)', 0),
#             percentage_to_nav=row.get('% age to NAV', 0),
#             isin=row.get('ISIN', ''),
#             yield_percentage=row.get('Yield %', None),
#             ytc=row.get('^YTC (AT1/Tier 2 bonds)', None),
#             instrument_type=instrument_type  # Assign dynamically
#         )
        
#         # Update totals based on categorized type
#         if instrument_type == 'Equity':
#             portfolio.equity_total += row.get('Market Value (Rs. In Lakhs)', 0)
#         elif instrument_type == 'Debt':
#             portfolio.debt_total += row.get('Market Value (Rs. In Lakhs)', 0)
#         else:
#             portfolio.other_total += row.get('Market Value (Rs. In Lakhs)', 0)

#     # Calculate the total market value
#     portfolio.total_market_value = portfolio.equity_total + portfolio.debt_total + portfolio.other_total
#     portfolio.save()

#     # Return the totals as a JSON response
#     data = {
#         'equity_total': portfolio.equity_total,
#         'debt_total': portfolio.debt_total,
#         'other_total': portfolio.other_total,
#     }

#     return JsonResponse(data)

# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio

# def categorize_instrument(instrument_name):
#     """Categorizes instruments based on their names."""
#     if pd.isna(instrument_name) or not isinstance(instrument_name, str):
#         return 'Other'
#     if any(keyword in instrument_name.lower() for keyword in ['equity', 'stock', 'share']):
#         return 'Equity'
#     elif any(keyword in instrument_name.lower() for keyword in ['bond', 'debt', 'fund']):
#         return 'Debt'
#     else:
#         return 'Other'

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (1).xlsx"
    
#     # Read the Excel file using pandas (header starts from row 4)
#     df = pd.read_excel(file_path, header=3)
#     df.columns = df.columns.str.strip()  # Remove spaces from column names

#     # Replace NaN values with defaults
#     df = df.fillna({
#         'Quantity': 0,  # Replace missing quantity with 0
#         'Market Value (Rs. In Lakhs)': 0,  # Replace missing market value with 0
#         '% age to NAV': 0,  # Replace missing percentage with 0
#         'ISIN': '',  # Replace missing ISIN with empty string
#     })

#     print("Columns in DataFrame:", df.columns.tolist())  # Debugging output

#     # Ensure Portfolio object exists
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     # Initialize totals
#     portfolio.equity_total = 0
#     portfolio.debt_total = 0
#     portfolio.other_total = 0

#     # Process the data and store it in the database
#     for index, row in df.iterrows():
#         instrument_name = row.get('Name of the Instruments', 'Unknown')
#         instrument_type = categorize_instrument(instrument_name)  # Auto-categorize

#         instrument = Instrument.objects.create(
#             instrument_name=instrument_name,
#             industry_rating=row.get('Industry/Rating', ''),
#             quantity=row['Quantity'],  # No more NaN issues
#             market_value=row['Market Value (Rs. In Lakhs)'],
#             percentage_to_nav=row['% age to NAV'],
#             isin=row['ISIN'],
#             yield_percentage=row.get('Yield %', None),
#             ytc=row.get('^YTC (AT1/Tier 2 bonds)', None),
#             instrument_type=instrument_type  # Assign dynamically
#         )
        
#         # Update totals based on categorized type
#         if instrument_type == 'Equity':
#             portfolio.equity_total += row['Market Value (Rs. In Lakhs)']
#         elif instrument_type == 'Debt':
#             portfolio.debt_total += row['Market Value (Rs. In Lakhs)']
#         else:
#             portfolio.other_total += row['Market Value (Rs. In Lakhs)']

#     # Calculate the total market value
#     portfolio.total_market_value = portfolio.equity_total + portfolio.debt_total + portfolio.other_total
#     portfolio.save()

#     # Return the totals as a JSON response
#     data = {
#         'equity_total': portfolio.equity_total,
#         'debt_total': portfolio.debt_total,
#         'other_total': portfolio.other_total,
#     }

#     return JsonResponse(data)

# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio

# def categorize_instrument(instrument_name):
#     """Categorizes instruments based on their names."""
#     if pd.isna(instrument_name) or not isinstance(instrument_name, str):
#         return 'Other'
#     if any(keyword in instrument_name.lower() for keyword in ['equity', 'stock', 'share']):
#         return 'Equity'
#     elif any(keyword in instrument_name.lower() for keyword in ['bond', 'debt', 'fund']):
#         return 'Debt'
#     else:
#         return 'Other'

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (1).xlsx"
    
#     # Read the Excel file using pandas (header starts from row 4)
#     df = pd.read_excel(file_path, header=3)
#     df.columns = df.columns.str.strip()  # Remove spaces from column names

#     # Replace "NIL" and NaN values with 0
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({
#         'Quantity': 0,
#         'Market Value (Rs. In Lakhs)': 0,
#         '% age to NAV': 0,
#         'ISIN': '',
#     }, inplace=True)

#     # Convert numeric columns to correct type
#     df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)

#     print("Columns in DataFrame:", df.columns.tolist())  # Debugging output

#     # Ensure Portfolio object exists
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     # Initialize totals
#     portfolio.equity_total = 0
#     portfolio.debt_total = 0
#     portfolio.other_total = 0

#     # Process the data and store it in the database
#     for index, row in df.iterrows():
#         instrument_name = row.get('Name of the Instruments', 'Unknown')
#         instrument_type = categorize_instrument(instrument_name)  # Auto-categorize

#         instrument = Instrument.objects.create(
#             instrument_name=instrument_name,
#             industry_rating=row.get('Industry/Rating', ''),
#             quantity=row['Quantity'],
#             market_value=row['Market Value (Rs. In Lakhs)'],
#             percentage_to_nav=row['% age to NAV'],
#             isin=row['ISIN'],
#             yield_percentage=row.get('Yield %', None),
#             ytc=row.get('^YTC (AT1/Tier 2 bonds)', None),
#             instrument_type=instrument_type  
#         )
        
#         # Update totals based on categorized type
#         if instrument_type == 'Equity':
#             portfolio.equity_total += row['Market Value (Rs. In Lakhs)']
#         elif instrument_type == 'Debt':
#             portfolio.debt_total += row['Market Value (Rs. In Lakhs)']
#         else:
#             portfolio.other_total += row['Market Value (Rs. In Lakhs)']

#     # Calculate the total market value
#     portfolio.total_market_value = portfolio.equity_total + portfolio.debt_total + portfolio.other_total
#     portfolio.save()

#     # Return the totals as a JSON response
#     data = {
#         'equity_total': portfolio.equity_total,
#         'debt_total': portfolio.debt_total,
#         'other_total': portfolio.other_total,
#     }

#     return JsonResponse(data)

# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (1).xlsx"
    
#     # Read the Excel file.
#     # header=3 means that row 4 (0-indexed row 3) is taken as the header row.
#     df = pd.read_excel(file_path, header=3)
#     df.columns = df.columns.str.strip()  # Clean column names

#     # Replace "NIL" with 0 and fill missing values for key numeric fields.
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({
#         'Quantity': 0,
#         'Market Value (Rs. In Lakhs)': 0,
#         '% age to NAV': 0,
#         'ISIN': '',
#     }, inplace=True)

#     # Ensure numeric conversion for these columns.
#     df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)

#     print("Columns in DataFrame:", df.columns.tolist())  # Debug output

#     # Get or create the Portfolio.
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
#     # Initialize totals.
#     portfolio.equity_total = 0
#     portfolio.debt_total = 0
#     portfolio.other_total = 0

#     # We'll use this variable to "remember" the current section.
#     current_category = None

#     # Process each row in the DataFrame.
#     for index, row in df.iterrows():
#         # Get the cell value from the "Name of the Instruments" column.
#         name = str(row.get('Name of the Instruments', '')).strip()

#         # Detect section headers (which are not instrument data rows)
#         # Adjust these checks as needed for your Excel's text.
#         lower_name = name.lower()
#         if lower_name.startswith('equity'):
#             current_category = 'Equity'
#             continue  # skip the header row itself
#         elif lower_name.startswith('debt'):
#             current_category = 'Debt'
#             continue
#         elif lower_name.startswith('money market'):
#             # If you want Money Market instruments to be counted as "Other"
#             current_category = 'Other'
#             continue
#         elif lower_name.startswith('subtotal'):
#             # Skip any subtotal rows.
#             continue

#         # Skip rows that are clearly not data rows (for example, missing a valid ISIN).
#         isin = str(row.get('ISIN', '')).strip()
#         if not isin:
#             continue

#         # If no category has been set yet, default to 'Other'
#         instrument_type = current_category if current_category is not None else 'Other'

#         # Create the Instrument record.
#         instrument = Instrument.objects.create(
#             instrument_name=name,
#             industry_rating=row.get('Industry/Rating', ''),
#             quantity=row['Quantity'],
#             market_value=row['Market Value (Rs. In Lakhs)'],
#             percentage_to_nav=row['% age to NAV'],
#             isin=isin,
#             yield_percentage=row.get('Yield %', None),
#             ytc=row.get('^YTC (AT1/Tier 2 bonds)', None),
#             instrument_type=instrument_type
#         )

#         # Update portfolio totals based on the instrument type.
#         if instrument_type == 'Equity':
#             portfolio.equity_total += row['Market Value (Rs. In Lakhs)']
#         elif instrument_type == 'Debt':
#             portfolio.debt_total += row['Market Value (Rs. In Lakhs)']
#         else:
#             portfolio.other_total += row['Market Value (Rs. In Lakhs)']

#     # Calculate overall total market value.
#     portfolio.total_market_value = portfolio.equity_total + portfolio.debt_total + portfolio.other_total
#     portfolio.save()

#     # Return the totals as JSON.
#     data = {
#         'equity_total': portfolio.equity_total,
#         'debt_total': portfolio.debt_total,
#         'other_total': portfolio.other_total,
#     }

#     return JsonResponse(data)


# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (1).xlsx"
    
#     # Read the Excel file.
#     # header=3 means row 4 (0-indexed: 3) is used as the header row.
#     df = pd.read_excel(file_path, header=3)
#     df.columns = df.columns.str.strip()  # Clean column names

#     # Replace "NIL" with 0 and fill missing values for numeric fields.
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({
#         'Quantity': 0,
#         'Market Value (Rs. In Lakhs)': 0,
#         '% age to NAV': 0,
#         'ISIN': '',
#     }, inplace=True)

#     # Ensure numeric conversion.
#     df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)

#     print("Columns in DataFrame:", df.columns.tolist())  # Debug output

#     # Get or create the Portfolio.
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
#     # Initialize totals.
#     portfolio.equity_total = 0
#     portfolio.debt_total = 0
#     portfolio.other_total = 0

#     # This variable will remember the current section.
#     current_category = None

#     # Process each row.
#     for index, row in df.iterrows():
#         # Get the cell value from the "Name of the Instruments" column.
#         name = str(row.get('Name of the Instruments', '')).strip()
#         lower_name = name.lower()

#         # Detect section header rows to update the current category.
#         if lower_name.startswith('equity'):
#             current_category = 'Equity'
#             continue  # Skip header row itself.
#         elif lower_name.startswith('debt'):
#             current_category = 'Debt'
#             continue
#         elif lower_name.startswith('money market') or lower_name.startswith('tri party repo'):
#             current_category = 'Other'
#             continue

#         # Check if the row is a subtotal/total row.
#         if lower_name.startswith('subtotal') or lower_name.startswith('total'):
#             # If we're in the Other section, use this row's market value.
#             if current_category == 'Other':
#                 # Here we assume the subtotal row contains the total for that section.
#                 portfolio.other_total = row['Market Value (Rs. In Lakhs)']
#             continue  # Skip further processing for subtotal rows.

#         # Optionally skip rows with no instrument name.
#         if not name:
#             continue

#         # For instrument rows, if ISIN is missing then assume it's not a data row.
#         isin = str(row.get('ISIN', '')).strip()
#         if not isin:
#             continue

#         # Determine instrument type based on the current section.
#         instrument_type = current_category if current_category is not None else 'Other'
#         market_value = row['Market Value (Rs. In Lakhs)']

#         # Create the Instrument record.
#         Instrument.objects.create(
#             instrument_name=name,
#             industry_rating=row.get('Industry/Rating', ''),
#             quantity=row['Quantity'],
#             market_value=market_value,
#             percentage_to_nav=row['% age to NAV'],
#             isin=isin,
#             yield_percentage=row.get('Yield %', None),
#             ytc=row.get('^YTC (AT1/Tier 2 bonds)', None),
#             instrument_type=instrument_type
#         )
        
#         # Update portfolio totals for Equity and Debt using individual rows.
#         if instrument_type == 'Equity':
#             portfolio.equity_total += market_value
#         elif instrument_type == 'Debt':
#             portfolio.debt_total += market_value
#         # For 'Other', if there are individual rows (in addition to a subtotal row),
#         # you can also add them here:
#         elif instrument_type == 'Other':
#             portfolio.other_total += market_value

#     # Calculate the overall total market value.
#     portfolio.total_market_value = (
#         portfolio.equity_total + portfolio.debt_total + portfolio.other_total
#     )
#     portfolio.save()

#     # Return the totals as a JSON response.
#     data = {
#         'equity_total': portfolio.equity_total,
#         'debt_total': portfolio.debt_total,
#         'other_total': portfolio.other_total,
#     }
#     return JsonResponse(data)

# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (1).xlsx"
    
#     # Read the Excel file (data starts from row 4; header row is at index 3)
#     df = pd.read_excel(file_path, header=3)
#     df.columns = df.columns.str.strip()  # Clean column names

#     # Replace "NIL" with 0 and fill missing values for numeric fields
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({
#         'Quantity': 0,
#         'Market Value (Rs. In Lakhs)': 0,
#         '% age to NAV': 0,
#         'ISIN': '',
#     }, inplace=True)

#     # Convert numeric columns
#     df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)

#     print("Columns in DataFrame:", df.columns.tolist())

#     # Get or create the Portfolio
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
#     # Initialize totals
#     portfolio.equity_total = 0
#     portfolio.debt_total = 0
#     portfolio.other_total = 0

#     # Variable to track the current section/category
#     current_category = None

#     # Process each row in the DataFrame
#     for index, row in df.iterrows():
#         # Get the instrument name and clean it up
#         name = str(row.get('Name of the Instruments', '')).strip()
#         lower_name = name.lower()

#         # Detect section headers to update the current category
#         if lower_name.startswith('equity'):
#             current_category = 'Equity'
#             continue  # Skip the header row itself
#         elif lower_name.startswith('debt'):
#             current_category = 'Debt'
#             continue
#         elif lower_name.startswith('money market') or lower_name.startswith('tri party repo'):
#             current_category = 'Other'
#             continue

#         # Check if the row is a subtotal/total row.
#         if lower_name.startswith('subtotal') or lower_name.startswith('total'):
#             # For Other, if there's a subtotal row, update the total.
#             if current_category == 'Other':
#                 portfolio.other_total = row['Market Value (Rs. In Lakhs)']
#             continue  # Skip subtotal rows

#         # Skip rows that don't seem to be instrument data (e.g., empty name or missing ISIN)
#         if not name:
#             continue
#         isin = str(row.get('ISIN', '')).strip()
#         if not isin:
#             continue

#         # Use the current category if available, else default to "Other"
#         instrument_type = current_category if current_category is not None else 'Other'
#         market_value = row['Market Value (Rs. In Lakhs)']

#         # Create the Instrument record
#         Instrument.objects.create(
#             instrument_name=name,
#             industry_rating=row.get('Industry/Rating', ''),
#             quantity=row['Quantity'],
#             market_value=market_value,
#             percentage_to_nav=row['% age to NAV'],
#             isin=isin,
#             yield_percentage=row.get('Yield %', None),
#             ytc=row.get('^YTC (AT1/Tier 2 bonds)', None),
#             instrument_type=instrument_type
#         )
        
#         # Update totals for Equity and Debt from individual rows.
#         if instrument_type == 'Equity':
#             portfolio.equity_total += market_value
#         elif instrument_type == 'Debt':
#             portfolio.debt_total += market_value
#         elif instrument_type == 'Other':
#             # If there are individual rows for 'Other' in addition to subtotal rows,
#             # you can accumulate them as well:
#             portfolio.other_total += market_value

#     # Calculate overall total market value.
#     portfolio.total_market_value = (
#         portfolio.equity_total + portfolio.debt_total + portfolio.other_total
#     )
#     portfolio.save()

#     # Calculate percentages. Avoid division by zero.
#     total = portfolio.total_market_value if portfolio.total_market_value > 0 else 1
#     equity_percentage = (portfolio.equity_total / total) * 100
#     debt_percentage = (portfolio.debt_total / total) * 100
#     other_percentage = (portfolio.other_total / total) * 100

#     # Return both the totals and their percentages as JSON.
#     data = {
#         'equity_total': portfolio.equity_total,
#         'debt_total': portfolio.debt_total,
#         'other_total': portfolio.other_total,
#         'equity_percentage': round(equity_percentage, 2),
#         'debt_percentage': round(debt_percentage, 2),
#         'other_percentage': round(other_percentage, 2),
#         'total_market_value': portfolio.total_market_value,
#     }

#     return JsonResponse(data)

# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (1).xlsx"
    
#     # Read the Excel file (data starts from row 4; header row is at index 3)
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
    
#     # Ensure numeric columns are converted correctly
#     df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)

#     print("Columns in DataFrame:", df.columns.tolist())  # Debugging output

#     # Get or create the Portfolio object
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
#     # Initialize totals
#     portfolio.equity_total = 0
#     portfolio.debt_total = 0
#     portfolio.other_total = 0

#     # Variable to track the current section/category from the Excel file.
#     current_category = None

#     # Process each row of the DataFrame
#     for index, row in df.iterrows():
#         # Get the cell value from the "Name of the Instruments" column
#         name = str(row.get('Name of the Instruments', '')).strip()
#         lower_name = name.lower()

#         # Detect section headers (update current_category accordingly)
#         if lower_name.startswith('equity'):
#             current_category = 'Equity'
#             continue  # Skip the header row itself
#         elif lower_name.startswith('debt'):
#             current_category = 'Debt'
#             continue
#         elif lower_name.startswith('money market'):
#             current_category = 'Others'  # Map Money Market into Others
#             continue
#         elif lower_name.startswith('other'):
#             current_category = 'Others'
#             continue

#         # Detect subtotal/total rows (if present) and use them to set totals for the Others category
#         if lower_name.startswith('subtotal') or lower_name.startswith('total'):
#             if current_category == 'Others':
#                 portfolio.other_total = row['Market Value (Rs. In Lakhs)']
#             continue  # Skip processing subtotal rows

#         # Skip rows that are not instrument data (e.g., empty names or missing ISIN)
#         if not name:
#             continue
#         isin = str(row.get('ISIN', '')).strip()
#         if not isin:
#             continue

#         # Use the current category if available; default to "Others" if not set
#         instrument_type = current_category if current_category is not None else 'Others'
#         market_value = row['Market Value (Rs. In Lakhs)']

#         # Create an Instrument record
#         Instrument.objects.create(
#             instrument_name=name,
#             industry_rating=row.get('Industry/Rating', ''),
#             quantity=row['Quantity'],
#             market_value=market_value,
#             percentage_to_nav=row['% age to NAV'],
#             isin=isin,
#             yield_percentage=row.get('Yield %', None),
#             ytc=row.get('^YTC (AT1/Tier 2 bonds)', None),
#             instrument_type=instrument_type
#         )
        
#         # Update portfolio totals based on instrument type
#         if instrument_type == 'Equity':
#             portfolio.equity_total += market_value
#         elif instrument_type == 'Debt':
#             portfolio.debt_total += market_value
#         elif instrument_type == 'Others':
#             portfolio.other_total += market_value

#     # Calculate overall total market value
#     portfolio.total_market_value = portfolio.equity_total + portfolio.debt_total + portfolio.other_total
#     portfolio.save()

#     # Calculate percentage shares (avoid division by zero)
#     total = portfolio.total_market_value if portfolio.total_market_value > 0 else 1
#     equity_percentage = (portfolio.equity_total / total) * 100
#     debt_percentage = (portfolio.debt_total / total) * 100
#     others_percentage = (portfolio.other_total / total) * 100

#     # Return totals and percentages as a JSON response
#     data = {
#         'equity_total': portfolio.equity_total,
#         'debt_total': portfolio.debt_total,
#         'others_total': portfolio.other_total,
#         'equity_percentage': round(equity_percentage, 2),
#         'debt_percentage': round(debt_percentage, 2),
#         'others_percentage': round(others_percentage, 2),
#         'total_market_value': portfolio.total_market_value,
#     }

#     return JsonResponse(data)


# import pandas as pd
# from django.shortcuts import render
# from django.http import JsonResponse
# from .models import Instrument, Portfolio

# def upload_excel(request):
#     file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (1).xlsx"
    
#     # Read the Excel file (data starts from row 4; header row is at index 3)
#     df = pd.read_excel(file_path, header=3)
#     df.columns = df.columns.str.strip()  # Clean column names

#     # Replace "NIL" with 0 and fill missing values for key numeric columns
#     df.replace("NIL", 0, inplace=True)
#     df.fillna({
#         'Quantity': 0,
#         'Market Value (Rs. In Lakhs)': 0,
#         '% age to NAV': 0,
#         'ISIN': '',
#     }, inplace=True)
    
#     # Convert these columns to numeric
#     df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
#     df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)

#     print("Columns in DataFrame:", df.columns.tolist())

#     # Get or create the Portfolio
#     portfolio, created = Portfolio.objects.get_or_create(
#         name="JM Financial Mutual Fund", 
#         portfolio_date="2025-01-31"
#     )
    
#     # Initialize category totals
#     equity_total = 0
#     debt_total = 0
#     money_market_total = 0
#     others_total = 0

#     # Use this variable to track the current section from the Excel file.
#     current_category = None

#     # Process each row of the DataFrame
#     for index, row in df.iterrows():
#         # Get the cell value for the instrument name
#         name = str(row.get('Name of the Instruments', '')).strip()
#         lower_name = name.lower()

#         # Detect section header rows to update the current_category.
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

#         # Detect subtotal or total rows; these may be present for any category.
#         if lower_name.startswith('subtotal') or lower_name.startswith('total'):
#             # If you wish to use the subtotal row only for a category, you might do:
#             if current_category == 'Money Market':
#                 money_market_total = row['Market Value (Rs. In Lakhs)']
#             elif current_category == 'Others':
#                 others_total = row['Market Value (Rs. In Lakhs)']
#             # For Equity or Debt, we assume individual instrument rows are provided.
#             continue

#         # Skip rows that are not actual instrument data (empty name or missing ISIN)
#         if not name:
#             continue
#         isin = str(row.get('ISIN', '')).strip()
#         if not isin:
#             continue

#         # Use the current category if available; default to Others if not set.
#         instrument_type = current_category if current_category is not None else 'Others'
#         market_value = row['Market Value (Rs. In Lakhs)']

#         # Create an Instrument record
#         Instrument.objects.create(
#             instrument_name=name,
#             industry_rating=row.get('Industry/Rating', ''),
#             quantity=row['Quantity'],
#             market_value=market_value,
#             percentage_to_nav=row['% age to NAV'],
#             isin=isin,
#             yield_percentage=row.get('Yield %', None),
#             ytc=row.get('^YTC (AT1/Tier 2 bonds)', None),
#             instrument_type=instrument_type
#         )
        
#         # Update category totals based on instrument type
#         if instrument_type == 'Equity':
#             equity_total += market_value
#         elif instrument_type == 'Debt':
#             debt_total += market_value
#         elif instrument_type == 'Money Market':
#             money_market_total += market_value
#         elif instrument_type == 'Others':
#             others_total += market_value

#     # Calculate overall total market value
#     total_market_value = equity_total + debt_total + money_market_total + others_total

#     # Save these totals into the Portfolio (if you want to store overall total in your model)
#     portfolio.equity_total = equity_total
#     portfolio.debt_total = debt_total
#     # Here, you might decide whether to store money market and others separately or combined.
#     # For this example, we store them separately:
#     portfolio.other_total = money_market_total + others_total
#     portfolio.total_market_value = total_market_value
#     portfolio.save()

#     # Calculate percentages for each category (avoid division by zero)
#     total = total_market_value if total_market_value > 0 else 1
#     equity_percentage = (equity_total / total) * 100
#     debt_percentage = (debt_total / total) * 100
#     money_market_percentage = (money_market_total / total) * 100
#     others_percentage = (others_total / total) * 100

#     # Return the results as a JSON response.
#     data = {
#         'equity_total': equity_total,
#         'debt_total': debt_total,
#         'money_market_total': money_market_total,
#         'others_total': others_total,
#         'total_market_value': total_market_value,
#         'equity_percentage': round(equity_percentage, 2),
#         'debt_percentage': round(debt_percentage, 2),
#         'money_market_percentage': round(money_market_percentage, 2),
#         'others_percentage': round(others_percentage, 2),
#     }

#     return JsonResponse(data)


import pandas as pd
from django.shortcuts import render
from django.http import JsonResponse
from .models import Instrument, Portfolio

def upload_excel(request):
    file_path = "C:/Users/91787/Downloads/Monthly Portfolio of Schemes (1).xlsx"
    
    # Read the Excel file (data starts from row 4; header is at index 3)
    df = pd.read_excel(file_path, header=3)
    df.columns = df.columns.str.strip()  # Clean column names

    # Replace "NIL" with 0 and fill missing values for numeric columns
    df.replace("NIL", 0, inplace=True)
    df.fillna({
        'Quantity': 0,
        'Market Value (Rs. In Lakhs)': 0,
        '% age to NAV': 0,
        'ISIN': '',
    }, inplace=True)
    
    # Convert to numeric types where needed
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
    df['Market Value (Rs. In Lakhs)'] = pd.to_numeric(df['Market Value (Rs. In Lakhs)'], errors='coerce').fillna(0)
    
    print("Columns in DataFrame:", df.columns.tolist())

    # Get or create the Portfolio object
    portfolio, created = Portfolio.objects.get_or_create(
        name="JM Financial Mutual Fund", 
        portfolio_date="2025-01-31"
    )
    
    # Initialize individual category totals
    equity_total = 0
    debt_total = 0
    money_market_total = 0
    others_total = 0

    # Variable to track the current section/category from the Excel file
    current_category = None

    # Process each row in the DataFrame
    for index, row in df.iterrows():
        name = str(row.get('Name of the Instruments', '')).strip()
        lower_name = name.lower()

        # Detect section header rows and update current_category
        if lower_name.startswith('equity'):
            current_category = 'Equity'
            continue  # Skip header row
        elif lower_name.startswith('debt'):
            current_category = 'Debt'
            continue
        elif lower_name.startswith('money market'):
            current_category = 'Money Market'
            continue
        elif lower_name.startswith('other'):
            current_category = 'Others'
            continue

        # Detect subtotal/total rows. For Money Market and Others, we can update the total if desired.
        if lower_name.startswith('subtotal') or lower_name.startswith('total'):
            if current_category == 'Money Market':
                money_market_total = row['Market Value (Rs. In Lakhs)']
            elif current_category == 'Others':
                others_total = row['Market Value (Rs. In Lakhs)']
            continue

        # Skip rows that are not instrument data (empty name or missing ISIN)
        if not name:
            continue
        isin = str(row.get('ISIN', '')).strip()
        if not isin:
            continue

        # Determine instrument type; default to 'Others' if current_category is not set.
        instrument_type = current_category if current_category is not None else 'Others'
        market_value = row['Market Value (Rs. In Lakhs)']

        # Create the Instrument record
        Instrument.objects.create(
            instrument_name=name,
            industry_rating=row.get('Industry/Rating', ''),
            quantity=row['Quantity'],
            market_value=market_value,
            percentage_to_nav=row['% age to NAV'],
            isin=isin,
            yield_percentage=row.get('Yield %', None),
            ytc=row.get('^YTC (AT1/Tier 2 bonds)', None),
            instrument_type=instrument_type
        )
        
        # Update totals based on the instrument type
        if instrument_type == 'Equity':
            equity_total += market_value
        elif instrument_type == 'Debt':
            debt_total += market_value
        elif instrument_type == 'Money Market':
            money_market_total += market_value
        elif instrument_type == 'Others':
            others_total += market_value

    # Combine Money Market and Others into one category
    combined_others_total = money_market_total + others_total

    # Calculate the overall total market value
    total_market_value = equity_total + debt_total + combined_others_total

    # Save these totals to the Portfolio (for Equity, Debt, and Others combined)
    portfolio.equity_total = equity_total
    portfolio.debt_total = debt_total
    portfolio.other_total = combined_others_total
    portfolio.total_market_value = total_market_value
    portfolio.save()

    # Calculate percentage shares (avoiding division by zero)
    total = total_market_value if total_market_value > 0 else 1
    equity_percentage = (equity_total / total) * 100
    debt_percentage = (debt_total / total) * 100
    others_percentage = (combined_others_total / total) * 100

    # Return the results as a JSON response
    data = {
        'equity_total': equity_total,
        'debt_total': debt_total,
        'others_total': combined_others_total,
        'total_market_value': total_market_value,
        'equity_percentage': round(equity_percentage, 2),
        'debt_percentage': round(debt_percentage, 2),
        'others_percentage': round(others_percentage, 2),
    }

    return JsonResponse(data)

# import matplotlib.pyplot as plt
# from io import BytesIO
# import base64
# from django.shortcuts import render
# from .models import Portfolio

# def bar_graph(request):
#     # Retrieve the latest portfolio record (you can adjust the filtering as needed)
#     portfolio = Portfolio.objects.latest('portfolio_date')
    
#     # Prepare data for the graph
#     labels = ['Equity', 'Debt', 'Combined Others']
#     values = [portfolio.equity_total, portfolio.debt_total, portfolio.other_total]
    
#     # Create a bar graph using Matplotlib
#     plt.figure(figsize=(8, 6))
#     bars = plt.bar(labels, values, color=['blue', 'green', 'orange'])
#     plt.xlabel("Categories")
#     plt.ylabel("Market Value (Rs. In Lakhs)")
#     plt.title("Market Value by Category")
    
#     # Optionally, add the value on top of each bar
#     for bar in bars:
#         height = bar.get_height()
#         plt.text(
#             bar.get_x() + bar.get_width() / 2.0, 
#             height, 
#             f'{height:.2f}', 
#             ha='center', 
#             va='bottom'
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

import matplotlib.pyplot as plt
from io import BytesIO
import base64
from django.shortcuts import render
from .models import Portfolio

def bar_graph(request):
    # Retrieve the latest portfolio record (you can adjust the filtering as needed)
    portfolio = Portfolio.objects.latest('portfolio_date')
    
    # Prepare data for the graph
    total_value = portfolio.equity_total + portfolio.debt_total + portfolio.other_total  # Total market value
    
    # Calculate percentages for each category
    equity_percentage = (portfolio.equity_total / total_value) * 100
    debt_percentage = (portfolio.debt_total / total_value) * 100
    other_percentage = (portfolio.other_total / total_value) * 100
    
    # Prepare the labels and values for the bar graph
    labels = ['Equity', 'Debt', 'Combined Others']
    values = [equity_percentage, debt_percentage, other_percentage]
    
    # Create a horizontal bar graph using Matplotlib
    plt.figure(figsize=(8, 6))
    bars = plt.barh(labels, values, color=['blue', 'green', 'orange'])
    plt.xlabel("Percentage (%)")
    plt.ylabel("Categories")
    plt.title("Portfolio Market Value Distribution by Category")
    
    # Optionally, add the percentage value on the bars
    for bar in bars:
        width = bar.get_width()
        plt.text(
            width, 
            bar.get_y() + bar.get_height() / 2.0, 
            f'{width:.2f}%', 
            ha='left', 
            va='center'
        )
    
    # Save the figure to a BytesIO buffer
    buffer = BytesIO()
    plt.savefig(buffer, format='png')
    buffer.seek(0)
    image_png = buffer.getvalue()
    buffer.close()
    
    # Encode the image to base64 string
    graph = base64.b64encode(image_png).decode('utf-8')
    plt.close()
    
    # Pass the graph to the template
    return render(request, 'bar_graph.html', {'graph': graph})
