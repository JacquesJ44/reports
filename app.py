import pandas as pd
import pdfkit
import os
from datetime import datetime

# CSV files downloaded from Bicoms must be opened in Excel first and converted. Then save the file again (as CSV). The files are then ready to be processed.

directory = r"C:\Users\Jacques\Documents\vsc\bicoms_reports\downloads"
# Columns for the pandas dataframe
columns = ["From", "To", "Date/Time", "Call Duration", "Rating Duration", "Cost", "Status", "ID", "Caller ID"]
# Options for creating the PDF file
options = {    
         'page-size': 'Letter',
         'margin-top': '0mm',
         'margin-right': '0mm',
         'margin-bottom': '0mm',
         'margin-left': '0mm'
      }
count = 0

# Read the CSV files into a pandas dataframe
for file in os.listdir(directory):
   if file.endswith(".csv"):
      rows_html = ""
      count += 1
      path = os.path.join(directory, file)
      print(path)
      
      # Read the CSV file, set the column names, and convert the dataframe to an HTML table. If the file does not have valid data then insert the placeholder html_table
      try:
         df = pd.read_csv(path, encoding='utf-16', sep='\t', header=None)
         df.columns = columns
         df[['Date', 'Time']] = df['Date/Time'].str.split(' ', expand=True)
         df = df.drop(columns=['Date/Time'])
         # print(df.head())
         
         # Iterate over the rows of the dataframe, finding calls made between 17:00 and 06:00
         for _,row in df.iterrows():
            time_str = row['Time']
            # print(time_str)
            time_obj = datetime.strptime(time_str, '%H:%M').time()
            if (time_obj >= datetime.strptime("17:00", "%H:%M").time() or
               time_obj <= datetime.strptime("06:00", "%H:%M").time()):
               row_style = 'style="color: red;"'  # Red text for night hours
            else:
               row_style = ""

            # Build the HTML row
            row_html = f"<tr {row_style}>"
            for value in row:
               row_html += f"<td>{value}</td>"
            row_html += "</tr>"
            rows_html += row_html
         
         # Convert the dataframe to an HTML table
         # html_table = df.to_html()

         # Add the HTML rows into a HTML template
         html_template = f"""
         <html>
            <head>
                  <style>
                     table {{
                        width: 100%;
                        border-collapse: collapse;
                     }}
                     table, th, td {{
                        border: 1px solid black;
                        padding: 8px;
                        text-align: left;
                     }}
                     th {{
                        background-color: #f2f2f2;
                     }}
                  </style>
            </head>
            <body>
               <h1>{file.strip(".csv")}</h1>
                  <table>   
                     <thead>
                     <tr>
                        {''.join(f'<th>{col}</th>' for col in df.columns)}
                     </tr>
                     </thead>
                     <body>
                        {rows_html}
                     </body>
                  </table>
         </html>
         """
      except Exception as e:
         # print(f"Error reading file with pandas: {e}")
         # Fallback HTML in case of an error
         html_template = f"""
         <html>
            <head><title>No Data Found</title></head>
            <body>
                  <h1>{file.strip(".csv")}</h1>
                  <table border="1">
                     <tr>{''.join(['<td>N/A</td>' for _ in range(10)])}</tr>
                  </table>
            </body>
         </html>
         """
         
      # Print the HTML table to the console
      # print(html_table)

      # Save the HTML table to a PDF file
      file = file.strip(".csv")
      print(f"Creating {file}.pdf...{count} of {len(os.listdir(directory))}")
      wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
      pdfkit.configuration(wkhtmltopdf=wkhtmltopdf)
      pdfkit.from_string(html_template, f'{file}.pdf', options=options)
      print(f"{file}.pdf created.")
