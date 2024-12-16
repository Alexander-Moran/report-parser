import pandas as pd

'''
This title, date implementation may work, but may need to be added
to the parsing paragraph loop.
'''

class reportParser:
    def __init__(self):
        self.tables_data = []

    def report_parser(self, document):
        title = document.paragraphs[0].text
        date = document.paragraphs[2].text.lower().split("date:")[-1].strip()

        # Initialize analyst as an empty string in case 'analyst:' is not found
        analyst = ""
        for paragraph in document.paragraphs:
            if paragraph.text[:8].lower() == 'analyst:':
                analyst = paragraph.text[9:].lower()

        #print(f"Report: {title}\nDate: {date}\nAnalyst: {analyst}")

        # Iterate over tables and extract rows
        for table in document.tables:
            table_content = []

            # Extract cell text as list format
            for row in table.rows:
                for cell in row.cells:
                    print(cell.text.hyperlink)

                row_data = [cell.text for cell in row.cells]
                table_content.append(row_data)

            # Convert table to DataFrame and set document name as header
            if table_content:  # Ensure the table is not empty
                if len(table_content) > 1:
                    headers = table_content[0]  # First row as headers
                    data = table_content[1:]  # Remaining rows as data
                else:
                    headers = [f"Column {i + 1}" for i in range(len(table_content[0]))]
                    data = []

                # Create DataFrame for each table
                df = pd.DataFrame(data, columns=headers)
                df['Report'] = title
                df['Date'] = date
                df['Analyst'] = analyst
                self.tables_data.append(df)

    def csv_export(self):
        if self.tables_data:
            combined_df = pd.concat(self.tables_data, ignore_index=True)
            combined_df.to_csv('final_report.csv', index=False)
            print("Exported to final_report.csv")


