import pandas as pd
class reportParser:
    def __init__(self):
        self.final_report = pd.DataFrame(columns=["Subject", "Sentiment", "Content", "Source",
                                                  "Report", "Date", "Analyst"])
    def report_parser(self, document):
        # Extract title, date, and analyst information
        title = document.paragraphs[0].text
        date = document.paragraphs[2].text.lower().split("date:")[-1].strip()
        records = []


        # Initialize analyst as an empty string in case 'analyst:' is not found
        analyst = ""
        for paragraph in document.paragraphs:
            if paragraph.text[:8].lower() == 'analyst:':
                analyst = paragraph.text[9:].lower()

        # Iterate over tables and extract rows
        for table in document.tables:


            # Extract cell text and check for hyperlinks in each cell
            for row in table.rows:

                subject = row.cells[0].text

                # Postive Sentiment
                if len(row.cells) > 3:
                    for paragraph in row.cells[1].paragraphs:
                        if paragraph.hyperlinks:
                            link = paragraph.hyperlinks[0].address

                        paragraph_format = paragraph.paragraph_format

                        if paragraph_format.left_indent:
                            content = paragraph.text
                            records.append([subject, "Positive", content, link])
                            #print(f"Subject: {subject}\nDate: {date}\nSentiment: Positive\n"
                             #     f"Report: {content}\nSource: {link}")

                # Negative Sentiment
                for paragraph in row.cells[-2].paragraphs:
                    if paragraph.hyperlinks:
                        link = paragraph.hyperlinks[0].address

                    paragraph_format = paragraph.paragraph_format

                    if paragraph_format.left_indent:
                        content = paragraph.text
                        records.append([subject, "Negative", content, link])
                        #print(f"Subject: {subject}\nDate: {date}\nSentiment: Negative\n"
                         #     f"Report: {content}\nSource: {link}")

                # Threats
                for paragraph in row.cells[-1].paragraphs:
                    if paragraph.hyperlinks:
                        link = paragraph.hyperlinks[0].address

                    paragraph_format = paragraph.paragraph_format

                    if paragraph_format.left_indent:
                        content = paragraph.text
                        records.append([subject, "Threat", content, link])
                        #print(f"Subject: {subject}\nDate: {date}\nSentiment: Threat\n"
                         #     f"Report: {content}\nSource: {link}")

        # After processing the document, print all hyperlinks
        df = pd.DataFrame(records, columns=["Subject", "Sentiment", "Content", "Source"])
        df['Report'] = title
        df['Date'] = date
        df['Analyst'] = analyst
        self.final_report = pd.concat([self.final_report, df], ignore_index=True)


    def csv_export(self):
        self.final_report.to_csv('final_report.csv', index=False)
        print("Exported to final_report.csv")

