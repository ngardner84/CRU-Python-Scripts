import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk
import chardet
from tkinter.filedialog import askopenfilename
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# Get the current date to include in the filename
current_date = datetime.now().strftime('%Y-%m-%d')
pdf_filename = f"{current_date} CRU Attribution Analysis.pdf"

# Set up the file selection dialog
tk.Tk().withdraw()  # We don't want a full GUI, so keep the root window from appearing
filename = askopenfilename()  # Show an "Open" dialog box and return the path to the selected file

if not filename:
    print("No file selected, exiting program.")
else:
    # Load the data from CSV
    with open(filename, 'rb') as f:
        result = chardet.detect(f.read())
    encoding = result['encoding']

    data = pd.read_csv(filename, encoding=encoding)

    # Check if "UTM Content" and "Source" exist and convert to string
    if 'UTM Content' not in data.columns or 'Source' not in data.columns:
        print("Required columns do not exist in the file.")
    else:
        data['UTM Content'] = data['UTM Content'].astype(str)
        # Filter data for specific event types
        event_types = ['Appointment Booked', 'Form Submitted', 'Incoming Call']
        filtered_data = data[data['Event Type'].isin(event_types)]

        # Create a column to check for numeric values in 'UTM Content'
        data['Has Numeric UTM Content'] = data['UTM Content'].str.contains(r'\d', na=False)
        numeric_utm_content_entries = data['Has Numeric UTM Content'].sum()
        total_entries = len(data)
        ratio_numeric_utm_content = numeric_utm_content_entries / total_entries

        # Calculate the percentage of events from social media
        social_media_entries = filtered_data[filtered_data['Source'] == 'Social media'].shape[0]
        ratio_social_media = social_media_entries / len(filtered_data) if len(filtered_data) > 0 else 0

        # Count the occurrences of each event type in the filtered data
        event_counts = filtered_data['Event Type'].value_counts()

        # Create plots and save as images
        plt.figure(figsize=(22, 6))
        sns.barplot(x=event_counts.index, y=event_counts.values)
        plt.title('Count of Specific Event Types')
        plt.xlabel('Event Type')
        plt.ylabel('Count')
        plt.xticks(rotation=45)
        plt.savefig('event_types.png')
        plt.close()

        plt.figure(figsize=(22, 6))
        sns.countplot(x='Has Numeric UTM Content', data=data)
        plt.title('Entries with and without Numeric UTM Content')
        plt.xlabel('Contains Numeric UTM Content')
        plt.ylabel('Count')
        plt.xticks([0, 1], ['No', 'Yes'])
        plt.savefig('numeric_utm_content.png')
        plt.close()

        # Pie chart for the ratio of Social Media vs Other Sources
        plt.figure(figsize=(8, 8))
        labels = ['Social Media', 'Other Sources']
        sizes = [social_media_entries, len(filtered_data) - social_media_entries]
        colors = ['#ff9999','#66b3ff']
        plt.pie(sizes, colors=colors, labels=labels, autopct='%1.1f%%', startangle=90)
        plt.axis('equal')
        plt.title('Ratio of Social Media to Other Sources')
        plt.savefig('social_media_ratio.png')
        plt.close()

        # Use a dictionary to track names and event types
        name_event_map = {}
        for index, row in filtered_data.iterrows():
            contact = row['Contact']
            created_at = row['Created At']
            event_type = row['Event Type']
            source = row['Source']
            entry = f"{contact}, {created_at}, Source: {source} - {event_type}"
            if contact not in name_event_map:
                name_event_map[contact] = []
            name_event_map[contact].append(entry)

        # Now create a single PDF with both images and text
        c = canvas.Canvas(pdf_filename, pagesize=letter)
        c.drawString(72, 750, "CRU Attribution Analysis Report")
        
        # Add pie chart to the PDF at a lowered position
        c.drawImage(ImageReader('social_media_ratio.png'), 72, 450, width=300, height=300)

        # Add other images to the PDF at adjusted positions
        c.drawImage(ImageReader('event_types.png'), 72, 275, width=300, height=200)
        c.drawImage(ImageReader('numeric_utm_content.png'), 72, 75, width=400, height=200)
        
        c.drawString(72, 65, f"Percentage of events that were generated from ads: {(ratio_numeric_utm_content)*100:.2f}%")
        c.drawString(72, 45, f"Percentage of events from social media: {(ratio_social_media)*100:.2f}%")
        c.drawString(72, 25, "Contact Names, Created At, Source and Event Types for Specified Events:")
        
        # Sort the names and print them
        sorted_names = sorted(name_event_map.items())
        y = 20
        last_name = None
        for name, entries in sorted_names:
            if last_name is not None:  # Add extra space when a new name is printed
                y -= 30
            for entry in entries:
                if y < 40:
                    c.showPage()
                    y = 800
                c.drawString(72, y, entry)
                y -= 20
            last_name = name
        
        c.save()  # Save the PDF

        print("PDF generated successfully at:", pdf_filename)

        # Optionally delete the temporary images
        import os
        os.remove('event_types.png')
        os.remove('numeric_utm_content.png')
        os.remove('social_media_ratio.png')
