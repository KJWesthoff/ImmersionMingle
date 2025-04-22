import webbrowser

import pandas as pd
from IPython.display import display
from tabulate import tabulate

## Inputs
my_name = "Karl-Johan Westhoff"

## Read all sheets
xls = pd.ExcelFile('ImmersionMingle.xlsx')
Immersion_names = xls.parse("Participants")

# Initialize a dictionary to track classes per name
name_to_classes = {name: [] for name in Immersion_names['Name']}

## Loop through the sheets and collect class info
for sheet in xls.sheet_names:
    if sheet == "Participants":
        continue  # Skip this sheet

    df = xls.parse(sheet)

    # Get the class of my_name in this sheet
    if my_name not in df['Name'].values:
        continue  # Skip if my_name not in this sheet

    my_class = df.loc[df['Name'] == my_name, 'Class'].values[0]

    # Find Immersion participants from the same class
    match = df.loc[(df['Class'] == my_class) & (df['Immersion'] == 'Immersion')]

    # Update class lists per matching name
    for name in match['Name']:
        if name in name_to_classes:
            name_to_classes[name].append(my_class)

# Add a new column with the list of classes
Immersion_names['Classes'] = Immersion_names['Name'].apply(lambda n: name_to_classes.get(n, []))


# Custom CSS with alternating row colors
css = """
<style>
table.dataframe {
    border-collapse: collapse;
    width: 100%;
    font-family: Arial, sans-serif;
}
table.dataframe th, table.dataframe td {
    border: 1px solid #ccc;
    padding: 8px;
    text-align: left;
}
table.dataframe th {
    background-color: #f5f5f5;
}
table.dataframe tbody tr:nth-child(odd) {
    background-color: #f9f9f9;
}
table.dataframe tbody tr:nth-child(even) {
    background-color: #ffffff;
}
</style>
"""

# Generate styled HTML
html = css + Immersion_names.sort_values("Classes", ascending=False).style.set_table_attributes('class="dataframe"').to_html()

# Save and open in browser
with open("ImmersionMingle.html", "w", encoding="utf-8") as f:
    f.write(html)

webbrowser.open("ImmersionMingle.html")





