import csv

# Read the original CSV
with open('nepremicnine_export.csv', 'r', encoding='utf-8') as infile:
    reader = csv.DictReader(infile, delimiter=';')

    # Define the desired columns in order
    desired_columns = [
        'VrstaObjekta',  # type of object
        'VelikostM2',    # area
        'LetoGradnje',   # year of construction
        'EnergetskiRazred',  # energy class
        'Lokacija',      # municipality/place
        'ZemljisteM2',   # land
        'StSob',         # number of rooms
        'Url'            # URL of the listing
    ]

    # Write to new CSV
    with open('nepremicnine_reorganized.csv', 'w', newline='', encoding='utf-8') as outfile:
        writer = csv.DictWriter(outfile, fieldnames=desired_columns, delimiter=';')
        writer.writeheader()

        for row in reader:
            # Create new row with only desired columns
            new_row = {col: row.get(col, '') for col in desired_columns}
            writer.writerow(new_row)

print("CSV reorganized successfully. New file: nepremicnine_reorganized.csv")
