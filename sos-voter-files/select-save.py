from csv import reader, writer
import io

file_name = "base-r.csv"

            
with open(file_name) as csv_in, open ("result.csv", 'w', newline='') as csv_out:
    csv_reader = reader(csv_in)
    csv_writer = writer(csv_out)
    
    for row in csv_reader:
        if row[6] == '40':
            csv_writer.writerow(row)