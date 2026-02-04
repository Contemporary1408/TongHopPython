#Dùng để xuất thành từng file excel đối với mỗi giá trị lọc của 1 cột trong bảng
import pandas as pd

# Load your Excel file into a DataFrame
input_file = r"D:\Labs\excelgen\Shipment List1.xlsx"  # Sửa lại tên file
df = pd.read_excel(input_file)

# Specify the column to filter on
filter_column = 'STOCK POINT'  # Sửa lại tên cột

# Get all unique values in the specified column
unique_values = df[filter_column].unique()

# Create a new Excel file for each unique value
for value in unique_values:
    filtered_df = df[df[filter_column] == value]
    
    # Define output file name based on the unique value
    output_file = f'D:/Labs/excelgen/filtered_output_{value}.xlsx'  # Chỗ này cần khai báo đường dẫn cụ thể thì mới lưu file được
    filtered_df.to_excel(output_file, index=False)

    print(f'Saved filtered data for "{value}" to {output_file}')
