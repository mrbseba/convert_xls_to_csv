import pandas as pd

def convert_csv_to_excel(csv_file, xls_file=None, xlsx_file=None, split_xls=True, chunk_size=65000, max_rows=None):
    """
    Convert CSV to both XLS and XLSX with automatic splitting for XLS format.
    Optionally limit conversion to the first 'max_rows' rows.

    Args:
        csv_file (str): Input CSV file path.
        xls_file (str): Output XLS path (will get _part suffixes if split).
        xlsx_file (str): Output XLSX path.
        split_xls (bool): Auto-split large files for XLS format (default True).
        chunk_size (int): Max rows per XLS file (default 65,000).
        max_rows (int): Optional maximum number of rows to convert (default None means no limit).
    """
    try:
        # Read the CSV file with specified settings
        df = pd.read_csv(csv_file, delimiter=';', encoding='latin1')
        total_rows = len(df)
        print(f"📊 Total rows detected: {total_rows:,}")

        # If max_rows is specified and total_rows exceeds it, limit the DataFrame
        if max_rows is not None and total_rows > max_rows:
            df = df.head(max_rows)
            total_rows = len(df)
            print(f"📉 Limiting conversion to first {max_rows:,} rows.")

        # XLSX Conversion (single file)
        if xlsx_file:
            if total_rows > 500001:
                print(f"⛔ XLSX creation skipped: Exceeds row limit")
            else:
                df.to_excel(xlsx_file, index=False, engine='openpyxl')
                print(f"✅ XLSX file created: {xlsx_file}")

        # XLS Conversion with auto-splitting
        if xls_file:
            if total_rows <= 65535:
                # Create single XLS file if under limit
                df.to_excel(xls_file, index=False, engine='xlwt')
                print(f"✅ XLS file created: {xls_file}")
            else:
                if split_xls:
                    base_name = xls_file.rsplit('.', 1)[0]
                    part_num = 0
                    current_row_count = 0
                    # Use chunks to read the CSV file; this supports limiting to max_rows
                    reader = pd.read_csv(csv_file, delimiter=';', encoding='latin1', chunksize=chunk_size)
                    for chunk in reader:
                        if max_rows is not None:
                            # Stop processing if we've already reached the limit
                            if current_row_count >= max_rows:
                                break
                            # Trim the chunk if adding all rows would exceed max_rows
                            if current_row_count + len(chunk) > max_rows:
                                chunk = chunk.head(max_rows - current_row_count)
                        part_num += 1
                        output_name = f"{base_name}_part{part_num}.xls"
                        chunk.to_excel(output_name, index=False, engine='xlwt')
                        print(f"✅ XLS part {part_num} created: {output_name}")
                        current_row_count += len(chunk)
                else:
                    print(f"⛔ XLS creation skipped: {total_rows:,} rows exceeds 65,535 limit (use split_xls=True)")

    except Exception as e:
        print(f"🔥 Error: {str(e)}")

if __name__ == "__main__":
    input_csv = r'C:\me_seba\Promo_Price_99MB.csv'
    
    convert_csv_to_excel(
        csv_file=input_csv,
        # xls_file=r'C:\me_seba\output.xls',  # Uncomment to create XLS file(s)
        xlsx_file=r'C:\me_seba\output.xlsx',
        split_xls=True,      # Auto-split large XLS files
        chunk_size=65000,    # 65,000 rows per XLS file
        max_rows=500000      # Limit conversion to first 500,000 rows
    )
