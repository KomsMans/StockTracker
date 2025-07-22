import pandas as pd
import os
def create():
    username = input("Name?: ").strip()
    while not username:
        print("Error: Name cannot be empty.")
        username = input("Name?: ").strip()

    filename = os.path.join("data", f"{username}.xlsx")
    records = []

    while True:
        stock = input("Stock name? (or 'done' to finish): ").strip()
        if stock.lower() == 'done':
            break
        if not stock:
            print("Error: Stock name cannot be empty.")
            continue

        while True:
            qty = input("Quantity: ").strip()
            try:
                qty = int(qty)
                if qty <= 0:
                    print("Error: Quantity must be positive.")
                else:
                    break
            except ValueError:
                print("Error: Quantity must be a whole number.")

        while True:
            avg_price = input("Average Price: ").strip()
            try:
                avg_price = float(avg_price)
                if avg_price <= 0:
                    print("Error: Price must be positive.")
                else:
                    break
            except ValueError:
                print("Error: Price must be a number.")

        records.append({'Stock Name': stock, 'Quantity': qty, 'Average Price': avg_price})

    try:
        os.makedirs("data", exist_ok=True)
        pd.DataFrame(records).to_excel(filename, index=False)
        print(f"Saved to {filename}")
    except Exception as e:
        print(f"Error saving file: {e}")

def view():
    files = [f for f in os.listdir("data") if f.endswith(".xlsx")]
    for i, file in enumerate(files):
        print(f"{i+1}. {file}")
    choice = int(input("Choose: ")) - 1
    filepath = os.path.join("data", files[choice])
    df = pd.read_excel(filepath)
    print("\nCurrent Data:")
    print(df)
    if input("Edit? (y/n): ").lower() == 'y':
        row = int(input("Row to edit: "))
        col = input("Column to edit (Stock Name/Quantity/Average Price): ")
        new_value = input("New value: ")
        if col == 'Quantity':
            new_value = int(new_value)
        elif col == 'Average Price':
            new_value = float(new_value)
        df.at[row, col] = new_value
        df.to_excel(filepath, index=False)
        print("Updated!")

def summary():
    if not os.path.exists("data"):
        print("Error: 'data' folder not found. No data to summarize.")
        return

    combined = []
    for file in os.listdir("data"):
        if file.endswith('.xlsx'):
            filepath = os.path.join("data", file)
            try:
                df = pd.read_excel(filepath)
                if not df.empty:
                    combined.append(df)
            except Exception as e:
                print(f"Error reading {file}: {e}")

    if not combined:
        print("No valid data found to summarize.")
        return

    all_data = pd.concat(combined, ignore_index=True)
    all_data["TotalValue"] = all_data["Quantity"] * all_data["Average Price"]
    summary = all_data.groupby("Stock Name").agg({
        "Quantity": "sum",
        "TotalValue": "sum"
    }).reset_index()
    summary["Weighted Average Price"] = summary["TotalValue"] / summary["Quantity"]
    summary = summary[["Stock Name", "Quantity", "Weighted Average Price"]]

    os.makedirs("output", exist_ok=True)
    summary_path = os.path.join("output", "summary.xlsx")
    try:
        summary.to_excel(summary_path, index=False)
        print("\nSummary Generated:")
        print(summary)
    except Exception as e:
        print(f"Error saving summary: {e}")

def main_menu():
    while True:
        print("1. Create New Spreadsheet")
        print("2. View/Edit Existing Spreadsheet")
        print("3. Generate Final Summary")
        print("4. Exit")

        choice = input("Choose an option: ")

        if choice == '1':
            create()
        elif choice == '2':
            view()
        elif choice == '3':
            summary()
        elif choice == '4':
            print("bye")
            break
        else:
            print("Error: Enter a number between 1 and 4.")

main_menu()