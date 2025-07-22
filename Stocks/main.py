import pandas as pd
import os
def create():
    username = input("Name?: ")
    filename = os.path.join("data", f"{username}.xlsx")
    records = []
    while True:
        stock = input("stock name? or done?: ")
        if stock == "done": 
            break
        qty = input("qty: ")
        avg_price = float(input("Average Price: "))
        records.append({'Stock Name': stock, 'Quantity': qty, 'Average Price': avg_price})
    pd.DataFrame(records).to_excel(filename)
    print("saved to {filename}")

def view():
    files = [f for f in os.listdir("data") if f.endswith(".xlsx")]
    for i, file in enumerate(files):
        print(f"{i+1}. {file}")
    choice = int(input("Choose: ")) - 1
    filepath = os.path.join("data", files[choice])
    df = pd.read_excel(filepath)
    print("\nCurrent Data:")
    print(df)

def summary():
    combined = []
    for file in os.listdir("data"):
        if file.endswith('.xlsx'):
            filepath = os.path.join("data", file)
            df = pd.read_excel(filepath)
            combined.append(df)
    all_data = pd.concat(combined, ignore_index=True)
    all_data["TotalValue"] = all_data["Quantity"] * all_data["Average Price"]
    summary = all_data.groupby("Stock Name").agg({
        "Quantity": "sum",
        "TotalValue": "sum"
    }).reset_index()
    summary["Weighted Average Price"] = summary["TotalValue"] / summary["Quantity"]
    summary = summary[["Stock Name", "Quantity", "Weighted Average Price"]]

    summary_path = os.path.join("output", "summary.xlsx")
    summary.to_excel(summary_path, index=False)
    print(summary)

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

main_menu()