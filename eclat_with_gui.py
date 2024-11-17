import pandas as pd
from itertools import combinations
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to read transactions from an Excel file
def read_transactions_from_excel(file_path):
    df = pd.read_excel(file_path)
    transactions = df.values.tolist()
    transactions = [[item for item in row if pd.notna(item)] for row in transactions]
    return transactions

# Convert data to vertical format
def convert_to_vertical_format(transactions):
    vertical_data = defaultdict(set)
    for tid, transaction in enumerate(transactions):
        for item in transaction:
            items = item.split(',') if isinstance(item, str) else [item]
            for individual_item in items:
                vertical_data[individual_item].add(tid)
    return vertical_data

# Generate frequent itemsets
def generate_frequent_itemsets(vertical_data, min_support):
    frequent_itemsets = []
    items = list(vertical_data.keys())
    for k in range(1, len(items) + 1):
        candidates = combinations(items, k)
        for candidate in candidates:
            tid_sets = [vertical_data[item] for item in candidate]
            intersection = set.intersection(*tid_sets)
            support = len(intersection)
            if support >= min_support:
                frequent_itemsets.append((candidate, support))
    return frequent_itemsets

# Generate association rules
def generate_association_rules(frequent_itemsets, min_confidence):
    rules = []
    for itemset, support in frequent_itemsets:
        if len(itemset) > 1:
            for i in range(1, len(itemset)):
                for antecedent in combinations(itemset, i):
                    consequent = tuple(set(itemset) - set(antecedent))
                    antecedent_support = next(
                        (sup for items, sup in frequent_itemsets if set(items) == set(antecedent)), 0
                    )
                    confidence = support / antecedent_support if antecedent_support > 0 else 0
                    if confidence >= min_confidence:
                        rules.append((antecedent, consequent, confidence))
    return rules

# Calculate lift
def calculate_lift(rules, frequent_itemsets, total_transactions):
    lifts = []
    for antecedent, consequent, confidence in rules:
        consequent_support = next(
            (sup for items, sup in frequent_itemsets if set(items) == set(consequent)), 0
        )
        lift = confidence / (consequent_support / total_transactions)
        lifts.append((antecedent, consequent, confidence, lift))
    return lifts

# Running the ECLAT Algorithm
def run_eclat(file_path, min_support, min_confidence):
    transactions = read_transactions_from_excel(file_path)
    total_transactions = len(transactions)
    vertical_data = convert_to_vertical_format(transactions)
    frequent_itemsets = generate_frequent_itemsets(vertical_data, min_support)
    rules = generate_association_rules(frequent_itemsets, min_confidence)
    lifts = calculate_lift(rules, frequent_itemsets, total_transactions)

    result_text = ""
    result_text += "Frequent Itemsets:\n"
    for itemset, support in frequent_itemsets:
        result_text += f"Itemset: {itemset}, Support: {support}\n"

    result_text += "\nStrong Association Rules:\n"
    for antecedent, consequent, confidence, lift in lifts:
        result_text += f"Rule: {antecedent} -> {consequent}, Confidence: {confidence:.2f}, Lift: {lift:.2f}\n"

    return result_text

# Tkinter GUI for ECLAT
class EclatGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("ECLAT Algorithm GUI")
        self.geometry("600x400")

        # File selection
        self.file_path = ""
        self.file_button = tk.Button(self, text="Choose Excel File", command=self.choose_file)
        self.file_button.pack(pady=10)

        # Min support input
        self.min_support_label = tk.Label(self, text="Enter Min Support:")
        self.min_support_label.pack()
        self.min_support_entry = tk.Entry(self)
        self.min_support_entry.pack(pady=5)

        # Min confidence input
        self.min_confidence_label = tk.Label(self, text="Enter Min Confidence:")
        self.min_confidence_label.pack()
        self.min_confidence_entry = tk.Entry(self)
        self.min_confidence_entry.pack(pady=5)

        # Run button
        self.run_button = tk.Button(self, text="Run ECLAT", command=self.run_eclat)
        self.run_button.pack(pady=20)

        # Result Text Box
        self.result_text = tk.Text(self, height=10, width=70)
        self.result_text.pack(pady=10)

    def choose_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            self.file_path = file_path

    def run_eclat(self):
        # Get min_support and min_confidence values
        try:
            min_support = int(self.min_support_entry.get())
            min_confidence = float(self.min_confidence_entry.get())

            if not self.file_path:
                raise ValueError("Please choose a file.")

            # Run the ECLAT algorithm
            result = run_eclat(self.file_path, min_support, min_confidence)
            self.result_text.delete(1.0, tk.END)  # Clear the text box
            self.result_text.insert(tk.END, result)  # Display the result

        except Exception as e:
            messagebox.showerror("Error", f"Error: {e}")

# Running the Tkinter Application
if __name__ == "__main__":
    app = EclatGUI()
    app.mainloop()
