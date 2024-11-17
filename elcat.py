import pandas as pd
from itertools import combinations
from collections import defaultdict

file_path = r"C:\Users\arqam\PycharmProjects\ECLAT_PROJECT\Horizontal_Format.xlsx"
# Step 1: Read Transactions Table from Excel
def read_transactions_from_excel(file_path):
    df = pd.read_excel(file_path)
    transactions = df.values.tolist()
    transactions = [[item for item in row if pd.notna(item)] for row in transactions]
    return transactions

# Step 2: Convert Data to Vertical Format
def convert_to_vertical_format(transactions):
    vertical_data = defaultdict(set)
    for tid, transaction in enumerate(transactions):
        for item in transaction:
            items = item.split(',') if isinstance(item, str) else [item]
            for individual_item in items:
                vertical_data[individual_item].add(tid)
    return vertical_data

# Step 3: Generate Frequent Itemsets
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

# Step 4: Generate Association Rules
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

# Step 5: Calculate Lift
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

    print("Frequent Itemsets:")
    for itemset, support in frequent_itemsets:
        print(f"Itemset: {itemset}, Support: {support}")

    print("\nStrong Association Rules:")
    for antecedent, consequent, confidence, lift in lifts:
        print(f"Rule: {antecedent} -> {consequent}, Confidence: {confidence:.2f}, Lift: {lift:.2f}")

# Example Usage
min_support = 2
min_confidence = 0.5
run_eclat(file_path, min_support, min_confidence)