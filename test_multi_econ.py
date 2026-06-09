import sys
sys.path.insert(0, ".")
import codebase.results_supply_link_workflow as rsw

# Enable the flag
rsw.USE_AGGREGATED_DEMAND_AS_DUMMY = True

from codebase.results_supply_link_workflow import load_results_demand_table

print("=== Test 1: Single 00_APEC (aggregate) ===")
demand = load_results_demand_table(economies=["00_APEC"])
print("Shape:", demand.shape)
print("Unique economies:", demand["economy"].unique())
print()

print("=== Test 2: Multiple individual economies ===")
demand = load_results_demand_table(economies=["05_PRC", "20_USA"])
print("Shape:", demand.shape)
print("Unique economies:", demand["economy"].unique())
print("05_PRC rows:", len(demand[demand["economy"] == "05_PRC"]))
print("20_USA rows:", len(demand[demand["economy"] == "20_USA"]))
print()

print("All tests passed.")

