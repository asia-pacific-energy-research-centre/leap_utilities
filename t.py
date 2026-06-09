import sys
sys.path.insert(0, ".")
import codebase.results_supply_link_workflow as rsw
rsw.USE_AGGREGATED_DEMAND_AS_DUMMY = True
from codebase.results_supply_link_workflow import load_results_demand_table
demand = load_results_demand_table(economies=["05_PRC", "20_USA"])
print("Shape:", demand.shape)
print("Economies:", sorted(demand["economy"].unique()))
