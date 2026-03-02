# Mappings for previously unmapped ninth data
# These names do not conflict with any already-mapped names

NINTH_UNMAPPED_MAPPINGS = {
    # Fuel mappings
    '15_solid_biomass_unallocated': 'Unallocated solid biomass',
    '16_others_unallocated': 'Unallocated others',
    '17_x_green_electricity': 'Green electricity',
    
    # Sector mappings
    '15_02_02_04_09_lng': 'Heavy truck - LNG (freight)',  # Under freight > heavy truck hierarchy
    '16_01_03_ai_training': 'AI training',  # Under buildings
    '16_01_04_traditional_data_centres': 'Traditional data centres',  # Under buildings
    '18_01_01_coal_power_ccs': 'Coal power CCS (electricity output)',  # Already exists - this is duplicate
    '18_01_02_gas_power_ccs': 'Gas power CCS (electricity output)',  # Already exists - this is duplicate
    '18_01_02_gas_power_h2': 'Gas power H2 (electricity output)',  # Under electricity plants
    '19_01_05_others': 'Other CHP (heat output)',  # Under CHP plants heat output
    '19_02_05_others': 'Other heat plants (heat output)',  # Under heat plants
    '19_02_17_electricity': 'Electric heat plants (heat output)',  # Under heat plants, using electricity as fuel
    '22_demand_supply_discrepancy': 'Demand-supply discrepancy',  # Top-level balance item
}
