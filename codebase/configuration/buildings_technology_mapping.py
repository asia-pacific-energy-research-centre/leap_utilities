"""Buildings technology mappings used by buildings fuel remap."""

# Minimal structure requested for manual mapping edits.
BUILDINGS_TECHNOLOGY_TARGET_PRODUCTS = {
    'RESIDENTIAL PER HOUSEHOLD': {
        'Cooking': {
            'Electric Stove': {
                'target_products': ['17 Electricity'],
            },
            'Electric stove': {
                'target_products': ['17 Electricity'],
            },
            'Natural gas stove': {
                'target_products': ['08.01 Natural gas'],
            },
            'City gas stove': {
                'target_products': ['08.01 Natural gas'],
            },
            'LPG Stove': {
                'target_products': ['07.09 LPG'],
            },
            'Kerosene stove': {
                'target_products': ['07.06 Kerosene'],
            },
            'Coal stove': {
                'target_products': ['01 Coal'],
            },
            'Biomass': {
                'target_products': ['15 Solid biomass'],
            },
            'Others': {
                'target_products': ['07.07 Gas/diesel oil', '16 Others', '12 Solar', '12.99 Solar nonspecified', '20 Total Renewables', '21 Modern renewables'],
            },
            'Cooking appliances': {
                'target_products': [],
            },
        },
        'Appliances': {
            'Dishwasher': {
                'target_products': ['17 Electricity'],
            },
            'Washer': {
                'target_products': ['17 Electricity'],
            },
            'Refrigeration': {
                'target_products': ['17 Electricity'],
            },
            'Electric dryer': {
                'target_products': ['17 Electricity'],
            },
            'Gas dryer': {
                'target_products': ['08.01 Natural gas'],
            },
            'Other appliances_incl electronics etc': {
                'target_products': [],
            },
            'Others': {
                'target_products': [],
            },
        },
        'Lighting': {
            'LED': {
                'target_products': ['17 Electricity'],
            },
            'Conventional lighting': {
                'target_products': ['17 Electricity'],
            },
            'Kerosene lamps': {
                'target_products': ['07.06 Kerosene'],
            },
        },
        'Water Heating': {
            'Electric water heating': {
                'target_products': ['17 Electricity'],
            },
            'Gas water heating': {
                'target_products': ['08.01 Natural gas'],
            },
            'District Heat': {
                'target_products': ['18 Heat'],
            },
            'Geothermal': {
                'target_products': ['11 Geothermal'],
            },
            'Solar thermal': {
                'target_products': ['12 Solar'],
            },
            'Heat pump': {
                'target_products': ['17 Electricity'],
            },
            'Heat Pump': {
                'target_products': [],
            },
        },
    },
    'RESIDENTIAL PER SQUARE METRE': {
        'Space Heating': {
            'Electric heating': {
                'target_products': ['17 Electricity'],
            },
            'Gas boiler_furnace': {
                'target_products': ['08.01 Natural gas'],
            },
            'Oil boiler_furnace': {
                'target_products': ['07.08 Fuel oil'],
            },
            'LPG boiler': {
                'target_products': ['07.09 LPG'],
            },
            'District heat': {
                'target_products': ['18 Heat'],
            },
            'Heat Pump': {
                'target_products': [],
            },
            'Heat pump': {
                'target_products': ['17 Electricity'],
            },
            'Biomass': {
                'target_products': ['15 Solid biomass'],
            },
            'Room heater': {
                'target_products': [],
            },
            'Others': {
                'target_products': [],
            },
        },
        'Space Cooling': {
            'Central AC': {
                'target_products': ['17 Electricity'],
            },
            'Room AC': {
                'target_products': ['17 Electricity'],
            },
            'Split AC': {
                'target_products': ['17 Electricity'],
            },
            'Fans': {
                'target_products': ['17 Electricity'],
            },
            'Heat pump': {
                'target_products': ['17 Electricity'],
            },
        },
    },
    'SERVICES PER FLOORSPACE': {
        'Space Heating': {
            'Electric heating': {
                'target_products': ['17 Electricity'],
            },
            'Gas boiler furnace': {
                'target_products': ['08.01 Natural gas'],
            },
            'Oil boiler furnace': {
                'target_products': ['07.08 Fuel oil'],
            },
            'District heat': {
                'target_products': ['18 Heat'],
            },
            'Heat Pump': {
                'target_products': [],
            },
            'Heat pump': {
                'target_products': ['17 Electricity'],
            },
            'Biomass': {
                'target_products': ['15 Solid biomass'],
            },
            'Coal': {
                'target_products': ['01 Coal'],
            },
            'Others': {
                'target_products': ['07.07 Gas/diesel oil', '16 Others', '12 Solar', '12.99 Solar nonspecified', '20 Total Renewables', '21 Modern renewables'],
            },
        },
        'Space Cooling': {
            'Central AC': {
                'target_products': ['17 Electricity'],
            },
            'Room AC': {
                'target_products': ['17 Electricity'],
            },
            'Fans': {
                'target_products': ['17 Electricity'],
            },
            'Heat Pump': {
                'target_products': [],
            },
            'Heat pump': {
                'target_products': [],
            },
        },
        'Lighting': {
            'Conventional': {
                'target_products': ['17 Electricity'],
            },
            'LED': {
                'target_products': ['17 Electricity'],
            },
        },
    },
    'SERVICES PER SRV GDP': {
        'Cooking': {
            'Electric stove': {
                'target_products': [],
            },
            'Natural gas stove': {
                'target_products': ['08.01 Natural gas'],
            },
            'LPG stove': {
                'target_products': ['07.09 LPG'],
            },
            'Biomass': {
                'target_products': ['15 Solid biomass'],
            },
            'Others': {
                'target_products': ['07.07 Gas/diesel oil', '16 Others', '12 Solar', '12.99 Solar nonspecified', '20 Total Renewables', '21 Modern renewables'],
            },
            'Electric Stove': {
                'target_products': [],
            },
        },
        'Water Heating': {
            'Electric': {
                'target_products': ['17 Electricity'],
            },
            'Gas heater': {
                'target_products': ['08.01 Natural gas'],
            },
            'Distrcit heat': {
                'target_products': [],
            },
            'Heat pump': {
                'target_products': ['17 Electricity'],
            },
            'Solar thermal': {
                'target_products': ['12 Solar'],
            },
            'Biomass': {
                'target_products': ['15 Solid biomass'],
            },
            'Diesel heater?': {
                'target_products': ['07.07 Gas/diesel oil'],
            },
            'Others': {
                'target_products': ['16 Others', '12.99 Solar nonspecified', '20 Total Renewables', '21 Modern renewables'],
            },
            'District heat': {
                'target_products': [],
            },
        },
        'Equipment_appliances': {
            'Refrigeration': {
                'target_products': ['17 Electricity'],
            },
            'Others': {
                'target_products': ['17 Electricity', '16 Others'],
            },
        },
    },
}

# Extended structure used by remap logic (preserves alias/split behavior).
BUILDINGS_TECHNOLOGY_MAPPING = {
    'RESIDENTIAL PER HOUSEHOLD': {
        'Cooking': {
            'Electric Stove': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Electric Stove',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Electric stove': {
                'target_products': [],
                'canonical_technology': 'Electric Stove',
                'mapping_mode': 'alias',
                'esto_flow': '16.02 Residential',
                'notes': 'Alias to Electric Stove',
            },
            'Natural gas stove': {
                'target_products': ['08.01 Natural gas'],
                'canonical_technology': 'Natural gas stove',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'City gas stove': {
                'target_products': ['08.01 Natural gas'],
                'canonical_technology': 'City gas stove',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'LPG Stove': {
                'target_products': ['07.09 LPG'],
                'canonical_technology': 'LPG Stove',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'LPG stove': {
                'target_products': [],
                'canonical_technology': 'LPG Stove',
                'mapping_mode': 'alias',
                'esto_flow': '16.02 Residential',
                'notes': 'Alias to LPG Stove',
            },
            'Kerosene stove': {
                'target_products': ['07.06 Kerosene'],
                'canonical_technology': 'Kerosene stove',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Coal stove': {
                'target_products': ['01 Coal'],
                'canonical_technology': 'Coal stove',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Biomass': {
                'target_products': ['15 Solid biomass'],
                'canonical_technology': 'Biomass',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Others': {
                'target_products': ['07.07 Gas/diesel oil', '16 Others', '12 Solar', '12.99 Solar nonspecified', '20 Total Renewables', '21 Modern renewables'],
                'canonical_technology': 'Others',
                'mapping_mode': 'split_base_year',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Cooking appliances': {
                'target_products': [],
                'canonical_technology': 'Others',
                'mapping_mode': 'alias',
                'esto_flow': '16.02 Residential',
                'notes': 'Alias to Others',
            },
        },
        'Appliances': {
            'Dishwasher': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Dishwasher',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Washer': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Washer',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Refrigeration': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Refrigeration',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Electric dryer': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Electric dryer',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Gas dryer': {
                'target_products': ['08.01 Natural gas'],
                'canonical_technology': 'Gas dryer',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Other appliances_incl electronics etc': {
                'target_products': [],
                'canonical_technology': 'Others',
                'mapping_mode': 'alias',
                'esto_flow': '16.02 Residential',
                'notes': 'Alias to Others',
            },
            'Others': {
                'target_products': ['17 Electricity', '08.01 Natural gas', '16 Others'],
                'canonical_technology': 'Others',
                'mapping_mode': 'split_base_year',
                'esto_flow': '16.02 Residential',
                'notes': "Added for alias 'Other appliances_incl electronics etc'",
            },
        },
        'Lighting': {
            'LED': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'LED',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Conventional lighting': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Conventional lighting',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Kerosene lamps': {
                'target_products': ['07.06 Kerosene'],
                'canonical_technology': 'Kerosene lamps',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
        },
        'Water Heating': {
            'Electric water heating': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Electric water heating',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Gas water heating': {
                'target_products': ['08.01 Natural gas'],
                'canonical_technology': 'Gas water heating',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'District Heat': {
                'target_products': ['18 Heat'],
                'canonical_technology': 'District Heat',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Geothermal': {
                'target_products': ['11 Geothermal'],
                'canonical_technology': 'Geothermal',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Solar thermal': {
                'target_products': ['12 Solar'],
                'canonical_technology': 'Solar thermal',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Heat pump': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Heat pump',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Heat Pump': {
                'target_products': [],
                'canonical_technology': 'Heat pump',
                'mapping_mode': 'alias',
                'esto_flow': '16.02 Residential',
                'notes': 'Alias to Heat pump',
            },
        },
    },
    'RESIDENTIAL PER SQUARE METRE': {
        'Space Heating': {
            'Electric heating': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Electric heating',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Gas boiler_furnace': {
                'target_products': ['08.01 Natural gas'],
                'canonical_technology': 'Gas boiler furnace',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Oil boiler_furnace': {
                'target_products': ['07.08 Fuel oil'],
                'canonical_technology': 'Oil boiler furnace',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'LPG boiler': {
                'target_products': ['07.09 LPG'],
                'canonical_technology': 'LPG boiler',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'District heat': {
                'target_products': ['18 Heat'],
                'canonical_technology': 'District heat',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Heat Pump': {
                'target_products': [],
                'canonical_technology': 'Heat pump',
                'mapping_mode': 'alias',
                'esto_flow': '16.02 Residential',
                'notes': 'Alias to Heat pump',
            },
            'Heat pump': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Heat pump',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Biomass': {
                'target_products': ['15 Solid biomass'],
                'canonical_technology': 'Biomass',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Room heater': {
                'target_products': [],
                'canonical_technology': 'Others',
                'mapping_mode': 'alias',
                'esto_flow': '16.02 Residential',
                'notes': 'Alias to Others',
            },
            'Others': {
                'target_products': ['17 Electricity', '08.01 Natural gas', '07.08 Fuel oil', '07.09 LPG', '18 Heat', '15 Solid biomass', '16 Others'],
                'canonical_technology': 'Others',
                'mapping_mode': 'split_base_year',
                'esto_flow': '16.02 Residential',
                'notes': "Added for alias 'Room heater'",
            },
        },
        'Space Cooling': {
            'Central AC': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Central AC',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Room AC': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Room AC',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Split AC': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Split AC',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Fans': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Fans',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
            'Heat pump': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Heat pump',
                'mapping_mode': 'direct',
                'esto_flow': '16.02 Residential',
                'notes': '',
            },
        },
    },
    'SERVICES PER FLOORSPACE': {
        'Space Heating': {
            'Electric heating': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Electric heating',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Gas boiler furnace': {
                'target_products': ['08.01 Natural gas'],
                'canonical_technology': 'Gas boiler furnace',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Oil boiler furnace': {
                'target_products': ['07.08 Fuel oil'],
                'canonical_technology': 'Oil boiler furnace',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'District heat': {
                'target_products': ['18 Heat'],
                'canonical_technology': 'District heat',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Heat Pump': {
                'target_products': [],
                'canonical_technology': 'Heat pump',
                'mapping_mode': 'alias',
                'esto_flow': '16.01 Commercial and public services',
                'notes': 'Alias to Heat pump',
            },
            'Heat pump': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Heat pump',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Biomass': {
                'target_products': ['15 Solid biomass'],
                'canonical_technology': 'Biomass',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Coal': {
                'target_products': ['01 Coal'],
                'canonical_technology': 'Coal',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Others': {
                'target_products': ['07.07 Gas/diesel oil', '16 Others', '12 Solar', '12.99 Solar nonspecified', '20 Total Renewables', '21 Modern renewables'],
                'canonical_technology': 'Others',
                'mapping_mode': 'split_base_year',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
        },
        'Space Cooling': {
            'Central AC': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Central AC',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Room AC': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Room AC',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Fans': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Fans',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Heat Pump': {
                'target_products': [],
                'canonical_technology': 'Heat pump',
                'mapping_mode': 'alias',
                'esto_flow': '16.01 Commercial and public services',
                'notes': 'Alias to Heat pump',
            },
            'Heat pump': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Heat pump',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': "Added for alias 'Heat Pump'",
            },
        },
        'Lighting': {
            'Conventional': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Conventional',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'LED': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'LED',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
        },
    },
    'SERVICES PER SRV GDP': {
        'Cooking': {
            'Electric stove': {
                'target_products': [],
                'canonical_technology': 'Electric Stove',
                'mapping_mode': 'alias',
                'esto_flow': '16.01 Commercial and public services',
                'notes': 'Alias to Electric Stove',
            },
            'Natural gas stove': {
                'target_products': ['08.01 Natural gas'],
                'canonical_technology': 'Natural gas stove',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'LPG stove': {
                'target_products': ['07.09 LPG'],
                'canonical_technology': 'LPG stove',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Biomass': {
                'target_products': ['15 Solid biomass'],
                'canonical_technology': 'Biomass',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Others': {
                'target_products': ['07.07 Gas/diesel oil', '16 Others', '12 Solar', '12.99 Solar nonspecified', '20 Total Renewables', '21 Modern renewables'],
                'canonical_technology': 'Others',
                'mapping_mode': 'split_base_year',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Electric Stove': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Electric Stove',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': "Added for alias 'Electric stove'",
            },
        },
        'Water Heating': {
            'Electric': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Electric',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Gas heater': {
                'target_products': ['08.01 Natural gas'],
                'canonical_technology': 'Gas heater',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Distrcit heat': {
                'target_products': [],
                'canonical_technology': 'District heat',
                'mapping_mode': 'alias',
                'esto_flow': '16.01 Commercial and public services',
                'notes': 'Alias typo to District heat',
            },
            'Heat pump': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Heat pump',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Solar thermal': {
                'target_products': ['12 Solar'],
                'canonical_technology': 'Solar thermal',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Biomass': {
                'target_products': ['15 Solid biomass'],
                'canonical_technology': 'Biomass',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Diesel heater?': {
                'target_products': ['07.07 Gas/diesel oil'],
                'canonical_technology': 'Diesel heater',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Others': {
                'target_products': ['16 Others', '12.99 Solar nonspecified', '20 Total Renewables', '21 Modern renewables'],
                'canonical_technology': 'Others',
                'mapping_mode': 'split_base_year',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'District heat': {
                'target_products': ['18 Heat'],
                'canonical_technology': 'District heat',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': "Added for alias 'Distrcit heat'",
            },
        },
        'Equipment_appliances': {
            'Refrigeration': {
                'target_products': ['17 Electricity'],
                'canonical_technology': 'Refrigeration',
                'mapping_mode': 'direct',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
            'Others': {
                'target_products': ['17 Electricity', '16 Others'],
                'canonical_technology': 'Others',
                'mapping_mode': 'split_base_year',
                'esto_flow': '16.01 Commercial and public services',
                'notes': '',
            },
        },
    },
}

__all__ = ['BUILDINGS_TECHNOLOGY_TARGET_PRODUCTS', 'BUILDINGS_TECHNOLOGY_MAPPING']
