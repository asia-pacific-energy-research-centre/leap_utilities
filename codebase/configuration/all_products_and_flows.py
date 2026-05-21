
# THE FOLLOWING MAPPINGS may be useful for helping agents with understanding what fuels and subfuels exist in ESTO and Ninth datasets.
ESTO_PRODUCT_LIST = [
    '01 Coal',
    '01.01 Coking coal',
    '01.02 Other bituminous coal',
    '01.03 Sub-bituminous coal',
    '01.04 Anthracite',
    '01.05 Lignite',
    '01.99 Coal nonspecified',
    '02 Coal products',
    '02.01 Coke oven coke',
    '02.02 Gas coke',
    '02.03 Coke oven gas',
    '02.04 Blast furnace gas',
    '02.05 Other recovered gases',
    '02.06 Patent fuel',
    '02.07 Coal tar',
    '02.08 BKB/PB',
    '03 Peat',
    '04 Peat products',
    '05 Oil shale and oil sands',
    '06 Crude oil & NGL',
    '06.01 Crude oil',
    '06.02 Natural gas liquids',
    '06.03 Refinery feedstocks',
    '06.04 Additives/  oxygenates',
    '06.05 Other hydrocarbons',
    '07 Petroleum products',
    '07.01 Motor gasoline',
    '07.02 Aviation gasoline',
    '07.03 Naphtha',
    '07.04 Gasoline type jet fuel',
    '07.05 Kerosene type jet fuel',
    '07.06 Kerosene',
    '07.07 Gas/diesel oil',
    '07.08 Fuel oil',
    '07.09 LPG',
    '07.10 Refinery gas (not liquefied)',
    '07.11 Ethane',
    '07.12 White spirit SBP',
    '07.13 Lubricants',
    '07.14 Bitumen',
    '07.15 Paraffin  waxes',
    '07.16 Petroleum coke',
    '07.17 Other products',
    '07.99 PetProd nonspecified',
    '08 Gas',
    '08.01 Natural gas',
    '08.02 LNG',
    '08.03 Gas works gas',
    '08.99 Gas nonspecified',
    '09 Nuclear',
    '10 Hydro',
    '11 Geothermal',
    '12 Solar',
    '12.01 of which: Photovoltaics',
    '12.99 Solar nonspecified',
    '13 Tide, wave, ocean',
    '14 Wind',
    '15 Solid biomass',
    '15.01 Fuelwood & woodwaste',
    '15.02 Bagasse',
    '15.03 Charcoal',
    '15.04 Black liqour',
    '15.05 Other biomass',
    '16 Others',
    '16.01 Biogas',
    '16.02 Industrial waste',
    '16.03 Municipal solid waste (renewable)',
    '16.04 Municipal solid waste (non-renewable)',
    '16.05 Biogasoline',
    '16.06 Biodiesel',
    '16.07 Bio jet kerosene',
    '16.08 Other liquid biofuels',
    '16.09 Other sources',
    '17 Electricity',
    '18 Heat',
    '19 Total',
    '20 Total Renewables',
    '21 Modern renewables',
]

NINTH_FUEL_SUBFUEL_LIST = [
    ('01_coal', '01_01_coking_coal'),
    ('01_coal', '01_05_lignite'),
    ('01_coal', '01_coal_unallocated'),
    ('01_coal', '01_x_thermal_coal'),
    ('01_coal', 'x'),
    ('02_coal_products', 'x'),
    ('03_peat', 'x'),
    ('04_peat_products', 'x'),
    ('05_oil_shale_and_oil_sands', 'x'),
    ('06_crude_oil_and_ngl', '06_01_crude_oil'),
    ('06_crude_oil_and_ngl', '06_02_natural_gas_liquids'),
    ('06_crude_oil_and_ngl', '06_crude_oil_and_ngl_unallocated'),
    ('06_crude_oil_and_ngl', '06_x_other_hydrocarbons'),
    ('06_crude_oil_and_ngl', 'x'),
    ('07_petroleum_products', '07_01_motor_gasoline'),
    ('07_petroleum_products', '07_02_aviation_gasoline'),
    ('07_petroleum_products', '07_03_naphtha'),
    ('07_petroleum_products', '07_06_kerosene'),
    ('07_petroleum_products', '07_07_gas_diesel_oil'),
    ('07_petroleum_products', '07_08_fuel_oil'),
    ('07_petroleum_products', '07_09_lpg'),
    ('07_petroleum_products', '07_10_refinery_gas_not_liquefied'),
    ('07_petroleum_products', '07_11_ethane'),
    ('07_petroleum_products', '07_petroleum_products_unallocated'),
    ('07_petroleum_products', '07_x_jet_fuel'),
    ('07_petroleum_products', '07_x_other_petroleum_products'),
    ('07_petroleum_products', 'x'),
    ('08_gas', '08_01_natural_gas'),
    ('08_gas', '08_02_lng'),
    ('08_gas', '08_03_gas_works_gas'),
    ('08_gas', '08_gas_unallocated'),
    ('08_gas', 'x'),
    ('09_nuclear', 'x'),
    ('10_hydro', 'x'),
    ('11_geothermal', 'x'),
    ('12_solar', '12_01_of_which_photovoltaics'),
    ('12_solar', '12_solar_unallocated'),
    ('12_solar', '12_x_other_solar'),
    ('12_solar', 'x'),
    ('13_tide_wave_ocean', 'x'),
    ('14_wind', 'x'),
    ('15_solid_biomass', '15_01_fuelwood_and_woodwaste'),
    ('15_solid_biomass', '15_02_bagasse'),
    ('15_solid_biomass', '15_03_charcoal'),
    ('15_solid_biomass', '15_04_black_liquor'),
    ('15_solid_biomass', '15_05_other_biomass'),
    ('15_solid_biomass', '15_solid_biomass_unallocated'),
    ('15_solid_biomass', 'x'),
    ('16_others', '16_01_biogas'),
    ('16_others', '16_02_industrial_waste'),
    ('16_others', '16_03_municipal_solid_waste_renewable'),
    ('16_others', '16_04_municipal_solid_waste_nonrenewable'),
    ('16_others', '16_05_biogasoline'),
    ('16_others', '16_06_biodiesel'),
    ('16_others', '16_07_bio_jet_kerosene'),
    ('16_others', '16_08_other_liquid_biofuels'),
    ('16_others', '16_09_other_sources'),
    ('16_others', '16_others_unallocated'),
    ('16_others', '16_x_ammonia'),
    ('16_others', '16_x_efuel'),
    ('16_others', '16_x_hydrogen'),
    ('16_others', 'x'),
    ('17_electricity', 'x'),
    ('17_x_green_electricity', 'x'),
    ('18_heat', 'x'),
]

ESTO_SECTORS = [
    '01 Production',
    '02 Imports',
    '03 Exports',
    '04 International marine bunkers',
    '05 International aviation bunkers',
    '06 Stock changes',
    '07 Total primary energy supply',
    '08 Transfers',
    '08.01 Recycled products',
    '08.02 Interproduct transfers',
    '08.03 Products transferred',
    '08.04 Gas separation',
    '08.99 Transfers nonspecified',
    '09 Total transformation sector',
    '09.01 Main activity producer',
    '09.01.01 Electricity plants',
    '09.01.02 CHP plants',
    '09.01.03 Heat plants',
    '09.02 Autoproducers',
    '09.02.01 Electricity plants',
    '09.02.02 CHP plants',
    '09.02.03 Heat plants',
    '09.03 Heat pumps',
    '09.04 Electric boilers',
    '09.05 Chemical heat for electricity production',
    '09.06 Gas processing plants',
    '09.06.01 Gas works plants',
    '09.06.02 Liquefaction/regasification plants',
    '09.06.03 Natural gas blending plants',
    '09.06.04 Gas-to-liquids plants',
    '09.07 Oil refineries',
    '09.08 Coal transformation',
    '09.08.01 Coke ovens',
    '09.08.02 Blast furnaces',
    '09.08.03 Patent fuel plants',
    '09.08.04 BKB/PB plants',
    '09.08.05 Liquefaction (coal to oil)',
    '09.09 Petrochemical industry',
    '09.10 Biofuels processing',
    '09.11 Charcoal processing',
    '09.12 Non-specified transformation',
    '10 Losses & own use',
    '10.01 Own Use',
    '10.01.01 Electricity, CHP and heat plants',
    '10.01.02 Gas works plants',
    '10.01.03 Liquefaction/regasification plants',
    '10.01.04 Gas-to-liquids plants',
    '10.01.05 Coke ovens',
    '10.01.06 Coal mines',
    '10.01.07 Blast furnaces',
    '10.01.08 Patent fuel plants',
    '10.01.09 BKB/PB plants',
    '10.01.10 Liquefaction plants (Coal to Oil)',
    '10.01.11 Oil refineries',
    '10.01.12 Oil and gas extraction',
    '10.01.13 Pump storage plants',
    '10.01.14 Nuclear industry',
    '10.01.15 Charcoal production plants',
    '10.01.16 Gasification plants for biogases',
    '10.01.17 Non-specified own uses',
    '10.02 Transmission and distribution losses',
    '11 Statistical discrepancy',
    '12 Total final consumption',
    '13 Total final energy consumption',
    '14 Industry sector',
    '14.01 Mining and quarrying',
    '14.02 Construction',
    '14.03 Manufacturing',
    '14.03.01 Iron and steel',
    '14.03.02 Chemical (incl. petrochemical)',
    '14.03.03 Non ferrous metals',
    '14.03.04 Non-metallic mineral products',
    '14.03.05 Transportation equipment',
    '14.03.06 Machinery',
    '14.03.07 Food, beverages and tobacco',
    '14.03.08 Pulp, paper and printing',
    '14.03.09 Wood and wood products',
    '14.03.10 Textiles and leather',
    '14.03.11 Non-specified industry',
    '15 Transport sector',
    '15.01 Domestic air transport',
    '15.02 Road',
    '15.03 Rail',
    '15.04 Domestic navigation',
    '15.05 Pipeline transport',
    '15.06 Non-specified transport',
    '16 Other sector',
    '16.01 Commercial and public services',
    '16.02 Residential',
    '16.03 Agriculture',
    '16.04 Fishing',
    '16.05 Non-specified others',
    '17 Non-energy use',
    '17.01 Transformation sector',
    '17.02 Industry sector',
    '17.03 Transport sector',
    '17.04 Other sector',
    '18 Electricity output in GWh',
    '18.01 MAP electricity plants',
    '18.02 MAP CHP plants',
    '18.03 AP electricity plants',
    '18.04 AP CHP plants',
    '19 Heat output in PJ',
    '19.01 MAP CHP plants',
    '19.02 MAP heat plants',
    '19.03 AP CHP plants',
    '19.04 AP heat plants',
]

NINTH_SECTORS = [
    '01_production',
    '02_imports',
    '03_exports',
    '04_international_marine_bunkers',
    '05_international_aviation_bunkers',
    '06_stock_changes',
    '07_total_primary_energy_supply',
    '08_transfers',
    '09_total_transformation_sector',
    '10_losses_and_own_use',
    '11_statistical_discrepancy',
    '12_total_final_consumption',
    '13_total_final_energy_consumption',
    '14_industry_sector',
    '15_transport_sector',
    '16_other_sector',
    '17_nonenergy_use',
    '18_electricity_output_in_gwh',
    '19_heat_output_in_pj',
    '22_demand_supply_discrepancy',
]

NINTH_SUB1SECTORS = [
    '09_01_electricity_plants',
    '09_02_chp_plants',
    '09_06_gas_processing_plants',
    '09_07_oil_refineries',
    '09_08_coal_transformation',
    '09_10_biofuels_processing',
    '09_12_nonspecified_transformation',
    '09_13_hydrogen_transformation',
    '09_x_heat_plants',
    '10_01_own_use',
    '10_02_transmission_and_distribution_losses',
    '14_01_mining_and_quarrying',
    '14_02_construction',
    '14_03_manufacturing',
    '15_01_domestic_air_transport',
    '15_02_road',
    '15_03_rail',
    '15_04_domestic_navigation',
    '15_05_pipeline_transport',
    '15_06_nonspecified_transport',
    '16_01_buildings',
    '16_02_agriculture_and_fishing',
    '16_05_nonspecified_others',
    '18_01_electricity_plants',
    '18_02_chp_plants',
    '19_01_chp_plants',
    '19_02_heat_plants',
    '08_03_products_transferred',
    '09_09_petrochemical_industry',
    '09_05_chemical_heat_for_electricity_production',
    '09_11_charcoal_processing',
    '08_02_interproduct_transfers',
    '09_04_electric_boilers',
]

NINTH_SUB2SECTORS = [
    '09_01_01_coal_power',
    '09_01_01_coal_power_ccs',
    '09_01_02_gas_power',
    '09_01_02_gas_power_ccs',
    '09_01_02_gas_power_h2',
    '09_01_03_oil',
    '09_01_04_nuclear',
    '09_01_05_hydro',
    '09_01_06_biomass',
    '09_01_07_geothermal',
    '09_01_08_solar',
    '09_01_09_wind',
    '09_01_10_otherrenewable',
    '09_01_11_otherfuel',
    '09_01_12_storage',
    '09_02_01_coal',
    '09_02_02_gas',
    '09_02_03_oil',
    '09_02_04_biomass',
    '09_02_05_others',
    '09_06_01_gas_works_plants',
    '09_06_02_liquefaction_regasification_plants',
    '09_06_03_natural_gas_blending_plants',
    '09_06_04_gastoliquids_plants',
    '09_08_01_coke_ovens',
    '09_08_02_blast_furnaces',
    '09_08_03_patent_fuel_plants',
    '09_08_04_bkb_pb_plants',
    '09_08_05_liquefaction_coal_to_oil',
    '09_13_01_electrolysers',
    '09_13_02_smr_wo_ccs',
    '09_13_03_smr_w_ccs',
    '09_13_04_coal_wo_ccs',
    '09_13_05_coal_w_ccs',
    '09_13_06_others',
    '09_x_01_coal',
    '09_x_02_gas',
    '09_x_03_oil',
    '09_x_04_biomass',
    '09_x_05_others',
    '10_01_01_electricity_chp_and_heat_plants',
    '10_01_02_gas_works_plants',
    '10_01_03_liquefaction_regasification_plants',
    '10_01_04_gastoliquids_plants',
    '10_01_05_coke_ovens',
    '10_01_06_coal_mines',
    '10_01_07_blast_furnaces',
    '10_01_08_patent_fuel_plants',
    '10_01_09_bkb_pb_plants',
    '10_01_10_liquefaction_plants_coal_to_oil',
    '10_01_11_oil_refineries',
    '10_01_12_oil_and_gas_extraction',
    '10_01_13_pump_storage_plants',
    '10_01_14_nuclear_industry',
    '10_01_15_charcoal_production_plants',
    '10_01_16_gasification_plants_for_biogases',
    '10_01_17_nonspecified_own_uses',
    '10_01_18_ccs',
    '14_03_01_iron_and_steel',
    '14_03_02_chemical_incl_petrochemical',
    '14_03_03_non_ferrous_metals',
    '14_03_04_nonmetallic_mineral_products',
    '14_03_05_transportation_equipment',
    '14_03_06_machinery',
    '14_03_07_food_beverages_and_tobacco',
    '14_03_08_pulp_paper_and_printing',
    '14_03_09_wood_and_wood_products',
    '14_03_10_textiles_and_leather',
    '14_03_11_nonspecified_industry',
    '15_01_01_passenger',
    '15_01_02_freight',
    '15_02_01_passenger',
    '15_02_02_freight',
    '15_03_01_passenger',
    '15_03_02_freight',
    '15_04_01_passenger',
    '15_04_02_freight',
    '16_01_01_commercial_and_public_services',
    '16_01_02_residential',
    '16_01_03_ai_training',
    '16_01_04_traditional_data_centres',
    '16_02_03_agriculture',
    '16_02_04_fishing',
    '18_01_01_coal_power',
    '18_01_01_coal_power_ccs',
    '18_01_02_gas_power',
    '18_01_02_gas_power_ccs',
    '18_01_02_gas_power_h2',
    '18_01_03_oil',
    '18_01_04_nuclear',
    '18_01_05_hydro',
    '18_01_06_biomass',
    '18_01_07_geothermal',
    '18_01_08_solar',
    '18_01_09_wind',
    '18_01_10_otherrenewable',
    '18_01_11_otherfuel',
    '18_01_12_storage',
    '18_02_01_coal',
    '18_02_02_gas',
    '18_02_03_oil',
    '18_02_04_biomass',
    '19_01_01_coal',
    '19_01_02_gas',
    '19_01_03_oil',
    '19_01_04_biomass',
    '19_01_05_others',
    '19_02_01_coal',
    '19_02_02_gas',
    '19_02_03_oil',
    '19_02_04_biomass',
    '19_02_05_others',
    '19_02_17_electricity',
]

NINTH_SUB3SECTORS = [
    '09_01_01_01_subcritical',
    '09_01_01_02_superultracritical',
    '09_01_01_03_advultracritical',
    '09_01_01_04_ccs',
    '09_01_02_01_gasturbine',
    '09_01_02_02_combinedcycle',
    '09_01_02_03_ccs',
    '09_01_05_01_large',
    '09_01_05_02_mediumsmall',
    '09_01_05_03_pump',
    '09_01_08_01_utility',
    '09_01_08_02_rooftop',
    '09_01_08_03_csp',
    '09_01_09_01_onshore',
    '09_01_09_02_offshore',
    '14_03_01_01_fs',
    '14_03_01_02_eaf',
    '14_03_01_03_ccs',
    '14_03_01_04_bfbof',
    '14_03_01_05_hydrogen',
    '14_03_02_01_fs',
    '14_03_02_02_ccs',
    '14_03_04_01_ccs',
    '14_03_04_02_nonccs',
    '15_02_01_01_two_wheeler',
    '15_02_01_02_car',
    '15_02_01_03_sports_utility_vehicle',
    '15_02_01_04_light_truck',
    '15_02_01_05_bus',
    '15_02_02_01_two_wheeler_freight',
    '15_02_02_02_light_commercial_vehicle',
    '15_02_02_03_medium_truck',
    '15_02_02_04_heavy_truck',
    '18_01_01_01_subcritical',
    '18_01_01_02_superultracritical',
    '18_01_01_03_advultracritical',
    '18_01_01_04_ccs',
    '18_01_05_01_large',
    '18_01_05_02_mediumsmall',
    '18_01_05_03_pump',
    '18_01_08_01_utility',
    '18_01_08_02_rooftop',
    '18_01_08_03_csp',
    '18_01_09_01_onshore',
    '18_01_09_02_offshore',
]

NINTH_SUB4SECTORS = [
    '15_02_01_01_01_diesel_engine',
    '15_02_01_01_02_gasoline_engine',
    '15_02_01_01_03_battery_ev',
    '15_02_01_01_04_compressed_natual_gas',
    '15_02_01_01_05_plugin_hybrid_ev_gasoline',
    '15_02_01_01_06_plugin_hybrid_ev_diesel',
    '15_02_01_01_07_liquified_petroleum_gas',
    '15_02_01_01_08_fuel_cell_ev',
    '15_02_01_01_09_lng',
    '15_02_01_02_01_diesel_engine',
    '15_02_01_02_02_gasoline_engine',
    '15_02_01_02_03_battery_ev',
    '15_02_01_02_04_compressed_natual_gas',
    '15_02_01_02_05_plugin_hybrid_ev_gasoline',
    '15_02_01_02_06_plugin_hybrid_ev_diesel',
    '15_02_01_02_07_liquified_petroleum_gas',
    '15_02_01_02_08_fuel_cell_ev',
    '15_02_01_02_09_lng',
    '15_02_01_03_01_diesel_engine',
    '15_02_01_03_02_gasoline_engine',
    '15_02_01_03_03_battery_ev',
    '15_02_01_03_04_compressed_natual_gas',
    '15_02_01_03_05_plugin_hybrid_ev_gasoline',
    '15_02_01_03_06_plugin_hybrid_ev_diesel',
    '15_02_01_03_07_liquified_petroleum_gas',
    '15_02_01_03_08_fuel_cell_ev',
    '15_02_01_03_09_lng',
    '15_02_01_04_01_diesel_engine',
    '15_02_01_04_02_gasoline_engine',
    '15_02_01_04_03_battery_ev',
    '15_02_01_04_04_compressed_natual_gas',
    '15_02_01_04_05_plugin_hybrid_ev_gasoline',
    '15_02_01_04_06_plugin_hybrid_ev_diesel',
    '15_02_01_04_07_liquified_petroleum_gas',
    '15_02_01_04_08_fuel_cell_ev',
    '15_02_01_04_09_lng',
    '15_02_01_05_01_diesel_engine',
    '15_02_01_05_02_gasoline_engine',
    '15_02_01_05_03_battery_ev',
    '15_02_01_05_04_compressed_natual_gas',
    '15_02_01_05_05_plugin_hybrid_ev_gasoline',
    '15_02_01_05_06_plugin_hybrid_ev_diesel',
    '15_02_01_05_07_liquified_petroleum_gas',
    '15_02_01_05_08_fuel_cell_ev',
    '15_02_01_05_09_lng',
    '15_02_02_01_01_diesel_engine',
    '15_02_02_01_02_gasoline_engine',
    '15_02_02_01_03_battery_ev',
    '15_02_02_01_04_compressed_natual_gas',
    '15_02_02_01_05_plugin_hybrid_ev_gasoline',
    '15_02_02_01_06_plugin_hybrid_ev_diesel',
    '15_02_02_01_07_liquified_petroleum_gas',
    '15_02_02_01_08_fuel_cell_ev',
    '15_02_02_01_09_lng',
    '15_02_02_02_01_diesel_engine',
    '15_02_02_02_02_gasoline_engine',
    '15_02_02_02_03_battery_ev',
    '15_02_02_02_04_compressed_natual_gas',
    '15_02_02_02_05_plugin_hybrid_ev_gasoline',
    '15_02_02_02_06_plugin_hybrid_ev_diesel',
    '15_02_02_02_07_liquified_petroleum_gas',
    '15_02_02_02_08_fuel_cell_ev',
    '15_02_02_02_09_lng',
    '15_02_02_03_01_diesel_engine',
    '15_02_02_03_02_gasoline_engine',
    '15_02_02_03_03_battery_ev',
    '15_02_02_03_04_compressed_natual_gas',
    '15_02_02_03_05_plugin_hybrid_ev_gasoline',
    '15_02_02_03_06_plugin_hybrid_ev_diesel',
    '15_02_02_03_07_liquified_petroleum_gas',
    '15_02_02_03_08_fuel_cell_ev',
    '15_02_02_03_09_lng',
    '15_02_02_04_01_diesel_engine',
    '15_02_02_04_02_gasoline_engine',
    '15_02_02_04_03_battery_ev',
    '15_02_02_04_04_compressed_natual_gas',
    '15_02_02_04_05_plugin_hybrid_ev_gasoline',
    '15_02_02_04_06_plugin_hybrid_ev_diesel',
    '15_02_02_04_07_liquified_petroleum_gas',
    '15_02_02_04_08_fuel_cell_ev',
    '15_02_02_04_09_lng',
]

#NAME MAPPINGS:
# 9th_label	9th_column	esto_label	esto_column	name
# 		06.04 Additives/  oxygenates	products	Additives/  oxygenates
# 09_01_01_03_advultracritical	sub3sectors			Advanced ultra-supercritical coal power
# 18_01_01_03_advultracritical	sub3sectors			Advanced ultra-supercritical coal power (electricity output)
# 16_02_03_agriculture	sub2sectors	16.03 Agriculture	flows	Agriculture
# 16_02_agriculture_and_fishing	sub1sectors			Agriculture and fishing
# 16_01_03_ai_training	sub2sectors			AI training
# 		01.04 Anthracite	products	Anthracite
# 		18.04 AP CHP plants	flows	AP CHP plants  (electricity)
# 		19.03 AP CHP plants	flows	AP CHP plants (heat)
# 		18.03 AP electricity plants	flows	AP electricity plants
# 		19.04 AP heat plants	flows	AP heat plants
# 		09.02 Autoproducers	flows	Autoproducers
# 07_02_aviation_gasoline	subfuels	07.02 Aviation gasoline	products	Aviation gasoline
# 15_02_bagasse	subfuels	15.02 Bagasse	products	Bagasse
# 15_02_01_01_03_battery_ev	sub4sectors			Battery EV
# 16_07_bio_jet_kerosene	subfuels	16.07 Bio jet kerosene	products	Bio jet kerosene
# 16_06_biodiesel	subfuels	16.06 Biodiesel	products	Biodiesel
# 09_10_biofuels_processing	sub1sectors	09.10 Biofuels processing	flows	Biofuels processing
# 16_01_biogas	subfuels	16.01 Biogas	products	Biogas
# 16_05_biogasoline	subfuels	16.05 Biogasoline	products	Biogasoline
# 09_02_04_biomass	sub2sectors			Biomass CHP
# 18_02_04_biomass	sub2sectors			Biomass CHP (electricity output)
# 19_01_04_biomass	sub2sectors			Biomass CHP heat
# 19_02_04_biomass	sub2sectors			Biomass heat plants
# 09_01_06_biomass	sub2sectors			Biomass power
# 18_01_06_biomass	sub2sectors			Biomass power (electricity output)
# 		07.14 Bitumen	products	Bitumen
# 		02.08 BKB/PB	products	BKB/PB
# 09_08_04_bkb_pb_plants	sub2sectors	09.08.04 BKB/PB plants	flows	BKB/PB plants
# 10_01_09_bkb_pb_plants	sub2sectors	10.01.09 BKB/PB plants	flows	BKB/PB plants (own-use)
# 15_04_black_liquor	subfuels	15.04 Black liqour	products	Black liqour
# 		02.04 Blast furnace gas	products	Blast furnace gas
# 14_03_01_04_bfbof	sub3sectors			Blast furnace/basic oxygen furnace
# 09_08_02_blast_furnaces	sub2sectors	09.08.02 Blast furnaces	flows	Blast furnaces
# 10_01_07_blast_furnaces	sub2sectors	10.01.07 Blast furnaces	flows	Blast furnaces (own-use)
# 16_01_buildings	sub1sectors			Buildings
# 15_02_01_05_bus	sub3sectors			Bus
# 15_02_01_05_03_battery_ev	sub4sectors			Bus - Battery EV
# 15_02_01_05_04_compressed_natual_gas	sub4sectors			Bus - Compressed natural gas
# 15_02_01_05_01_diesel_engine	sub4sectors			Bus - Diesel engine
# 15_02_01_05_08_fuel_cell_ev	sub4sectors			Bus - Fuel cell EV
# 15_02_01_05_02_gasoline_engine	sub4sectors			Bus - Gasoline engine
# 15_02_01_05_09_lng	sub4sectors			Bus - LNG
# 15_02_01_05_07_liquified_petroleum_gas	sub4sectors			Bus - LPG
# 15_02_01_05_06_plugin_hybrid_ev_diesel	sub4sectors			Bus - Plug-in hybrid EV (diesel)
# 15_02_01_05_05_plugin_hybrid_ev_gasoline	sub4sectors			Bus - Plug-in hybrid EV (gasoline)
# 15_02_01_02_car	sub3sectors			Car
# 15_02_01_02_03_battery_ev	sub4sectors			Car - Battery EV
# 15_02_01_02_04_compressed_natual_gas	sub4sectors			Car - Compressed natural gas
# 15_02_01_02_01_diesel_engine	sub4sectors			Car - Diesel engine
# 15_02_01_02_08_fuel_cell_ev	sub4sectors			Car - Fuel cell EV
# 15_02_01_02_02_gasoline_engine	sub4sectors			Car - Gasoline engine
# 15_02_01_02_09_lng	sub4sectors			Car - LNG
# 15_02_01_02_07_liquified_petroleum_gas	sub4sectors			Car - LPG
# 15_02_01_02_06_plugin_hybrid_ev_diesel	sub4sectors			Car - Plug-in hybrid EV (diesel)
# 15_02_01_02_05_plugin_hybrid_ev_gasoline	sub4sectors			Car - Plug-in hybrid EV (gasoline)
# 15_03_charcoal	subfuels	15.03 Charcoal	products	Charcoal
# 09_11_charcoal_processing	sub1sectors	09.11 Charcoal processing	flows	Charcoal processing
# 10_01_15_charcoal_production_plants	sub2sectors	10.01.15 Charcoal production plants	flows	Charcoal production plants
# 14_03_02_01_fs	sub3sectors			Chemical - FS
# 14_03_02_chemical_incl_petrochemical	sub2sectors	14.03.02 Chemical (incl. petrochemical)	flows	Chemical (incl. petrochemical)
# 14_03_02_02_ccs	sub3sectors			Chemical (incl. petrochemical) with CCS
# 09_05_chemical_heat_for_electricity_production	sub1sectors	09.05 Chemical heat for electricity production	flows	Chemical heat for electricity production
# 09_02_chp_plants	sub1sectors	09.01.02 CHP plants	flows	CHP plants
# 		09.02.02 CHP plants	flows	CHP plants (autoproducers)
# 18_02_chp_plants	sub1sectors			CHP plants (electricity output)
# 19_01_chp_plants	sub1sectors			CHP plants (heat output)
# 01_coal	fuels	01 Coal	products	Coal
# 09_02_01_coal	sub2sectors			Coal CHP
# 18_02_01_coal	sub2sectors			Coal CHP (electricity output)
# 19_01_01_coal	sub2sectors			Coal CHP heat
# 19_02_01_coal	sub2sectors			Coal heat plants
# 10_01_06_coal_mines	sub2sectors	10.01.06 Coal mines	flows	Coal mines
# 		01.99 Coal nonspecified	products	Coal nonspecified
# 09_01_01_coal_power	sub2sectors			Coal power
# 18_01_01_coal_power	sub2sectors			Coal power (electricity output)
# 18_01_01_04_ccs	sub3sectors			Coal power CCS (electricity output) 1
# 18_01_01_coal_power_ccs	sub2sectors			Coal power CCS (electricity output) 2
# 09_01_01_04_ccs	sub3sectors			Coal power CCS 1
# 09_01_01_coal_power_ccs	sub2sectors			Coal power CCS 2
# 02_coal_products	fuels	02 Coal products	products	Coal products
# 		02.07 Coal tar	products	Coal tar
# 09_08_coal_transformation	sub1sectors	09.08 Coal transformation	flows	Coal transformation
# 09_13_05_coal_w_ccs	sub2sectors			Coal with CCS
# 09_13_04_coal_wo_ccs	sub2sectors			Coal without CCS
# 		02.01 Coke oven coke	products	Coke oven coke
# 		02.03 Coke oven gas	products	Coke oven gas
# 09_08_01_coke_ovens	sub2sectors	09.08.01 Coke ovens	flows	Coke ovens
# 10_01_05_coke_ovens	sub2sectors	10.01.05 Coke ovens	flows	Coke ovens (own-use)
# 01_01_coking_coal	subfuels	01.01 Coking coal	products	Coking coal
# 09_01_02_02_combinedcycle	sub3sectors			Combined cycle power
# 18_01_02_02_combinedcycle	sub3sectors			Combined cycle power (electricity output)
# 16_01_01_commercial_and_public_services	sub2sectors	16.01 Commercial and public services	flows	Commercial and public services
# 15_02_01_01_04_compressed_natual_gas	sub4sectors			Compressed natural gas
# 14_02_construction	sub1sectors	14.02 Construction	flows	Construction
# 06_01_crude_oil	subfuels	06.01 Crude oil	products	Crude oil
# 06_crude_oil_and_ngl	fuels	06 Crude oil & NGL	products	Crude oil & NGL
# 09_01_08_03_csp	sub3sectors			CSP solar
# 18_01_08_03_csp	sub3sectors			CSP solar (electricity output)
# 22_demand_supply_discrepancy	sectors			Demand-supply discrepancy
# 15_02_01_01_01_diesel_engine	sub4sectors			Diesel engine
# 15_01_domestic_air_transport	sub1sectors	15.01 Domestic air transport	flows	Domestic air transport
# 15_04_domestic_navigation	sub1sectors	15.04 Domestic navigation	flows	Domestic navigation
# 15_04_02_freight	sub2sectors			Domestic navigation freight
# 15_04_01_passenger	sub2sectors			Domestic navigation passenger
# 14_03_01_02_eaf	sub3sectors			Electric arc furnace
# 09_04_electric_boilers	sub1sectors	09.04 Electric boilers	flows	Electric boilers
# 19_02_17_electricity	sub2sectors			Electric heat plants (heat output)
# 17_electricity	fuels	17 Electricity	products	Electricity
# 18_electricity_output_in_gwh	sectors	18 Electricity output in GWh	flows	Electricity output in GWh
# 09_01_electricity_plants	sub1sectors	09.01.01 Electricity plants	flows	Electricity plants
# 		09.02.01 Electricity plants	flows	Electricity plants (autoproducers)
# 18_01_electricity_plants	sub1sectors			Electricity plants (electricity output)
# 10_01_01_electricity_chp_and_heat_plants	sub2sectors	10.01.01 Electricity, CHP and heat plants	flows	Electricity, CHP and heat plants
# 09_13_01_electrolysers	sub2sectors			Electrolysers
# 07_11_ethane	subfuels	07.11 Ethane	products	Ethane
# 03_exports	sectors	03 Exports	flows	Exports
# 16_02_04_fishing	sub2sectors	16.04 Fishing	flows	Fishing
# 14_03_07_food_beverages_and_tobacco	sub2sectors	14.03.07 Food, beverages and tobacco	flows	Food, beverages and tobacco
# 15_01_02_freight	sub2sectors			Freight
# 07_08_fuel_oil	subfuels	07.08 Fuel oil	products	Fuel oil
# 15_01_fuelwood_and_woodwaste	subfuels	15.01 Fuelwood & woodwaste	products	Fuelwood & woodwaste
# 08_gas	fuels	08 Gas	products	Gas
# 09_02_02_gas	sub2sectors			Gas CHP
# 18_02_02_gas	sub2sectors			Gas CHP (electricity output)
# 19_01_02_gas	sub2sectors			Gas CHP heat
# 		02.02 Gas coke	products	Gas coke
# 19_02_02_gas	sub2sectors			Gas heat plants
# 		08.99 Gas nonspecified	products	Gas nonspecified
# 09_01_02_gas_power	sub2sectors			Gas power
# 18_01_02_gas_power	sub2sectors			Gas power (electricity output)
# 18_01_02_03_ccs	sub3sectors			Gas power CCS (electricity output) 1
# 18_01_02_gas_power_ccs	sub2sectors			Gas power CCS (electricity output) 2
# 09_01_02_gas_power_ccs	sub2sectors			Gas power CCS 1
# 09_01_02_03_ccs	sub3sectors			Gas power CCS 2
# 09_01_02_gas_power_h2	sub2sectors			Gas power H2
# 18_01_02_gas_power_h2	sub2sectors			Gas power H2 (electricity output)
# 09_06_gas_processing_plants	sub1sectors	09.06 Gas processing plants	flows	Gas processing plants
# 		08.04 Gas separation	flows	Gas separation
# 09_01_02_01_gasturbine	sub3sectors			Gas turbine power
# 18_01_02_01_gasturbine	sub3sectors			Gas turbine power (electricity output)
# 08_03_gas_works_gas	subfuels	08.03 Gas works gas	products	Gas works gas
# 09_06_01_gas_works_plants	sub2sectors	09.06.01 Gas works plants	flows	Gas works plants
# 10_01_02_gas_works_plants	sub2sectors	10.01.02 Gas works plants	flows	Gas works plants (own-use)
# 07_07_gas_diesel_oil	subfuels	07.07 Gas/diesel oil	products	Gas/diesel oil
# 10_01_16_gasification_plants_for_biogases	sub2sectors	10.01.16 Gasification plants for biogases	flows	Gasification plants for biogases
# 15_02_01_01_02_gasoline_engine	sub4sectors			Gasoline engine
# 		07.04 Gasoline type jet fuel	products	Gasoline type jet fuel
# 09_06_04_gastoliquids_plants	sub2sectors	09.06.04 Gas-to-liquids plants	flows	Gas-to-liquids plants
# 10_01_04_gastoliquids_plants	sub2sectors	10.01.04 Gas-to-liquids plants	flows	Gas-to-liquids plants (own-use)
# 11_geothermal	fuels	11 Geothermal	products	Geothermal
# 09_01_07_geothermal	sub2sectors			Geothermal power
# 18_01_07_geothermal	sub2sectors			Geothermal power (electricity output)
# 17_x_green_electricity	fuels			Green electricity
# 18_heat	fuels	18 Heat	products	Heat
# 19_heat_output_in_pj	sectors	19 Heat output in PJ	flows	Heat output in PJ
# 19_02_heat_plants	sub1sectors	09.01.03 Heat plants	flows	Heat plants
# 		09.02.03 Heat plants	flows	Heat plants (autoproducers)
# 09_03_heat_pumps	sub1sectors	09.03 Heat pumps	flows	Heat pumps
# 15_02_02_04_03_battery_ev	sub4sectors			Heavy truck - Battery EV
# 15_02_02_04_04_compressed_natual_gas	sub4sectors			Heavy truck - Compressed natural gas
# 15_02_02_04_01_diesel_engine	sub4sectors			Heavy truck - Diesel engine
# 15_02_02_04_08_fuel_cell_ev	sub4sectors			Heavy truck - Fuel cell EV
# 15_02_02_04_02_gasoline_engine	sub4sectors			Heavy truck - Gasoline engine
# 15_02_02_04_09_lng	sub4sectors			Heavy truck - LNG (freight)
# 15_02_02_04_07_liquified_petroleum_gas	sub4sectors			Heavy truck - LPG
# 15_02_02_04_06_plugin_hybrid_ev_diesel	sub4sectors			Heavy truck - Plug-in hybrid EV (diesel)
# 15_02_02_04_05_plugin_hybrid_ev_gasoline	sub4sectors			Heavy truck - Plug-in hybrid EV (gasoline)
# 15_02_02_04_heavy_truck	sub3sectors			Heavy truck 1
# 15_02_01_04_heavy_truck	sub3sectors			Heavy truck 2
# 10_hydro	fuels	10 Hydro	products	Hydro
# 09_01_05_hydro	sub2sectors			Hydro power
# 18_01_05_hydro	sub2sectors			Hydro power (electricity output)
# 14_03_01_05_hydrogen	sub3sectors			Hydrogen
# 09_13_hydrogen_transformation	sub1sectors			Hydrogen transformation
# 09_13_06_others	sub2sectors			Hydrogen transformation (Others)
# 02_imports	sectors	02 Imports	flows	Imports
# 16_02_industrial_waste	subfuels	16.02 Industrial waste	products	Industrial waste
# 14_industry_sector	sectors	14 Industry sector	flows	Industry sector
# 17_02_industry_sector	sub1sectors	17.02 Industry sector	flows	Industry sector (non-energy)
# 05_international_aviation_bunkers	sectors	05 International aviation bunkers	flows	International aviation bunkers
# 04_international_marine_bunkers	sectors	04 International marine bunkers	flows	International marine bunkers
# 08_02_interproduct_transfers	sub1sectors	08.02 Interproduct transfers	flows	Interproduct transfers
# 14_03_01_iron_and_steel	sub2sectors	14.03.01 Iron and steel	flows	Iron and steel
# 14_03_01_01_fs	sub3sectors			Iron and steel - FS
# 14_03_01_03_ccs	sub3sectors			Iron and steel CCS
# 07_x_jet_fuel	subfuels			Jet fuel
# 07_06_kerosene	subfuels	07.06 Kerosene	products	Kerosene
# 		07.05 Kerosene type jet fuel	products	Kerosene type jet fuel
# 09_01_05_01_large	sub3sectors			Large hydro
# 18_01_05_01_large	sub3sectors			Large hydro (electricity output)
# 15_02_01_03_light_truck	sub3sectors			Light truck
# 15_02_02_03_light_truck	sub3sectors			Light truck (freight)
# 15_02_02_03_03_battery_ev	sub4sectors			Light truck (freight) - Battery EV
# 15_02_02_03_04_compressed_natual_gas	sub4sectors			Light truck (freight) - Compressed natural gas
# 15_02_02_03_01_diesel_engine	sub4sectors			Light truck (freight) - Diesel engine
# 15_02_02_03_02_gasoline_engine	sub4sectors			Light truck (freight) - Gasoline engine
# 15_02_02_03_06_plugin_hybrid_ev_diesel	sub4sectors			Light truck (freight) - Plug-in hybrid EV (diesel)
# 15_02_02_03_05_plugin_hybrid_ev_gasoline	sub4sectors			Light truck (freight) - Plug-in hybrid EV (gasoline)
# 15_02_01_04_light_truck	sub3sectors			Light truck (passenger)
# 15_02_01_04_03_battery_ev	sub4sectors			Light truck (passenger) - Battery EV
# 15_02_01_04_04_compressed_natual_gas	sub4sectors			Light truck (passenger) - Compressed natural gas
# 15_02_01_04_01_diesel_engine	sub4sectors			Light truck (passenger) - Diesel engine
# 15_02_01_04_08_fuel_cell_ev	sub4sectors			Light truck (passenger) - Fuel cell EV
# 15_02_01_04_02_gasoline_engine	sub4sectors			Light truck (passenger) - Gasoline engine
# 15_02_01_04_09_lng	sub4sectors			Light truck (passenger) - LNG
# 15_02_01_04_07_liquified_petroleum_gas	sub4sectors			Light truck (passenger) - LPG
# 15_02_01_04_06_plugin_hybrid_ev_diesel	sub4sectors			Light truck (passenger) - Plug-in hybrid EV (diesel)
# 15_02_01_04_05_plugin_hybrid_ev_gasoline	sub4sectors			Light truck (passenger) - Plug-in hybrid EV (gasoline)
# 15_02_01_02_light_vehicle	sub3sectors			Light vehicle
# 15_02_02_02_03_battery_ev	sub4sectors			Light vehicle (freight) - Battery EV
# 15_02_02_02_04_compressed_natual_gas	sub4sectors			Light vehicle (freight) - Compressed natural gas
# 15_02_02_02_01_diesel_engine	sub4sectors			Light vehicle (freight) - Diesel engine
# 15_02_02_02_08_fuel_cell_ev	sub4sectors			Light vehicle (freight) - Fuel cell EV
# 15_02_02_02_02_gasoline_engine	sub4sectors			Light vehicle (freight) - Gasoline engine
# 15_02_02_02_09_lng	sub4sectors			Light vehicle (freight) - LNG
# 15_02_02_02_07_liquified_petroleum_gas	sub4sectors			Light vehicle (freight) - LPG
# 15_02_02_02_06_plugin_hybrid_ev_diesel	sub4sectors			Light vehicle (freight) - Plug-in hybrid EV (diesel)
# 15_02_02_02_05_plugin_hybrid_ev_gasoline	sub4sectors			Light vehicle (freight) - Plug-in hybrid EV (gasoline)
# 15_02_02_02_light_vehicle	sub3sectors			Light vehicle (freight) 1
# 15_02_02_02_light_commercial_vehicle	sub3sectors			Light vehicle (freight) 2
# 01_05_lignite	subfuels	01.05 Lignite	products	Lignite
# 09_08_05_liquefaction_coal_to_oil	sub2sectors	09.08.05 Liquefaction (coal to oil)	flows	Liquefaction (coal to oil)
# 		10.01.10 Liquefaction plants (Coal to Oil)	flows	Liquefaction plants (Coal to Oil)
# 10_01_10_liquefaction_plants_coal_to_oil	sub2sectors			Liquefaction plants coal to oil (own use)
# 09_06_02_liquefaction_regasification_plants	sub2sectors	09.06.02 Liquefaction/regasification plants	flows	Liquefaction/regasification plants
# 10_01_03_liquefaction_regasification_plants	sub2sectors	10.01.03 Liquefaction/regasification plants	flows	Liquefaction/regasification plants (own-use)
# 08_02_lng	subfuels	08.02 LNG	products	LNG
# 10_losses_and_own_use	sectors	10 Losses & own use	flows	Losses & own use
# 07_09_lpg	subfuels	07.09 LPG	products	LPG
# 		07.13 Lubricants	products	Lubricants
# 14_03_06_machinery	sub2sectors	14.03.06 Machinery	flows	Machinery
# 		09.01 Main activity producer	flows	Main activity producer
# 14_03_manufacturing	sub1sectors	14.03 Manufacturing	flows	Manufacturing
# 		18.02 MAP CHP plants	flows	MAP CHP plants (electricity)
# 		19.01 MAP CHP plants	flows	MAP CHP plants (heat)
# 		18.01 MAP electricity plants	flows	MAP electricity plants
# 		19.02 MAP heat plants	flows	MAP heat plants
# 15_02_02_03_medium_truck	sub3sectors			Medium truck
# 15_02_02_03_08_fuel_cell_ev	sub4sectors			Medium truck - Fuel cell EV
# 15_02_02_03_09_lng	sub4sectors			Medium truck - LNG
# 15_02_02_03_07_liquified_petroleum_gas	sub4sectors			Medium truck - LPG
# 09_01_05_02_mediumsmall	sub3sectors			Medium/small hydro
# 18_01_05_02_mediumsmall	sub3sectors			Medium/small hydro (electricity output)
# 14_01_mining_and_quarrying	sub1sectors	14.01 Mining and quarrying	flows	Mining and quarrying
# 21_modern_renewables	fuels	21 Modern renewables	products	Modern renewables
# 07_01_motor_gasoline	subfuels	07.01 Motor gasoline	products	Motor gasoline
# 16_04_municipal_solid_waste_nonrenewable	subfuels	16.04 Municipal solid waste (non-renewable)	products	Municipal solid waste (non-renewable)
# 16_03_municipal_solid_waste_renewable	subfuels	16.03 Municipal solid waste (renewable)	products	Municipal solid waste (renewable)
# 07_03_naphtha	subfuels	07.03 Naphtha	products	Naphtha
# 08_01_natural_gas	subfuels	08.01 Natural gas	products	Natural gas
# 09_06_03_natural_gas_blending_plants	sub2sectors	09.06.03 Natural gas blending plants	flows	Natural gas blending plants
# 06_02_natural_gas_liquids	subfuels	06.02 Natural gas liquids	products	Natural gas liquids
# 14_03_03_non_ferrous_metals	sub2sectors	14.03.03 Non ferrous metals	flows	Non ferrous metals
# 14_03_04_02_nonccs	sub3sectors			Non-CCS
# 17_nonenergy_use	sectors	17 Non-energy use	flows	Non-energy use
# 14_03_04_nonmetallic_mineral_products	sub2sectors	14.03.04 Non-metallic mineral products	flows	Non-metallic mineral products
# 14_03_04_01_ccs	sub3sectors			Non-metallic mineral products CCS
# 		01.99 Non-specified Coal	products	Non-specified Coal
# 		08.99 Non-specified Gas	products	Non-specified Gas
# 14_03_11_nonspecified_industry	sub2sectors	14.03.11 Non-specified industry	flows	Non-specified industry
# 16_05_nonspecified_others	sub1sectors	16.05 Non-specified others	flows	Non-specified others
# 10_01_17_nonspecified_own_uses	sub2sectors	10.01.17 Non-specified own uses	flows	Non-specified own uses
# 		07.99 Non-specified Petroleum Products	products	Non-specified Petroleum Products
# 		12.02 Non-specified Solar	products	Non-specified Solar
# 		08.99 Transfers nonspecified	flows	Transfers nonspecified
# 09_12_nonspecified_transformation	sub1sectors	09.12 Non-specified transformation	flows	Non-specified transformation
# 15_06_nonspecified_transport	sub1sectors	15.06 Non-specified transport	flows	Non-specified transport
# 09_nuclear	fuels	09 Nuclear	products	Nuclear
# 		10.01.14 Nuclear industry	flows	Nuclear industry
# 10_01_14_nuclear_industry	sub2sectors			Nuclear industry (own use)
# 09_01_04_nuclear	sub2sectors			Nuclear power
# 18_01_04_nuclear	sub2sectors			Nuclear power (electricity output)
# 12_01_of_which_photovoltaics	subfuels	12.01 of which: Photovoltaics	products	of which: Photovoltaics
# 09_01_09_02_offshore	sub3sectors			Offshore wind
# 18_01_09_02_offshore	sub3sectors			Offshore wind (electricity output)
# 10_01_12_oil_and_gas_extraction	sub2sectors	10.01.12 Oil and gas extraction	flows	Oil and gas extraction
# 09_02_03_oil	sub2sectors			Oil CHP
# 18_02_03_oil	sub2sectors			Oil CHP (electricity output)
# 19_01_03_oil	sub2sectors			Oil CHP heat
# 19_02_03_oil	sub2sectors			Oil heat plants
# 09_01_03_oil	sub2sectors			Oil power
# 18_01_03_oil	sub2sectors			Oil power (electricity output)
# 09_07_oil_refineries	sub1sectors	09.07 Oil refineries	flows	Oil refineries
# 10_01_11_oil_refineries	sub2sectors	10.01.11 Oil refineries	flows	Oil refineries (own-use)
# 05_oil_shale_and_oil_sands	fuels	05 Oil shale and oil sands	products	Oil shale and oil sands
# 09_01_09_01_onshore	sub3sectors			Onshore wind
# 18_01_09_01_onshore	sub3sectors			Onshore wind (electricity output)
# 16_x_ammonia	subfuels			Other ammonia
# 09_x_04_biomass	sub2sectors			Other biomass (transformation)
# 		01.02 Other bituminous coal	products	Other bituminous coal
# 09_02_05_others	sub2sectors			Other CHP
# 19_01_05_others	sub2sectors			Other CHP (heat output)
# 09_x_01_coal	sub2sectors			Other coal
# 16_x_efuel	subfuels			Other E-fuel
# 09_01_11_otherfuel	sub2sectors			Other fuel power
# 18_01_11_otherfuel	sub2sectors			Other fuel power (electricity output)
# 09_x_02_gas	sub2sectors			Other gas processing
# 19_02_05_others	sub2sectors			Other heat plants (heat output)
# 09_x_heat_plants	sub1sectors			Other heat plants 1
# 09_x_05_others	sub2sectors			Other heat plants 2
# 06_x_other_hydrocarbons	subfuels	06.05 Other hydrocarbons	products	Other hydrocarbons
# 16_x_hydrogen	subfuels			Other hydrogen
# 16_08_other_liquid_biofuels	subfuels	16.08 Other liquid biofuels	products	Other liquid biofuels
# 09_x_03_oil	sub2sectors			Other oil
# 07_x_other_petroleum_products	subfuels			Other petroleum products
# 		07.17 Other products	products	Other products
# 		02.05 Other recovered gases	products	Other recovered gases
# 09_01_10_otherrenewable	sub2sectors			Other renewable power
# 18_01_10_otherrenewable	sub2sectors			Other renewable power (electricity output)
# 16_other_sector	sectors	16 Other sector	flows	Other sector
# 17_04_other_sector	sub1sectors	17.04 Other sector	flows	Other sector (non-energy)
# 12_x_other_solar	subfuels			Other solar
# 15_05_other_biomass	subfuels	15.05 Other biomass	products	Other biomass
# 16_09_other_sources	subfuels	16.09 Other sources	products	Other sources
# 01_x_thermal_coal	subfuels			Other thermal coal
# 16_others	fuels	16 Others	products	Others
# 10_01_own_use	sub1sectors			Own use
# 		10.01 Own Use	flows	Own Use
# 10_01_18_ccs	sub2sectors			Own use CCS
# 		07.15 Paraffin  waxes	products	Paraffin  waxes
# 15_01_01_passenger	sub2sectors			Passenger
# 		02.06 Patent fuel	products	Patent fuel
# 09_08_03_patent_fuel_plants	sub2sectors	09.08.03 Patent fuel plants	flows	Patent fuel plants
# 10_01_08_patent_fuel_plants	sub2sectors	10.01.08 Patent fuel plants	flows	Patent fuel plants (own-use)
# 03_peat	fuels	03 Peat	products	Peat
# 04_peat_products	fuels	04 Peat products	products	Peat products
# 		07.99 PetProd nonspecified	products	PetProd nonspecified
# 09_09_petrochemical_industry	sub1sectors	09.09 Petrochemical industry	flows	Petrochemical industry
# 		07.16 Petroleum coke	products	Petroleum coke
# 07_petroleum_products	fuels	07 Petroleum products	products	Petroleum products
# 15_05_pipeline_transport	sub1sectors	15.05 Pipeline transport	flows	Pipeline transport
# 15_02_01_01_06_plugin_hybrid_ev_diesel	sub4sectors			Plug-in hybrid EV (diesel)
# 15_02_01_01_05_plugin_hybrid_ev_gasoline	sub4sectors			Plug-in hybrid EV (gasoline)
# 01_production	sectors	01 Production	flows	Production
# 08_03_products_transferred	sub1sectors	08.03 Products transferred	flows	Products transferred
# 14_03_08_pulp_paper_and_printing	sub2sectors	14.03.08 Pulp, paper and printing	flows	Pulp, paper and printing
# 10_01_13_pump_storage_plants	sub2sectors	10.01.13 Pump storage plants	flows	Pump storage plants
# 09_01_05_03_pump	sub3sectors			Pumped storage hydro
# 18_01_05_03_pump	sub3sectors			Pumped storage hydro (electricity output)
# 15_03_rail	sub1sectors	15.03 Rail	flows	Rail
# 15_03_02_freight	sub2sectors			Rail freight
# 15_03_01_passenger	sub2sectors			Rail passenger
# 		08.01 Recycled products	flows	Recycled products
# 		06.03 Refinery feedstocks	products	Refinery feedstocks
# 07_10_refinery_gas_not_liquefied	subfuels	07.10 Refinery gas (not liquefied)	products	Refinery gas (not liquefied)
# 16_01_02_residential	sub2sectors	16.02 Residential	flows	Residential
# 15_02_road	sub1sectors	15.02 Road	flows	Road
# 15_02_02_freight	sub2sectors			Road freight
# 15_02_01_passenger	sub2sectors			Road passenger
# 09_01_08_02_rooftop	sub3sectors			Rooftop solar
# 18_01_08_02_rooftop	sub3sectors			Rooftop solar (electricity output)
# 09_13_03_smr_w_ccs	sub2sectors			SMR with CCS
# 09_13_02_smr_wo_ccs	sub2sectors			SMR without CCS
# 12_solar	fuels	12 Solar	products	Solar
# 		12.99 Solar nonspecified	products	Solar nonspecified
# 09_01_08_solar	sub2sectors			Solar power
# 18_01_08_solar	sub2sectors			Solar power (electricity output)
# 15_solid_biomass	fuels	15 Solid biomass	products	Solid biomass
# 15_02_01_03_sports_utility_vehicle	sub3sectors			Sports utility vehicle
# 15_02_01_03_03_battery_ev	sub4sectors			Sports utility vehicle - Battery EV
# 15_02_01_03_04_compressed_natual_gas	sub4sectors			Sports utility vehicle - Compressed natural gas
# 15_02_01_03_01_diesel_engine	sub4sectors			Sports utility vehicle - Diesel engine
# 15_02_01_03_08_fuel_cell_ev	sub4sectors			Sports utility vehicle - Fuel cell EV
# 15_02_01_03_02_gasoline_engine	sub4sectors			Sports utility vehicle - Gasoline engine
# 15_02_01_03_09_lng	sub4sectors			Sports utility vehicle - LNG
# 15_02_01_03_07_liquified_petroleum_gas	sub4sectors			Sports utility vehicle - LPG
# 15_02_01_03_06_plugin_hybrid_ev_diesel	sub4sectors			Sports utility vehicle - Plug-in hybrid EV (diesel)
# 15_02_01_03_05_plugin_hybrid_ev_gasoline	sub4sectors			Sports utility vehicle - Plug-in hybrid EV (gasoline)
# 11_statistical_discrepancy	sectors	11 Statistical discrepancy	flows	Statistical discrepancy
# 06_stock_changes	sectors	06 Stock changes	flows	Stock changes
# 09_01_12_storage	sub2sectors			Storage power
# 18_01_12_storage	sub2sectors			Storage power (electricity output)
# 		01.03 Sub-bituminous coal	products	Sub-bituminous coal
# 09_01_01_01_subcritical	sub3sectors			Subcritical coal power
# 18_01_01_01_subcritical	sub3sectors			Subcritical coal power (electricity output)
# 09_01_01_02_superultracritical	sub3sectors			Super ultra-supercritical coal power
# 18_01_01_02_superultracritical	sub3sectors			Super ultra-supercritical coal power (electricity output)
# 14_03_10_textiles_and_leather	sub2sectors	14.03.10 Textiles and leather	flows	Textiles and leather
# 13_tide_wave_ocean	fuels			Tide wave ocean
# 		13 Tide, wave, ocean	products	Tide, wave, ocean
# 19_total	fuels	19 Total	products	Total
# 22_total_combustion_emissions	fuels			Total combustion emissions (fuels column)
# 20_total_combustion_emissions	sectors			Total combustion emissions (sectors column)
# 12_total_final_consumption	sectors	12 Total final consumption	flows	Total final consumption
# 13_total_final_energy_consumption	sectors	13 Total final energy consumption	flows	Total final energy consumption
# 07_total_primary_energy_supply	sectors	07 Total primary energy supply	flows	Total primary energy supply
# 20_total_renewables	fuels			Total renewables
# 		20 Total Renewables	products	Total Renewables
# 09_total_transformation_sector	sectors	09 Total transformation sector	flows	Total transformation sector
# 16_01_04_traditional_data_centres	sub2sectors			Traditional data centres
# 08_transfers	sectors	08 Transfers	flows	Transfers
# 17_01_transformation_sector	sub1sectors	17.01 Transformation sector	flows	Transformation sector
# 		10.02 Transmision and distribution losses	flows	Transmision and distribution losses
# 10_02_transmission_and_distribution_losses	sub1sectors	10.02 Transmission and distribution losses	flows	Transmission and distribution losses
# 15_transport_sector	sectors	15 Transport sector	flows	Transport sector
# 17_03_transport_sector	sub1sectors	17.03 Transport sector	flows	Transport sector (non-energy)
# 14_03_05_transportation_equipment	sub2sectors	14.03.05 Transportation equipment	flows	Transportation equipment
# 15_02_01_01_two_wheeler	sub3sectors			Two-wheeler
# 15_02_01_01_08_fuel_cell_ev	sub4sectors			Two-wheeler - Fuel cell EV
# 15_02_01_01_09_lng	sub4sectors			Two-wheeler - LNG
# 15_02_01_01_07_liquified_petroleum_gas	sub4sectors			Two-wheeler - LPG
# 15_02_02_01_03_battery_ev	sub4sectors			Two-wheeler (freight) - Battery EV
# 15_02_02_01_04_compressed_natual_gas	sub4sectors			Two-wheeler (freight) - Compressed natural gas
# 15_02_02_01_01_diesel_engine	sub4sectors			Two-wheeler (freight) - Diesel engine
# 15_02_02_01_08_fuel_cell_ev	sub4sectors			Two-wheeler (freight) - Fuel cell EV
# 15_02_02_01_02_gasoline_engine	sub4sectors			Two-wheeler (freight) - Gasoline engine
# 15_02_02_01_09_lng	sub4sectors			Two-wheeler (freight) - LNG
# 15_02_02_01_07_liquified_petroleum_gas	sub4sectors			Two-wheeler (freight) - LPG
# 15_02_02_01_06_plugin_hybrid_ev_diesel	sub4sectors			Two-wheeler (freight) - Plug-in hybrid EV (diesel)
# 15_02_02_01_05_plugin_hybrid_ev_gasoline	sub4sectors			Two-wheeler (freight) - Plug-in hybrid EV (gasoline)
# 15_02_02_01_two_wheeler	sub3sectors			Two-wheeler (freight) 1
# 15_02_02_01_two_wheeler_freight	sub3sectors			Two-wheeler (freight) 2
# 01_coal_unallocated	subfuels			Unallocated Coal
# 06_crude_oil_and_ngl_unallocated	subfuels			Unallocated Crude oil and NGL
# 08_gas_unallocated	subfuels			Unallocated Gas
# 16_others_unallocated	subfuels			Unallocated others
# 07_petroleum_products_unallocated	subfuels			Unallocated Petroleum Products
# 12_solar_unallocated	subfuels			Unallocated Solar
# 15_solid_biomass_unallocated	subfuels			Unallocated solid biomass
# 09_01_08_01_utility	sub3sectors			Utility-scale solar
# 18_01_08_01_utility	sub3sectors			Utility-scale solar (electricity output)
# 		07.12 White spirit SBP	products	White spirit SBP
# 14_wind	fuels	14 Wind	products	Wind
# 09_01_09_wind	sub2sectors			Wind power
# 18_01_09_wind	sub2sectors			Wind power (electricity output)
# 14_03_09_wood_and_wood_products	sub2sectors	14.03.09 Wood and wood products	flows	Wood and wood products
