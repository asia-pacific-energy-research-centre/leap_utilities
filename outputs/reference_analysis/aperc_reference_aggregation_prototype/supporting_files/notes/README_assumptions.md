# APERC Aggregated Reference Prototype - Assumptions

- 7th uses BAU as reference-equivalent.
- 7th alternating Mtoe/PJ structure is parsed by selecting adjacent PJ columns.
- 8th uses Demand - Reference and Final consumption by fuel (PJ).
- 9th workbook contributes total via Reference TFC/TFC in the TFEC note block.
- 9th workbook does not include a fully comparable fuel table for all buckets.
- Optional 9th fuel source used: Yes.
- Fuel harmonization bucket: Renewables_total
  - 7th: Renewables
  - 8th: Biomass + Other renewables
- 9th optional: Renewables_total directly, or Biomass + Other renewables
- Aggregated series are blank before 2022, using the configured start year from the 9th series.
- From 2022 onward, aggregated series are shifted to match the 9th value at the start year while preserving the aggregated trend.
- Aggregation method: weighted average using 33/66/99 for 7th/8th/9th.
- 7th data are interpolated to annual values before aggregation.
- 7th/8th series are extended from 2051 to 2060 using each series historical trend.
- Missing years are not imputed in output tables; charts connect gaps visually.
