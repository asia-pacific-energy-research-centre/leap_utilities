param(
  [Parameter(Mandatory=$true)][string[]]$AreaNames,
  [int]$ScenarioMax = 200,
  [int]$RegionMax = 100,
  [int]$FavoriteMax = 1000,
  [int]$VersionMax = 500,
  [int]$UserVarMax = 4000,
  [int]$TimeSliceMax = 1000,
  [int]$YearlyShapeMax = 2000,
  [int]$BranchScanMax = 6000
)

throw "LEAP API usage is disabled in this repository due a known LEAP API bug. This script intentionally blocks COM calls."

function Scan-ByIndex($getter, $maxIndex) {
  $rows = @()
  for($i=1; $i -le $maxIndex; $i++) {
    try {
      $obj = & $getter $i
      $name = ""
      $id = $null
      try { $name = [string]$obj.Name } catch {}
      try { $id = [int]$obj.ID } catch {}
      $rows += [PSCustomObject]@{idx=$i; id=$id; name=$name}
    } catch {
      $rows += [PSCustomObject]@{idx=$i; id=$null; name=""}
    }
  }
  return $rows
}

$o = New-Object -ComObject LEAP.LEAPApplication
$o.Visible = $false

$out = @()
foreach($area in $AreaNames) {
  $o.ActiveArea = $area

  $sc = Scan-ByIndex {param($i) $o.Scenarios($i)} $ScenarioMax
  $rg = Scan-ByIndex {param($i) $o.Regions($i)} $RegionMax
  $fv = Scan-ByIndex {param($i) $o.Favorites($i)} $FavoriteMax
  $vr = Scan-ByIndex {param($i) $o.Versions($i)} $VersionMax
  $uv = Scan-ByIndex {param($i) $o.UserVariables($i)} $UserVarMax
  $ts = Scan-ByIndex {param($i) $o.TimeSlices($i)} $TimeSliceMax
  $ys = Scan-ByIndex {param($i) $o.YearlyShapes($i)} $YearlyShapeMax
  $br = Scan-ByIndex {param($i) $o.Branches($i)} $BranchScanMax

  $scenarioNames = @($sc | Where-Object {$_.name -ne ''} | Select-Object -ExpandProperty name -Unique)
  $regionNames = @($rg | Where-Object {$_.name -ne ''} | Select-Object -ExpandProperty name -Unique)
  $favoriteNames = @($fv | Where-Object {$_.name -ne ''} | Select-Object -ExpandProperty name -Unique)
  $versionNames = @($vr | Where-Object {$_.name -ne ''} | Select-Object -ExpandProperty name -Unique)
  $uvNames = @($uv | Where-Object {$_.name -ne ''} | Select-Object -ExpandProperty name -Unique)
  $tsNames = @($ts | Where-Object {$_.name -ne ''} | Select-Object -ExpandProperty name -Unique)
  $ysNames = @($ys | Where-Object {$_.name -ne ''} | Select-Object -ExpandProperty name -Unique)

  $branchIds = @($br | Where-Object { $_.id -ne $null -and $_.id -gt 0 } | Select-Object -ExpandProperty id -Unique)
  $branchRootHits = @($br | Where-Object { $_.id -eq -1 }).Count

  $out += [PSCustomObject]@{
    area = $area
    base_year = $o.BaseYear
    end_year = $o.EndYear
    first_scenario_year = $o.FirstScenarioYear
    results_every = $o.ResultsEvery
    all_results_saved = $o.AllResultsSaved
    scenarios_count = $scenarioNames.Count
    scenarios = $scenarioNames
    regions_count = $regionNames.Count
    regions = $regionNames
    favorites_count = $favoriteNames.Count
    versions_count = $versionNames.Count
    user_variables_count = $uvNames.Count
    timeslices_count = $tsNames.Count
    yearlyshapes_count = $ysNames.Count
    branch_count_est_unique_ids = $branchIds.Count
    branch_scan_root_hits = $branchRootHits
  }
}

$out | ConvertTo-Json -Depth 8
