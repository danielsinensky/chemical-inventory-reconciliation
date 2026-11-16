// Power Query (M) custom columns used in inventory_reconciliation.xlsm

// ---- People query: flag contacts not confirmed in over a year (cutoff = data date minus 365 days) ----
Stale =
    [last_confirmed] < #date(2025, 10, 1)

// ---- System query: standardize units (added as UnitClean because Power Query treats Unit and unit as the same name) ----
UnitClean =
    if List.Contains({"ml", "milliliters"}, Text.Lower([unit])) then "mL"
    else if List.Contains({"g", "grams"}, Text.Lower([unit])) then "g"
    else [unit]

// ---- System query: final chemical name. CAS lookup first, then the alias table, then the typed name ----
ChemicalStd =
    if [CasName] <> null then [CasName]
    else if [AliasName] <> null then [AliasName]
    else [chemical_name]

// ---- Reconciliation query: Full Outer merge of SystemVisited and Field on container_id, then five custom columns ----

// Status =
    if [container_id] = null then "Not in system"
    else if [Field.container_id] = null then "Missing on shelf"
    else if [room_id] <> [Field.room_id] then "Room mismatch"
    else if [quantity] <> [Field.quantity] then "Quantity mismatch"
    else "Match"

// LabID =
    if [lab_id] = null then [Field.lab_id] else [lab_id]

// ContainerID =
    if [container_id] = null then [Field.container_id] else [container_id]

// Chemical =
    if [ChemicalStd] = null then [Field.chemical_name] else [ChemicalStd]

// Detail =
    if [Status] = "Room mismatch" then "Listed in " & [room_id] & ", found in " & [Field.room_id]
    else if [Status] = "Quantity mismatch" then "Listed " & Text.From([quantity]) & ", found " & Text.From([Field.quantity])
    else if [Status] = "Missing on shelf" then "Listed in " & [room_id] & ", not found"
    else if [Status] = "Not in system" then "Found in " & [Field.room_id] & ", not in system"
    else ""
