-- Clean staging_inventory into inventory: remove duplicates, blank empty CAS numbers,
-- standardize units, and convert MM/DD/YYYY dates to YYYY-MM-DD. Expected result: 564 rows.

INSERT OR IGNORE INTO inventory (container_id, lab_id, room_id, cas_number, quantity, unit, last_verified)
SELECT container_id, lab_id, room_id, NULLIF(cas_number, ''), CAST(quantity AS REAL),
  CASE WHEN lower(unit) IN ('ml', 'milliliters') THEN 'mL'
       WHEN lower(unit) IN ('g', 'grams') THEN 'g' ELSE unit END,
  CASE WHEN last_verified LIKE '__/__/____'
       THEN substr(last_verified, 7, 4) || '-' || substr(last_verified, 1, 2) || '-' || substr(last_verified, 4, 2)
       ELSE last_verified END
FROM staging_inventory;
