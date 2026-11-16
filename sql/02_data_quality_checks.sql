-- Data-quality checks on staging_inventory. Run one at a time.
-- Expected: 16 duplicate container IDs, 12 missing CAS numbers, 7 unit spellings of 2 units,
-- 32 names that disagree with the CAS reference list, and stale contacts only in L04 and L11 (3 each).

-- Duplicate container records
SELECT container_id, COUNT(*) AS copies FROM staging_inventory
GROUP BY container_id HAVING COUNT(*) > 1;

-- Missing CAS numbers (you should see 12)
SELECT * FROM staging_inventory WHERE cas_number IS NULL OR cas_number = '';

-- Unit spellings (you should see 7 variants of 2 units)
SELECT unit, COUNT(*) AS n FROM staging_inventory GROUP BY unit ORDER BY n DESC;

-- Names that disagree with the CAS reference list
SELECT s.container_id, s.chemical_name AS typed_name, c.chemical_name AS standard_name
FROM staging_inventory s JOIN chemicals c ON c.cas_number = s.cas_number
WHERE s.chemical_name <> c.chemical_name;

-- Contacts not confirmed in over a year, by lab
SELECT lab_id, COUNT(*) AS stale_contacts FROM people
WHERE last_confirmed < date('2026-10-01', '-365 days')
GROUP BY lab_id ORDER BY stale_contacts DESC;
