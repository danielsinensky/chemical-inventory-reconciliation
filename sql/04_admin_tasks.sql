-- Administration practice: account creation, value update, deactivation, saved view.
-- SQLite has no built-in logins, so app_users models how an application handles roles.

-- Account creation request
INSERT INTO app_users (person_id, role) VALUES ('P012', 'Lab Manager');
INSERT INTO audit_log (action, detail, done_by) VALUES ('create_account', 'P012 as Lab Manager', 'daniel');

-- Value update: a container moved rooms
UPDATE inventory SET room_id = 'R002' WHERE container_id = 'C00003';
INSERT INTO audit_log (action, detail, done_by) VALUES ('update_room', 'C00003 moved to R002', 'daniel');

-- Deactivate accounts of people not confirmed in over a year (least privilege)
UPDATE app_users SET is_active = 0
WHERE person_id IN (SELECT person_id FROM people WHERE last_confirmed < date('2026-10-01', '-365 days'));

-- Status report as a saved view
CREATE VIEW v_lab_inventory_summary AS
SELECT l.lab_id, l.building, l.pi_name, COUNT(i.container_id) AS containers,
  SUM(CASE WHEN c.hazard_class = 'Flammable' THEN 1 ELSE 0 END) AS flammables,
  SUM(CASE WHEN i.cas_number IS NULL THEN 1 ELSE 0 END) AS missing_cas
FROM labs l
LEFT JOIN inventory i ON i.lab_id = l.lab_id
LEFT JOIN chemicals c ON c.cas_number = i.cas_number
GROUP BY l.lab_id, l.building, l.pi_name;
